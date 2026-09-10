"""
Voice-Controlled Excel Assistant
Main Flask Application
"""
import os
import json
import base64
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, session
from flask_cors import CORS
from werkzeug.utils import secure_filename

from config import Config
from excel_automation import ExcelAutomation
from nlp_engine import NLPEngine
from speech_handler import get_speech_handler

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = Config.SECRET_KEY
app.config['UPLOAD_FOLDER'] = Config.UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = Config.MAX_CONTENT_LENGTH
CORS(app)

# Initialize folders
Config.init_folders()

# Store active Excel instances (in production, use Redis or similar)
excel_instances = {}
nlp_engine = NLPEngine()
speech_handler = get_speech_handler()


def get_excel_instance(session_id: str) -> ExcelAutomation:
    """Get or create Excel instance for a session."""
    if session_id not in excel_instances:
        excel_instances[session_id] = ExcelAutomation()
    return excel_instances[session_id]


def allowed_file(filename: str) -> bool:
    """Check if file extension is allowed."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Render the main application page."""
    return render_template('index.html')


@app.route('/api/session', methods=['POST'])
def create_session():
    """Create a new session and return session ID."""
    session_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    excel_instances[session_id] = ExcelAutomation()
    
    return jsonify({
        'success': True,
        'session_id': session_id,
        'message': 'Session created successfully'
    })


@app.route('/api/command', methods=['POST'])
def process_command():
    """Process a voice/text command."""
    try:
        data = request.get_json()
        command = data.get('command', '').strip()
        session_id = data.get('session_id', 'default')
        
        if not command:
            return jsonify({
                'success': False,
                'message': 'No command provided'
            })
        
        # Get Excel instance
        excel = get_excel_instance(session_id)
        
        # Interpret the command using NLP
        interpretation = nlp_engine.interpret_command(command)
        action = interpretation.get('action')
        params = interpretation.get('params', {})
        
        # Execute the action
        result = execute_action(excel, action, params)
        
        # Generate audio feedback if requested
        audio_data = None
        if data.get('voice_feedback', False):
            audio_data = speech_handler.text_to_speech_base64(result['message'])
        
        # Get updated data preview
        data_preview = excel.get_data_preview(10)
        sheet_info = excel.get_sheet_info()
        
        # Clean data preview to remove NaN values (not JSON serializable)
        import math
        def clean_value(val):
            if val is None:
                return None
            if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                return None
            return val
        
        clean_preview = [[clean_value(cell) for cell in row] for row in data_preview]
        
        return jsonify({
            'success': result['success'],
            'message': result['message'],
            'action': action,
            'params': params,
            'data_preview': clean_preview,
            'sheet_info': sheet_info,
            'audio': audio_data,
            'file_path': excel.get_file_path()
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error processing command: {str(e)}'
        })


def execute_action(excel: ExcelAutomation, action: str, params: dict) -> dict:
    """Execute an Excel action based on interpreted command."""
    try:
        message = ""
        success = True
        
        if action == 'write_cell':
            cell = params.get('cell', 'A1')
            value = params.get('value', '')
            # Try to convert to number if possible
            try:
                if '.' in str(value):
                    value = float(value)
                else:
                    value = int(value)
            except (ValueError, TypeError):
                pass
            message = excel.write_cell(cell, value)
        
        elif action == 'read_cell':
            cell = params.get('cell', 'A1')
            message = excel.read_cell(cell)
        
        elif action == 'clear_cell':
            if 'start_cell' in params:
                message = excel.clear_range(params['start_cell'], params['end_cell'])
            else:
                message = excel.clear_cell(params.get('cell', 'A1'))
        
        elif action == 'format_bold':
            if 'start_cell' in params:
                message = excel.format_range(params['start_cell'], params['end_cell'], bold=True)
            else:
                message = excel.format_cell(params.get('cell', 'A1'), bold=True)
        
        elif action == 'format_italic':
            message = excel.format_cell(params.get('cell', 'A1'), italic=True)
        
        elif action == 'format_color':
            cell = params.get('cell', 'A1')
            color = params.get('bg_color', 'FFFF00')
            message = excel.format_cell(cell, bg_color=color)
        
        elif action == 'format_currency':
            if 'start_cell' in params:
                for row in excel.active_sheet[params['start_cell']:params['end_cell']]:
                    for cell in row:
                        excel.format_as_currency(cell.coordinate)
                message = f"Formatted range as currency"
            else:
                message = excel.format_as_currency(params.get('cell', 'A1'))
        
        elif action == 'insert_row':
            row = params.get('row', 1)
            count = params.get('count', 1)
            message = excel.insert_row(row, count)
        
        elif action == 'delete_row':
            row = params.get('row', 1)
            count = params.get('count', 1)
            message = excel.delete_row(row, count)
        
        elif action == 'insert_column':
            column = params.get('column', 'A')
            count = params.get('count', 1)
            message = excel.insert_column(column, count)
        
        elif action == 'delete_column':
            column = params.get('column', 'A')
            count = params.get('count', 1)
            message = excel.delete_column(column, count)
        
        elif action == 'sum':
            start = params.get('start', 'A1')
            end = params.get('end', 'A10')
            result_cell = params.get('result', params.get('result_cell'))
            
            # Auto-determine result cell if not provided
            if not result_cell:
                import re
                end_row = int(re.findall(r'\d+', end)[0])
                col = re.findall(r'[A-Za-z]+', end)[0]
                result_cell = f"{col}{end_row + 1}"
            
            message = excel.sum_range(result_cell, start, end)
        
        elif action == 'sum_row':
            row = params.get('row', 1)
            message = excel.sum_row(row)
        
        elif action == 'sum_column':
            column = params.get('column', 'A')
            message = excel.sum_column(column)
        
        elif action == 'average_row':
            row = params.get('row', 1)
            message = excel.average_row(row)
        
        elif action == 'average_column':
            column = params.get('column', 'A')
            message = excel.average_column(column)
        
        elif action == 'format_column':
            column = params.get('column', 'A')
            bg_color = params.get('bg_color')
            bold = params.get('bold', False)
            if bg_color:
                message = excel.format_column(column, bg_color=bg_color)
            elif bold:
                message = excel.format_column(column, bold=True)
            else:
                message = excel.format_column(column)
        
        elif action == 'average':
            start = params.get('start', 'A1')
            end = params.get('end', 'A10')
            result_cell = params.get('result', params.get('result_cell'))
            
            if not result_cell:
                import re
                end_row = int(re.findall(r'\d+', end)[0])
                col = re.findall(r'[A-Za-z]+', end)[0]
                result_cell = f"{col}{end_row + 1}"
            
            message = excel.average_range(result_cell, start, end)
        
        elif action == 'max':
            start = params.get('start', 'A1')
            end = params.get('end', 'A10')
            result_cell = params.get('result', params.get('result_cell'))
            
            if not result_cell:
                import re
                end_row = int(re.findall(r'\d+', end)[0])
                col = re.findall(r'[A-Za-z]+', end)[0]
                result_cell = f"{col}{end_row + 1}"
            
            message = excel.max_range(result_cell, start, end)
        
        elif action == 'min':
            start = params.get('start', 'A1')
            end = params.get('end', 'A10')
            result_cell = params.get('result', params.get('result_cell'))
            
            if not result_cell:
                import re
                end_row = int(re.findall(r'\d+', end)[0])
                col = re.findall(r'[A-Za-z]+', end)[0]
                result_cell = f"{col}{end_row + 1}"
            
            message = excel.min_range(result_cell, start, end)
        
        elif action == 'formula':
            cell = params.get('cell', 'A1')
            formula = params.get('formula', '')
            message = excel.insert_formula(cell, formula)
        
        elif action == 'chart_bar':
            data_range = f"{params.get('start', 'A1')}:{params.get('end', 'B10')}"
            if 'data_range' in params:
                data_range = params['data_range']
            title = params.get('title', 'Bar Chart')
            message = excel.create_bar_chart(data_range, title)
        
        elif action == 'chart_pie':
            data_range = f"{params.get('start', 'A1')}:{params.get('end', 'B10')}"
            if 'data_range' in params:
                data_range = params['data_range']
            title = params.get('title', 'Pie Chart')
            message = excel.create_pie_chart(data_range, title)
        
        elif action == 'chart_line':
            data_range = f"{params.get('start', 'A1')}:{params.get('end', 'B10')}"
            if 'data_range' in params:
                data_range = params['data_range']
            title = params.get('title', 'Line Chart')
            message = excel.create_line_chart(data_range, title)
        
        elif action == 'sort':
            column = params.get('column', 'A')
            ascending = params.get('ascending', True)
            message = excel.sort_column(column, ascending)
        
        elif action == 'filter':
            column = params.get('column', 'A')
            value = params.get('value', '')
            message = excel.filter_data(column, value)
        
        elif action == 'create_sheet':
            name = params.get('name', 'NewSheet')
            message = excel.create_sheet(name)
        
        elif action == 'delete_sheet':
            name = params.get('name', '')
            message = excel.delete_sheet(name)
        
        elif action == 'rename_sheet':
            old_name = params.get('old_name', '')
            new_name = params.get('new_name', '')
            message = excel.rename_sheet(old_name, new_name)
        
        elif action == 'switch_sheet':
            name = params.get('name', '')
            message = excel.switch_sheet(name)
        
        elif action == 'list_sheets':
            message = excel.list_sheets()
        
        elif action == 'merge_cells':
            start = params.get('start', 'A1')
            end = params.get('end', 'A2')
            message = excel.merge_cells(start, end)
        
        elif action == 'border':
            if 'start' in params:
                message = excel.add_border_range(params['start'], params['end'])
            else:
                message = excel.add_border(params.get('cell', 'A1'))
        
        elif action == 'sample_data':
            data_type = params.get('type', 'sales')
            message = excel.create_sample_data(data_type)
        
        elif action == 'save':
            filename = params.get('filename')
            if filename:
                save_path = os.path.join(Config.OUTPUT_FOLDER, secure_filename(filename))
                if not save_path.endswith('.xlsx'):
                    save_path += '.xlsx'
                message = excel.save_workbook(save_path)
            else:
                message = excel.save_workbook()
        
        elif action == 'new_workbook':
            message = excel.create_new_workbook()
        
        elif action == 'help':
            message = nlp_engine.get_help_text()
        
        elif action == 'copy_cell':
            source = params.get('source', 'A1')
            destination = params.get('destination', 'B1')
            message = excel.copy_cell(source, destination)
        
        elif action == 'find':
            value = params.get('value', '')
            message = excel.find_value(value)
        
        elif action == 'replace':
            search = params.get('search', '')
            replace = params.get('replace', '')
            message = excel.replace_value(search, replace)
        
        elif action == 'count_column':
            column = params.get('column', 'A')
            message = excel.count_column(column)
        
        elif action == 'max_column':
            column = params.get('column', 'A')
            message = excel.max_column(column)
        
        elif action == 'min_column':
            column = params.get('column', 'A')
            message = excel.min_column(column)
        
        elif action == 'auto_fit':
            message = excel.auto_fit_columns()
        
        elif action == 'clear_sheet':
            message = excel.clear_sheet()
        
        elif action == 'freeze':
            row = params.get('row', 2)
            cell = f"A{row}" if row else "A2"
            message = excel.freeze_panes(cell)
        
        elif action == 'unfreeze':
            message = excel.unfreeze_panes()
        
        elif action == 'unknown':
            message = params.get('message', "I couldn't understand that command. Say 'help' for available commands.")
            success = False
        
        else:
            message = f"Action '{action}' is not yet implemented."
            success = False
        
        return {'success': success, 'message': message}
    
    except Exception as e:
        return {'success': False, 'message': f'Error executing action: {str(e)}'}


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload."""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No file provided'})
        
        file = request.files['file']
        session_id = request.form.get('session_id', 'default')
        
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No file selected'})
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(Config.UPLOAD_FOLDER, filename)
            file.save(filepath)
            
            # Load the file into Excel instance
            excel = get_excel_instance(session_id)
            result = excel.load_workbook(filepath)
            
            return jsonify({
                'success': True,
                'message': result,
                'data_preview': excel.get_data_preview(10),
                'sheet_info': excel.get_sheet_info()
            })
        
        return jsonify({'success': False, 'message': 'Invalid file type'})
    
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/download', methods=['GET'])
def download_file():
    """Download the current workbook."""
    try:
        session_id = request.args.get('session_id', 'default')
        excel = get_excel_instance(session_id)
        
        # Save the current workbook
        excel.save_workbook()
        file_path = excel.get_file_path()
        
        if file_path and os.path.exists(file_path):
            return send_file(
                file_path,
                as_attachment=True,
                download_name=os.path.basename(file_path)
            )
        
        return jsonify({'success': False, 'message': 'No file to download'})
    
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/preview', methods=['GET'])
def get_preview():
    """Get current data preview."""
    try:
        session_id = request.args.get('session_id', 'default')
        excel = get_excel_instance(session_id)
        
        return jsonify({
            'success': True,
            'data_preview': excel.get_data_preview(20),
            'sheet_info': excel.get_sheet_info()
        })
    
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/speak', methods=['POST'])
def text_to_speech():
    """Convert text to speech."""
    try:
        data = request.get_json()
        text = data.get('text', '')
        
        if not text:
            return jsonify({'success': False, 'message': 'No text provided'})
        
        audio_base64 = speech_handler.text_to_speech_base64(text)
        
        if audio_base64:
            return jsonify({
                'success': True,
                'audio': audio_base64
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Text-to-speech not available'
            })
    
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        'status': 'healthy',
        'message': 'Voice-Controlled Excel Assistant is running'
    })


if __name__ == '__main__':
    print("\n" + "="*60)
    print("🎤 Voice-Controlled Excel Assistant")
    print("="*60)
    print("Starting server on http://localhost:5000")
    print("Open your browser and navigate to the URL above")
    print("="*60 + "\n")
    
    app.run(debug=Config.DEBUG, host='0.0.0.0', port=5000)

