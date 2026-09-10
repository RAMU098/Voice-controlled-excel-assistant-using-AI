# 🎤 Voice-Controlled Excel Assistant

> **AI-Powered Voice Interface for Microsoft Excel Operations**

A modern web application that allows users to interact with Microsoft Excel using natural voice commands. Built with Python, Flask, and advanced NLP capabilities.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Flask](https://img.shields.io/badge/Flask-3.0-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## ✨ Features

### 🎙️ Voice Control
- **Real-time speech recognition** using Web Speech API
- Hands-free Excel operation
- Space bar shortcut for quick voice input
- Voice feedback for command confirmation

### 📊 Excel Operations
- **Data Entry**: Write, read, and clear cells
- **Formatting**: Bold, italic, colors, currency, borders
- **Formulas**: SUM, AVERAGE, MAX, MIN, COUNT, IF, VLOOKUP
- **Charts**: Bar, Pie, Line charts
- **Data Operations**: Sort, filter, auto-filter
- **Sheet Management**: Create, delete, rename, switch sheets

### 🤖 AI-Powered NLP
- Natural language understanding
- Rule-based command parsing
- Optional OpenAI GPT integration for enhanced understanding
- Intelligent command interpretation

### 🎨 Modern UI
- Beautiful dark theme interface
- Real-time data preview
- Drag-and-drop file upload
- Command history tracking
- Responsive design

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)
- Modern web browser (Chrome or Edge recommended for voice features)
- Microphone for voice input

### Installation

1. **Clone or navigate to the project directory:**
   ```bash
   cd /path/to/voice-excel
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment (optional for enhanced AI):**
   Create a `.env` file:
   ```
   OPENAI_API_KEY=your_openai_api_key_here
   SECRET_KEY=your-secret-key
   DEBUG=True
   ```

5. **Run the application:**
   ```bash
   python app.py
   ```

6. **Open your browser:**
   Navigate to `http://localhost:5000`

## 📖 Usage Guide

### Voice Commands

#### Data Entry
- "Write 'Hello' to cell A1"
- "Set cell B2 to 100"
- "Read cell A1"
- "Clear cell B2"
- "Clear cells A1 to B10"

#### Formatting
- "Make cell A1 bold"
- "Format B2 as italic"
- "Color cell C3 red"
- "Highlight A1 yellow"
- "Format D4 as currency"
- "Add border to A1:D10"
- "Merge cells A1 to C1"

#### Rows & Columns
- "Insert row at row 5"
- "Delete row 3"
- "Insert column at column B"
- "Delete column C"

#### Formulas
- "Sum A1 to A10 in cell A11"
- "Average B1 through B10"
- "Find the maximum in C1 to C20"
- "Find the minimum in D1 to D10"
- "Insert formula =B1*C1 in cell D1"

#### Charts
- "Create a bar chart for A1:B10"
- "Create a pie chart for sales data"
- "Create a line chart"

#### Data Operations
- "Sort by column A ascending"
- "Sort data descending"
- "Filter column B where value is Sales"

#### Sheets
- "Create new sheet named 'Report'"
- "Delete sheet 'Old Data'"
- "Switch to sheet 'Summary'"
- "Rename sheet 'Sheet1' to 'Data'"
- "List all sheets"

#### Sample Data
- "Create sample sales data"
- "Add sample employee data"
- "Generate inventory data"

#### File Operations
- "Save the workbook"
- "Create new workbook"

## 🏗️ Project Structure

```
voice-excel/
├── app.py                 # Main Flask application
├── config.py              # Configuration settings
├── excel_automation.py    # Excel operations module
├── nlp_engine.py          # NLP command interpreter
├── speech_handler.py      # Text-to-speech module
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── templates/
│   └── index.html         # Main HTML template
├── static/
│   ├── css/
│   │   └── style.css      # Stylesheet
│   └── js/
│       └── app.js         # Frontend JavaScript
├── uploads/               # Uploaded Excel files
└── output/                # Generated Excel files
```

## 🔧 Technology Stack

| Component | Technology |
|-----------|------------|
| Backend | Python, Flask |
| Excel Automation | OpenPyXL, Pandas |
| NLP | Rule-based + OpenAI GPT (optional) |
| Speech Recognition | Web Speech API |
| Text-to-Speech | gTTS |
| Frontend | HTML5, CSS3, JavaScript |
| UI Design | Modern Dark Theme |

## 📋 Requirements

### Software Requirements
- Operating System: Windows 10+, macOS, or Linux
- Python 3.8+
- Modern web browser (Chrome/Edge recommended)
- Microphone for voice input

### Hardware Requirements
- Processor: Intel i5 or equivalent
- RAM: 8 GB minimum (16 GB recommended)
- Storage: 500 MB free space
- Internet: Required for speech recognition

## 🔐 API Configuration

### OpenAI API (Optional)
For enhanced natural language understanding:
1. Get an API key from [OpenAI](https://platform.openai.com/api-keys)
2. Set the `OPENAI_API_KEY` environment variable

The application works without OpenAI using rule-based NLP, but AI integration provides better understanding of complex or ambiguous commands.

## 🐛 Troubleshooting

### Voice Recognition Not Working
- Ensure you're using Chrome or Edge browser
- Allow microphone permissions when prompted
- Check that your microphone is working properly

### Commands Not Recognized
- Speak clearly and at a normal pace
- Try the text input as an alternative
- Check the help modal for correct command formats

### Excel File Not Downloading
- Check browser download settings
- Ensure the workbook has data to download

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## 📄 License

This project is licensed under the MIT License.

## 🙏 Acknowledgments

- OpenAI for GPT API
- Google for Web Speech API
- OpenPyXL team for Excel automation library
- All contributors and testers

---

**Made with ❤️ for the School of Computing - CSE**

*Voice-Controlled Excel Assistant - Making Excel Accessible to Everyone*

