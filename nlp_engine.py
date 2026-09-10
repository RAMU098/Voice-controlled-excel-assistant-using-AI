"""
NLP Engine for Voice-Controlled Excel Assistant
Interprets natural language commands and maps them to Excel operations
"""
import re
import os
from typing import Dict, Any, Optional, Tuple, List
from config import Config


class NLPEngine:
    """
    Natural Language Processing Engine that interprets voice commands
    and maps them to Excel operations.
    """
    
    def __init__(self):
        """Initialize the NLP Engine."""
        self.openai_client = None
        self._init_openai()
        
        # Command patterns for rule-based fallback
        self.command_patterns = self._build_command_patterns()
        
        # Number word mappings
        self.number_words = {
            'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4,
            'five': 5, 'six': 6, 'seven': 7, 'eight': 8, 'nine': 9,
            'ten': 10, 'eleven': 11, 'twelve': 12, 'thirteen': 13,
            'fourteen': 14, 'fifteen': 15, 'sixteen': 16, 'seventeen': 17,
            'eighteen': 18, 'nineteen': 19, 'twenty': 20, 'first': 1,
            'second': 2, 'third': 3, 'fourth': 4, 'fifth': 5
        }
        
        # Color mappings
        self.color_map = {
            'red': 'FF0000', 'green': '00FF00', 'blue': '0000FF',
            'yellow': 'FFFF00', 'orange': 'FFA500', 'purple': '800080',
            'pink': 'FFC0CB', 'black': '000000', 'white': 'FFFFFF',
            'gray': '808080', 'grey': '808080', 'cyan': '00FFFF',
            'magenta': 'FF00FF', 'light blue': 'ADD8E6', 'dark blue': '00008B',
            'light green': '90EE90', 'dark green': '006400'
        }
    
    def _init_openai(self):
        """Initialize OpenAI client if API key is available."""
        if Config.OPENAI_API_KEY:
            try:
                from openai import OpenAI
                self.openai_client = OpenAI(api_key=Config.OPENAI_API_KEY)
            except Exception as e:
                print(f"OpenAI initialization failed: {e}")
                self.openai_client = None
    
    def _build_command_patterns(self) -> Dict[str, List[Dict]]:
        """Build regex patterns for command recognition."""
        return {
            'write_cell': [
                {'pattern': r'(?:write|put|enter|type|add|set)\s+["\']?(.+?)["\']?\s+(?:in|to|at)\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['value', 'cell']},
                {'pattern': r'(?:cell\s+)?([A-Za-z]+\d+)\s+(?:should be|equals?|is)\s+["\']?(.+?)["\']?$', 'groups': ['cell', 'value']},
                {'pattern': r'(?:set|make)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:to|as|equal to)\s+["\']?(.+?)["\']?$', 'groups': ['cell', 'value']},
            ],
            'read_cell': [
                {'pattern': r'(?:read|get|show|what(?:\'s| is))\s+(?:in\s+)?(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
                {'pattern': r'(?:what(?:\'s| is))\s+(?:the\s+)?(?:value\s+)?(?:in|of|at)\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
            ],
            'clear_cell': [
                {'pattern': r'(?:clear|delete|remove|erase)\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
                {'pattern': r'(?:clear|delete|remove|erase)\s+(?:cells?\s+)?(?:from\s+)?([A-Za-z]+\d+)\s+(?:to|through)\s+([A-Za-z]+\d+)', 'groups': ['start_cell', 'end_cell']},
            ],
            'format_bold': [
                {'pattern': r'(?:make|format|set)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:as\s+)?bold', 'groups': ['cell']},
                {'pattern': r'bold\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
                {'pattern': r'(?:make|format|set)\s+(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)\s+(?:as\s+)?bold', 'groups': ['start_cell', 'end_cell']},
            ],
            'format_italic': [
                {'pattern': r'(?:make|format|set)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:as\s+)?italic', 'groups': ['cell']},
                {'pattern': r'italic\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
            ],
            'format_color': [
                {'pattern': r'(?:make|set|change)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:background\s+)?(?:color\s+)?(?:to\s+)?(\w+(?:\s+\w+)?)\s*(?:color)?', 'groups': ['cell', 'color']},
                {'pattern': r'(?:color|highlight)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:with\s+)?(\w+(?:\s+\w+)?)', 'groups': ['cell', 'color']},
                {'pattern': r'(?:highlight|color)\s+(?:the\s+)?(?:top\s+)?row\s+(?:with\s+)?(\w+(?:\s+\w+)?)', 'groups': ['color']},
            ],
            'format_column': [
                # "color column D red", "make column B yellow", etc.
                {'pattern': r'(?:color|colour|highlight)\s+column\s+([A-Za-z]+)\s+(?:with\s+)?(\w+)', 'groups': ['column', 'color']},
                {'pattern': r'(?:make|set|change)\s+column\s+([A-Za-z]+)\s+(?:color\s+)?(?:to\s+)?(\w+)', 'groups': ['column', 'color']},
                {'pattern': r'column\s+([A-Za-z]+)\s+(?:color|colour)\s+(\w+)', 'groups': ['column', 'color']},
                {'pattern': r'(?:do\s+)?column\s+([A-Za-z]+)\s+(\w+)\s+(?:color|colour)', 'groups': ['column', 'color']},
                {'pattern': r'(\w+)\s+(?:color|colour)\s+(?:for\s+)?column\s+([A-Za-z]+)', 'groups': ['color', 'column']},
            ],
            'format_currency': [
                {'pattern': r'(?:format|make|set)\s+(?:cell\s+)?([A-Za-z]+\d+)\s+(?:as\s+)?currency', 'groups': ['cell']},
                {'pattern': r'(?:format|make|set)\s+(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)\s+(?:as\s+)?currency', 'groups': ['start_cell', 'end_cell']},
            ],
            'insert_row': [
                {'pattern': r'(?:insert|add)\s+(?:a\s+)?(?:new\s+)?row(?:\s+at)?\s*(?:row\s+)?(\d+|first|second|third|fourth|fifth|\w+)', 'groups': ['row']},
                {'pattern': r'(?:insert|add)\s+(\d+|one|two|three|four|five|\w+)\s+rows?\s+(?:at\s+)?(?:row\s+)?(\d+)?', 'groups': ['count', 'row']},
            ],
            'delete_row': [
                {'pattern': r'(?:delete|remove)\s+row\s*(\d+)', 'groups': ['row']},
                {'pattern': r'(?:delete|remove)\s+(?:the\s+)?(\d+|first|second|third|fourth|fifth|\w+)\s+row', 'groups': ['row']},
            ],
            'insert_column': [
                {'pattern': r'(?:insert|add)\s+(?:a\s+)?(?:new\s+)?column(?:\s+at)?\s*(?:column\s+)?([A-Za-z])', 'groups': ['column']},
                {'pattern': r'(?:insert|add)\s+(\d+|one|two|three|\w+)\s+columns?\s+(?:at\s+)?(?:column\s+)?([A-Za-z])?', 'groups': ['count', 'column']},
            ],
            'delete_column': [
                {'pattern': r'(?:delete|remove)\s+column\s*([A-Za-z])', 'groups': ['column']},
            ],
            'sum': [
                # Row-based sum patterns (check these FIRST - more specific)
                {'pattern': r'(?:do\s+)?(?:the\s+)?sum\s+(?:of\s+)?row\s+(\d+)', 'groups': ['row'], 'row_sum': True},
                {'pattern': r'(?:sum|total|add)\s+(?:up\s+)?row\s+(\d+)', 'groups': ['row'], 'row_sum': True},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?sum\s+(?:of\s+)?row\s+(\d+)', 'groups': ['row'], 'row_sum': True},
                {'pattern': r'row\s+(\d+)\s+(?:sum|total)', 'groups': ['row'], 'row_sum': True},
                # Column-based sum patterns (BEFORE cell range patterns)
                {'pattern': r'(?:do\s+)?(?:the\s+)?sum\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_sum': True},
                {'pattern': r'(?:sum|total|add)\s+(?:up\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_sum': True},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?sum\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_sum': True},
                {'pattern': r'column\s+([A-Za-z]+)\s+(?:sum|total)', 'groups': ['column'], 'column_sum': True},
                # Cell range patterns
                {'pattern': r'(?:sum|add up|total)\s+(?:cells?\s+)?(?:from\s+)?([A-Za-z]+\d+)\s+(?:to|through)\s+([A-Za-z]+\d+)(?:\s+(?:in|at|to)\s+(?:cell\s+)?([A-Za-z]+\d+))?', 'groups': ['start', 'end', 'result']},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?sum\s+(?:of\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'average': [
                # Row-based average patterns (FIRST - more specific)
                {'pattern': r'(?:do\s+)?(?:the\s+)?average\s+(?:of\s+)?row\s+(\d+)', 'groups': ['row'], 'row_avg': True},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?average\s+(?:of\s+)?row\s+(\d+)', 'groups': ['row'], 'row_avg': True},
                {'pattern': r'(?:average|mean)\s+(?:of\s+)?row\s+(\d+)', 'groups': ['row'], 'row_avg': True},
                # Column-based average patterns
                {'pattern': r'(?:do\s+)?(?:the\s+)?average\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_avg': True},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?average\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_avg': True},
                {'pattern': r'(?:average|mean)\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column'], 'column_avg': True},
                # Cell range patterns
                {'pattern': r'(?:average|mean)\s+(?:of\s+)?(?:cells?\s+)?(?:from\s+)?([A-Za-z]+\d+)\s+(?:to|through)\s+([A-Za-z]+\d+)(?:\s+(?:in|at|to)\s+(?:cell\s+)?([A-Za-z]+\d+))?', 'groups': ['start', 'end', 'result']},
                {'pattern': r'(?:calculate|compute)\s+(?:the\s+)?average\s+(?:of\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'max': [
                {'pattern': r'(?:max|maximum|highest)\s+(?:of\s+)?(?:cells?\s+)?(?:from\s+)?([A-Za-z]+\d+)\s+(?:to|through)\s+([A-Za-z]+\d+)', 'groups': ['start', 'end']},
                {'pattern': r'(?:find|get)\s+(?:the\s+)?(?:max|maximum|highest)\s+(?:value\s+)?(?:in|of)\s+([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'min': [
                {'pattern': r'(?:min|minimum|lowest)\s+(?:of\s+)?(?:cells?\s+)?(?:from\s+)?([A-Za-z]+\d+)\s+(?:to|through)\s+([A-Za-z]+\d+)', 'groups': ['start', 'end']},
                {'pattern': r'(?:find|get)\s+(?:the\s+)?(?:min|minimum|lowest)\s+(?:value\s+)?(?:in|of)\s+([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'formula': [
                {'pattern': r'(?:insert|add|put)\s+formula\s+["\']?(.+?)["\']?\s+(?:in|to|at)\s+(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['formula', 'cell']},
                {'pattern': r'(?:cell\s+)?([A-Za-z]+\d+)\s+(?:formula|equals?)\s+["\']?(.+?)["\']?$', 'groups': ['cell', 'formula']},
            ],
            'chart_bar': [
                {'pattern': r'(?:create|make|add|insert)\s+(?:a\s+)?bar\s+chart(?:\s+(?:for|from|using))?\s*(?:data\s+)?(?:in\s+)?(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)?', 'groups': ['start', 'end']},
                {'pattern': r'(?:create|make|add|insert)\s+(?:a\s+)?bar\s+chart\s+(?:for|from|of)\s+(?:the\s+)?(\w+)\s+data', 'groups': ['data_type']},
            ],
            'chart_pie': [
                {'pattern': r'(?:create|make|add|insert)\s+(?:a\s+)?pie\s+chart(?:\s+(?:for|from|using))?\s*(?:data\s+)?(?:in\s+)?(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)?', 'groups': ['start', 'end']},
                {'pattern': r'(?:create|make|add|insert)\s+(?:a\s+)?pie\s+chart\s+(?:for|from|of)\s+(?:the\s+)?(\w+)\s+data', 'groups': ['data_type']},
            ],
            'chart_line': [
                {'pattern': r'(?:create|make|add|insert)\s+(?:a\s+)?line\s+chart(?:\s+(?:for|from|using))?\s*(?:data\s+)?(?:in\s+)?(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)?', 'groups': ['start', 'end']},
            ],
            'sort': [
                {'pattern': r'sort\s+(?:by\s+)?column\s+([A-Za-z])\s+(?:in\s+)?(?:ascending|asc)', 'groups': ['column'], 'ascending': True},
                {'pattern': r'sort\s+(?:by\s+)?column\s+([A-Za-z])\s+(?:in\s+)?(?:descending|desc)', 'groups': ['column'], 'ascending': False},
                {'pattern': r'sort\s+(?:data\s+)?(?:by\s+)?column\s+([A-Za-z])', 'groups': ['column'], 'ascending': True},
                {'pattern': r'sort\s+(?:the\s+)?data\s+(?:in\s+)?(?:ascending|asc)', 'groups': [], 'ascending': True},
                {'pattern': r'sort\s+(?:the\s+)?data\s+(?:in\s+)?(?:descending|desc)', 'groups': [], 'ascending': False},
            ],
            'filter': [
                {'pattern': r'filter\s+(?:by\s+)?column\s+([A-Za-z])\s+(?:where|for|with)\s+["\']?(.+?)["\']?$', 'groups': ['column', 'value']},
                {'pattern': r'filter\s+(?:data\s+)?(?:where\s+)?([A-Za-z])\s*(?:=|equals?|is)\s*["\']?(.+?)["\']?$', 'groups': ['column', 'value']},
                {'pattern': r'(?:show|display)\s+(?:only\s+)?(?:rows?\s+)?(?:where\s+)?column\s+([A-Za-z])\s+(?:=|equals?|is)\s+["\']?(.+?)["\']?$', 'groups': ['column', 'value']},
            ],
            'create_sheet': [
                {'pattern': r'(?:create|add|make)\s+(?:a\s+)?(?:new\s+)?sheet\s+(?:named?|called?)\s+["\']?(.+?)["\']?$', 'groups': ['name']},
                {'pattern': r'(?:new|add)\s+sheet\s+["\']?(.+?)["\']?$', 'groups': ['name']},
            ],
            'delete_sheet': [
                {'pattern': r'(?:delete|remove)\s+sheet\s+["\']?(.+?)["\']?$', 'groups': ['name']},
            ],
            'rename_sheet': [
                {'pattern': r'rename\s+sheet\s+["\']?(.+?)["\']?\s+(?:to|as)\s+["\']?(.+?)["\']?$', 'groups': ['old_name', 'new_name']},
            ],
            'switch_sheet': [
                {'pattern': r'(?:switch|go)\s+to\s+sheet\s+["\']?(.+?)["\']?$', 'groups': ['name']},
                {'pattern': r'(?:open|select)\s+sheet\s+["\']?(.+?)["\']?$', 'groups': ['name']},
            ],
            'merge_cells': [
                {'pattern': r'merge\s+(?:cells?\s+)?([A-Za-z]+\d+)\s*(?:to|:|-|through)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'border': [
                {'pattern': r'(?:add|put)\s+(?:a\s+)?border\s+(?:to\s+)?(?:cell\s+)?([A-Za-z]+\d+)', 'groups': ['cell']},
                {'pattern': r'(?:add|put)\s+(?:a\s+)?border\s+(?:to\s+)?(?:range\s+)?([A-Za-z]+\d+)\s*(?:to|:|-)\s*([A-Za-z]+\d+)', 'groups': ['start', 'end']},
            ],
            'sample_data': [
                {'pattern': r'(?:create|add|generate|insert)\s+(?:some\s+)?sample\s+(?:sales\s+)?data', 'groups': [], 'type': 'sales'},
                {'pattern': r'(?:create|add|generate)\s+(?:sample\s+)?(?:sales|revenue)\s+data', 'groups': [], 'type': 'sales'},
                {'pattern': r'(?:create|add|generate)\s+(?:sample\s+)?employee\s+data', 'groups': [], 'type': 'employees'},
                {'pattern': r'(?:create|add|generate)\s+(?:sample\s+)?inventory\s+data', 'groups': [], 'type': 'inventory'},
            ],
            'context_chart': [
                {'pattern': r'(?:create|make|show|display)\s+(?:a\s+)?pie\s+chart\s+(?:for|of|showing)\s+(\w+)\s+(?:by|per|for each)\s+(\w+)', 'groups': ['data_field', 'category_field']},
                {'pattern': r'(?:create|make|show)\s+(?:a\s+)?bar\s+chart\s+(?:for|of|showing)\s+(\w+)\s+(?:by|per|for each)\s+(\w+)', 'groups': ['data_field', 'category_field']},
                {'pattern': r'(?:visualize|chart|graph)\s+(\w+)\s+(?:by|per|grouped by)\s+(\w+)', 'groups': ['data_field', 'category_field']},
            ],
            'context_analysis': [
                {'pattern': r'(?:show|calculate|get|find)\s+(?:the\s+)?average\s+(?:of\s+)?(?:column\s+)?([A-Za-z])\s+(?:for|where|when)\s+(\w+)\s+(?:only|equals?|is)?', 'groups': ['column', 'filter_context']},
                {'pattern': r'(?:show|calculate|what(?:\'s| is))\s+(?:the\s+)?(?:total|sum)\s+(?:of\s+)?(\w+)\s+(?:for|in|during)\s+(\w+)', 'groups': ['data_field', 'time_context']},
                {'pattern': r'(?:filter|show|display)\s+(?:data\s+)?(?:for|where)\s+(\w+)\s+(?:is|equals?|=)\s+["\']?(\w+)["\']?', 'groups': ['field', 'value']},
                {'pattern': r'(?:analyze|summarize)\s+(\w+)\s+(?:data|values)?', 'groups': ['data_type']},
            ],
            'save': [
                {'pattern': r'save\s+(?:the\s+)?(?:workbook|file|document)', 'groups': []},
                {'pattern': r'save\s+(?:as\s+)?["\']?(.+?)["\']?$', 'groups': ['filename']},
            ],
            'new_workbook': [
                {'pattern': r'(?:create|start|open)\s+(?:a\s+)?new\s+(?:workbook|file|spreadsheet)', 'groups': []},
            ],
            'list_sheets': [
                {'pattern': r'(?:list|show|display)\s+(?:all\s+)?sheets', 'groups': []},
                {'pattern': r'what\s+sheets?\s+(?:do\s+)?(?:we\s+)?have', 'groups': []},
            ],
            'help': [
                {'pattern': r'^help$|^what can you do|^commands?$|^show commands', 'groups': []},
            ],
            'copy_cell': [
                {'pattern': r'copy\s+(?:cell\s+)?([A-Za-z]+\d+)\s+to\s+([A-Za-z]+\d+)', 'groups': ['source', 'destination']},
                {'pattern': r'copy\s+([A-Za-z]+\d+)\s+(?:to|into)\s+([A-Za-z]+\d+)', 'groups': ['source', 'destination']},
            ],
            'find': [
                {'pattern': r'find\s+["\']?(.+?)["\']?(?:\s+in\s+sheet)?$', 'groups': ['value']},
                {'pattern': r'search\s+(?:for\s+)?["\']?(.+?)["\']?$', 'groups': ['value']},
                {'pattern': r'where\s+is\s+["\']?(.+?)["\']?$', 'groups': ['value']},
            ],
            'replace': [
                {'pattern': r'replace\s+["\']?(.+?)["\']?\s+with\s+["\']?(.+?)["\']?$', 'groups': ['search', 'replace']},
                {'pattern': r'find\s+and\s+replace\s+["\']?(.+?)["\']?\s+(?:with|to)\s+["\']?(.+?)["\']?$', 'groups': ['search', 'replace']},
            ],
            'count_column': [
                {'pattern': r'count\s+(?:cells?\s+in\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
                {'pattern': r'how\s+many\s+(?:cells?\s+)?(?:in\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
            ],
            'max_column': [
                {'pattern': r'(?:max|maximum)\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
                {'pattern': r'(?:find|get|what(?:\'s| is))\s+(?:the\s+)?(?:max|maximum|highest)\s+(?:value\s+)?(?:in\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
            ],
            'min_column': [
                {'pattern': r'(?:min|minimum)\s+(?:of\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
                {'pattern': r'(?:find|get|what(?:\'s| is))\s+(?:the\s+)?(?:min|minimum|lowest)\s+(?:value\s+)?(?:in\s+)?column\s+([A-Za-z]+)', 'groups': ['column']},
            ],
            'auto_fit': [
                {'pattern': r'auto\s*fit\s+(?:all\s+)?columns?', 'groups': []},
                {'pattern': r'fit\s+(?:column\s+)?widths?', 'groups': []},
                {'pattern': r'adjust\s+(?:column\s+)?widths?', 'groups': []},
            ],
            'clear_sheet': [
                {'pattern': r'clear\s+(?:all\s+)?(?:the\s+)?sheet', 'groups': []},
                {'pattern': r'clear\s+(?:all\s+)?data', 'groups': []},
                {'pattern': r'delete\s+all\s+data', 'groups': []},
            ],
            'freeze': [
                {'pattern': r'freeze\s+(?:panes?\s+)?(?:at\s+)?(?:row\s+)?(\d+)?', 'groups': ['row']},
                {'pattern': r'freeze\s+(?:the\s+)?(?:top\s+)?(?:header\s+)?row', 'groups': []},
            ],
            'unfreeze': [
                {'pattern': r'unfreeze\s+(?:panes?|all)?', 'groups': []},
            ],
        }
    
    def parse_number(self, text: str) -> Optional[int]:
        """Parse a number from text (handles both digits and words)."""
        text = text.lower().strip()
        if text.isdigit():
            return int(text)
        return self.number_words.get(text)
    
    def get_color_code(self, color_name: str) -> Optional[str]:
        """Get hex color code from color name."""
        return self.color_map.get(color_name.lower())
    
    def interpret_command(self, command: str) -> Dict[str, Any]:
        """
        Interpret a natural language command and return structured action.
        
        Args:
            command: Natural language command string
            
        Returns:
            Dictionary with 'action' and 'params' keys
        """
        command = command.strip().lower()
        
        # Fix common speech recognition errors
        command = self._fix_speech_errors(command)
        
        # Try AI-powered interpretation first if available
        if self.openai_client:
            try:
                result = self._interpret_with_ai(command)
                if result and result.get('action') != 'unknown':
                    return result
            except Exception as e:
                print(f"AI interpretation failed: {e}")
        
        # Fall back to rule-based interpretation
        return self._interpret_with_rules(command)
    
    def _fix_speech_errors(self, command: str) -> str:
        """Fix common speech recognition errors with context awareness."""
        
        # Context-specific corrections for "some" → "sum"
        # Only convert when followed by Excel-related words
        excel_context_patterns = [
            (r'\bsome\s+of\s+row\b', 'sum of row'),
            (r'\bsome\s+of\s+column\b', 'sum of column'),
            (r'\bsome\s+of\s+col\b', 'sum of column'),
            (r'\bsome\s+row\b', 'sum row'),
            (r'\bsome\s+column\b', 'sum column'),
            (r'\bsome\s+col\b', 'sum column'),
            (r'\bdo\s+some\s+of\s+row\b', 'do sum of row'),
            (r'\bdo\s+some\s+of\s+column\b', 'do sum of column'),
            (r'\bdo\s+some\s+row\b', 'do sum row'),
            (r'\bdo\s+some\s+column\b', 'do sum column'),
            (r'\bdo\s+some\s+of\b', 'do sum of'),  # Generic fallback
            (r'\bcalculate\s+some\b', 'calculate sum'),
            (r'\bfind\s+some\b', 'find sum'),
            (r'\bget\s+some\b', 'get sum'),
            (r'\badd\s+some\b', 'add sum'),
            (r'\btotal\s+some\b', 'total sum'),
            (r'\bthe\s+some\s+of\b', 'the sum of'),
            (r'\bshow\s+some\s+of\b', 'show sum of'),
        ]
        
        for pattern, replacement in excel_context_patterns:
            command = re.sub(pattern, replacement, command, flags=re.IGNORECASE)
        
        # General word corrections
        general_replacements = {
            r'\braw\b': 'row',  # "raw" → "row"
            r'\braws\b': 'rows',
            r'\bcaller\b': 'column',
            r'\bcalm\b': 'column',
            r'\bcolum\b': 'column',
            r'\bsheet\s+one\b': 'sheet 1',
            r'\bsheet\s+to\b': 'sheet 2',
            r'\bsheet\s+too\b': 'sheet 2',
            r'\bsheet\s+three\b': 'sheet 3',
            r'\bbowled\b': 'bold',
            r'\bweight\b': 'write',
            r'\bright\s+(\d+|to)\b': r'write \1',  # "right 100 to" → "write 100 to"
        }
        
        for pattern, replacement in general_replacements.items():
            command = re.sub(pattern, replacement, command, flags=re.IGNORECASE)
        
        return command
    
    def _interpret_with_ai(self, command: str) -> Dict[str, Any]:
        """Use OpenAI to interpret the command."""
        system_prompt = """You are an Excel command interpreter. Analyze the user's natural language command and return a JSON response with the action and parameters.

Available actions and their parameters:
- write_cell: {cell, value}
- read_cell: {cell}
- clear_cell: {cell} or {start_cell, end_cell}
- format_bold: {cell} or {start_cell, end_cell}
- format_italic: {cell}
- format_color: {cell, bg_color} (color as 6-digit hex without #)
- format_currency: {cell}
- insert_row: {row, count}
- delete_row: {row, count}
- insert_column: {column, count}
- delete_column: {column, count}
- sum: {start, end, result_cell}
- average: {start, end, result_cell}
- max: {start, end, result_cell}
- min: {start, end, result_cell}
- formula: {cell, formula}
- chart_bar: {data_range, title}
- chart_pie: {data_range, title}
- chart_line: {data_range, title}
- sort: {column, ascending}
- filter: {column, value}
- create_sheet: {name}
- delete_sheet: {name}
- rename_sheet: {old_name, new_name}
- switch_sheet: {name}
- merge_cells: {start, end}
- border: {cell} or {start, end}
- sample_data: {type} (sales/employees/inventory)
- save: {filename} (optional)
- new_workbook: {}
- list_sheets: {}
- help: {}

Respond with JSON only: {"action": "action_name", "params": {...}}
If you can't understand the command, respond with: {"action": "unknown", "params": {}, "message": "explanation"}"""

        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": command}
                ],
                max_tokens=200,
                temperature=0.1
            )
            
            import json
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            print(f"AI parsing error: {e}")
            return None
    
    def _interpret_with_rules(self, command: str) -> Dict[str, Any]:
        """Use rule-based pattern matching to interpret the command."""
        command = command.lower().strip()
        
        # Check each command type
        for action_type, patterns in self.command_patterns.items():
            for pattern_info in patterns:
                pattern = pattern_info['pattern']
                match = re.search(pattern, command, re.IGNORECASE)
                
                if match:
                    params = {}
                    groups = pattern_info.get('groups', [])
                    
                    for i, group_name in enumerate(groups):
                        if i < len(match.groups()) and match.group(i + 1):
                            value = match.group(i + 1).strip()
                            
                            # Handle special conversions
                            if group_name in ['row', 'count']:
                                parsed = self.parse_number(value)
                                params[group_name] = parsed if parsed else 1
                            elif group_name == 'color':
                                color_code = self.get_color_code(value)
                                params['bg_color'] = color_code if color_code else value
                            elif group_name == 'cell':
                                params[group_name] = value.upper()
                            elif group_name in ['start', 'end', 'start_cell', 'end_cell', 'result', 'result_cell']:
                                params[group_name] = value.upper()
                            elif group_name == 'column':
                                params[group_name] = value.upper()
                            else:
                                params[group_name] = value
                    
                    # Add any static params from pattern definition
                    if 'ascending' in pattern_info:
                        params['ascending'] = pattern_info['ascending']
                    if 'type' in pattern_info:
                        params['type'] = pattern_info['type']
                    
                    # Handle row-based sum
                    if pattern_info.get('row_sum') and 'row' in params:
                        row_num = params['row']
                        return {
                            'action': 'sum_row',
                            'params': {'row': int(row_num)}
                        }
                    
                    # Handle row-based average
                    if pattern_info.get('row_avg') and 'row' in params:
                        row_num = params['row']
                        return {
                            'action': 'average_row',
                            'params': {'row': int(row_num)}
                        }
                    
                    # Handle column-based average
                    if pattern_info.get('column_avg') and 'column' in params:
                        col = params['column'].upper()
                        return {
                            'action': 'average_column',
                            'params': {'column': col}
                        }
                    
                    # Handle column-based sum
                    if pattern_info.get('column_sum') and 'column' in params:
                        col = params['column'].upper()
                        return {
                            'action': 'sum_column',
                            'params': {'column': col}
                        }
                    
                    # Handle format_column - convert color name to hex
                    if action_type == 'format_column':
                        column = params.get('column', 'A').upper()
                        color_name = params.get('color', '').lower()
                        color_code = self.get_color_code(color_name)
                        return {
                            'action': 'format_column',
                            'params': {
                                'column': column,
                                'bg_color': color_code if color_code else 'FFFF00'  # default yellow
                            }
                        }
                    
                    # Handle context-aware actions
                    if action_type == 'context_chart':
                        # Convert to appropriate chart type
                        if 'pie' in command:
                            return {'action': 'chart_pie', 'params': {'data_range': 'A1:B10', 'title': f"{params.get('data_field', 'Data')} by {params.get('category_field', 'Category')}", 'context': params}}
                        else:
                            return {'action': 'chart_bar', 'params': {'data_range': 'A1:B10', 'title': f"{params.get('data_field', 'Data')} by {params.get('category_field', 'Category')}", 'context': params}}
                    
                    if action_type == 'context_analysis':
                        # Handle context-aware analysis
                        params['context_aware'] = True
                        if 'average' in command:
                            return {'action': 'average', 'params': {'start': 'A1', 'end': 'A10', 'context': params}}
                        elif 'sum' in command or 'total' in command:
                            return {'action': 'sum', 'params': {'start': 'A1', 'end': 'A10', 'context': params}}
                        else:
                            return {'action': 'filter', 'params': params}
                    
                    return {
                        'action': action_type,
                        'params': params
                    }
        
        # Try fuzzy matching for common intents
        result = self._fuzzy_match_intent(command)
        if result:
            return result
        
        # If no pattern matched
        return {
            'action': 'unknown',
            'params': {},
            'message': f"I couldn't understand the command: '{command}'. Try saying 'help' for available commands."
        }
    
    def _fuzzy_match_intent(self, command: str) -> Optional[Dict[str, Any]]:
        """Fuzzy matching for common intents when exact patterns don't match."""
        command = command.lower()
        
        # Chart intent detection
        if any(word in command for word in ['chart', 'graph', 'visualize', 'plot']):
            chart_type = 'bar'
            if 'pie' in command:
                chart_type = 'pie'
            elif 'line' in command:
                chart_type = 'line'
            
            title = "Data Visualization"
            # Extract context from command
            if 'sales' in command:
                title = "Sales Chart"
            elif 'revenue' in command:
                title = "Revenue Chart"
            elif 'employee' in command:
                title = "Employee Data"
            
            return {
                'action': f'chart_{chart_type}',
                'params': {'data_range': 'A1:B10', 'title': title}
            }
        
        # Formula intent detection  
        if any(word in command for word in ['sum', 'add up', 'total']):
            return {'action': 'sum', 'params': {'start': 'A1', 'end': 'A10'}}
        
        if any(word in command for word in ['average', 'mean', 'avg']):
            return {'action': 'average', 'params': {'start': 'A1', 'end': 'A10'}}
        
        if any(word in command for word in ['maximum', 'max', 'highest', 'largest']):
            return {'action': 'max', 'params': {'start': 'A1', 'end': 'A10'}}
        
        if any(word in command for word in ['minimum', 'min', 'lowest', 'smallest']):
            return {'action': 'min', 'params': {'start': 'A1', 'end': 'A10'}}
        
        # Data creation intent
        if 'sample' in command or 'demo' in command or 'example' in command:
            data_type = 'sales'
            if 'employee' in command:
                data_type = 'employees'
            elif 'inventory' in command:
                data_type = 'inventory'
            return {'action': 'sample_data', 'params': {'type': data_type}}
        
        return None
    
    def get_help_text(self) -> str:
        """Return help text with available commands."""
        return """
📊 **Voice-Controlled Excel Assistant - Available Commands**

**Data Entry:**
• "Write 'Hello' to cell A1"
• "Set cell B2 to 100"
• "Read cell A1"
• "Clear cell B2" / "Clear cells A1 to B10"

**Formatting:**
• "Make cell A1 bold"
• "Format B2 as italic"
• "Color cell C3 red" / "Highlight A1 yellow"
• "Format D4 as currency"
• "Add border to A1:D10"
• "Merge cells A1 to C1"

**Rows & Columns:**
• "Insert row at row 5"
• "Delete row 3"
• "Insert column at column B"
• "Delete column C"

**Formulas:**
• "Sum A1 to A10 in cell A11"
• "Average B1 through B10"
• "Find the maximum in C1 to C20"
• "Insert formula =B1*C1 in cell D1"

**Charts:**
• "Create a bar chart for A1:B10"
• "Create a pie chart for sales data"
• "Create a line chart"

**Data Operations:**
• "Sort by column A ascending"
• "Sort data descending"
• "Filter column B where value is Sales"

**Sheets:**
• "Create new sheet named 'Report'"
• "Delete sheet 'Old Data'"
• "Switch to sheet 'Summary'"
• "List all sheets"

**Sample Data:**
• "Create sample sales data"
• "Add sample employee data"
• "Generate inventory data"

**File Operations:**
• "Save the workbook"
• "Create new workbook"
"""

