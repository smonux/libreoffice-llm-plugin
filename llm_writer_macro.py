import uno
import unohelper
import datetime
import urllib.request
import urllib.parse
import json
import traceback
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
import os

PARAMS_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_params.json")
LOG_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_api_logs.log")

def init_db():
        """Initialize JSON files for parameters and logs"""
        # Initialize parameters file
        if not os.path.exists(PARAMS_PATH):
            with open(PARAMS_PATH, 'w') as f:
                json.dump({
                    'OPENAI_ENDPOINT': 'https://api.openai.com/v1/chat/completions',
                    'OPENAI_API_KEY': '',
                    'MODEL': 'gpt-4o-mini',
                    'MAX_GENERATION_TOKENS': '30',
                    'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally in the [COMPLETE] ',
                    'CONTEXT_PREVIOUS_CHARS': '100',
                    'CONTEXT_NEXT_CHARS': '100',
                    'TEMPERATURE': '0.7'
                }, f, indent=4)
        
        # Initialize logs file
        if not os.path.exists(LOG_PATH):
            with open(LOG_PATH, 'w') as f:
                f.write("initializing...\n")

def get_param(key):
        """Get parameter from JSON file"""
        with open(PARAMS_PATH, 'r') as f:
            params = json.load(f)
            return params.get(key)

def set_param(key, value):
        """Set parameter in JSON file atomically"""
        # Load existing params first
        with open(PARAMS_PATH, 'r') as f:
            params = json.load(f)
        
        # Update the requested parameter
        params[key] = value
        
        # Write to temporary file first to prevent corruption
        temp_path = PARAMS_PATH + ".tmp"
        with open(temp_path, 'w') as f:
            json.dump(params, f)
        
        # Atomically replace original file
        os.replace(temp_path, PARAMS_PATH)

def get_context(cursor):
        """Get previous and next tokens around cursor position"""
        text = cursor.getText()
        prev_chars = int(get_param('CONTEXT_PREVIOUS_CHARS'))
        next_chars = int(get_param('CONTEXT_NEXT_CHARS'))

        text_cursor = cursor.getText().createTextCursorByRange(cursor)
        text_cursor.goLeft(prev_chars, True)
        previous_context = text_cursor.getString()

        text_cursor = cursor.getText().createTextCursorByRange(cursor)
        text_cursor.goRight(next_chars, True)
        next_context = text_cursor.getString()

        #show_message(previous_context + next_context)
        
        return previous_context, next_context


def call_llm(data): 
        """Make API call to OpenAI-compatible endpoint"""
        url = get_param('OPENAI_ENDPOINT') 
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f"Bearer {get_param('OPENAI_API_KEY')}"
        }
        
        req = urllib.request.Request(url, 
                                   data=json.dumps(data).encode('utf-8'),
                                   headers=headers)
        
        try:
            with urllib.request.urlopen(req) as response:
                response_data = json.loads(response.read().decode('utf-8'))
                _log_api_call(url, data, response_data, response.status)
                return response_data
        except urllib.error.HTTPError as e:
            # Read the response body to get the detailed error message
            error_response = e.read().decode('utf-8')
            _log_api_call(url, data, error_response, e.code)
            raise

def _log_api_call(endpoint, request, response, status_code):
        """Log API call details to a regular text file"""
        with open(LOG_PATH, 'a') as f:
            f.write(f"Timestamp: {datetime.datetime.now().isoformat()}\n")
            f.write(f"Endpoint: {endpoint}\n")
            f.write(f"Status Code: {status_code}\n")
            f.write(f"Request: {request}\n")
            f.write(f"Response: {response}\n")
            f.write("-" * 40 + "\n")

def get_api_logs(limit=100):
        """Retrieve API logs from text file"""
        with open(LOG_PATH, 'r') as f:
            logs = f.readlines()[:limit]
            return logs

def show_message(message):
        """Show message dialog"""
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        parent = toolkit.getDesktopWindow()
        msgbox = toolkit.createMessageBox(
            parent,  "infobox", MSG_BUTTONS.BUTTONS_OK, "LLM Writer", str(message))
        msgbox.execute()

def _get_cursor():
        """Get text cursor from current selection"""
        xModel = uno.getComponentContext().getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", uno.getComponentContext()).getCurrentComponent()
        xSelectionSupplier = xModel.getCurrentController()
        xIndexAccess = xSelectionSupplier.getSelection()
        return xIndexAccess.getByIndex(0)

def autocomplete():
        """Generate autocomplete suggestions using LLM"""
        try:
            cursor = _get_cursor()
            previous_context, next_context = get_context(cursor)
            
            data = {
                'model': get_param('MODEL'),
                'messages': [
                    {"role": "system", "content": get_param('AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS')},
                    {"role": "user", "content": f"{previous_context}[COMPLETE HERE]{next_context}"}
                ],
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': float(get_param('TEMPERATURE')),
                'stop': ['\n'] 
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['message']['content'])
                
        except Exception as e:
            error_msg = f"ERROR: {str(e)}"
            full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
            show_message(full_error) 

def transform_text():
        """Transform selected text based on instruction"""
        try:
            cursor = _get_cursor()
            instruction = show_input_dialog("Enter transformation instructions:")
            selected_text = cursor.getString()
            
            if not instruction:
                return
                
            data = {
                'model': get_param('MODEL'),
                'messages': [
                    {"role": "system", "content": instruction},
                    {"role": "user", "content": f"Original text: {selected_text}\n\nTransformed text:"}
                ],
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': float(get_param('TEMPERATURE'))
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['message']['content'])
                
        except Exception as e:
            error_msg = f"ERROR: {str(e)}"
            full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
            show_message(full_error) 

def show_logs():
        """Display API logs in message box"""
        logs = get_api_logs()
        if not logs:
            show_message("No API logs found")
            return
            
        log_text = "API Logs:\n\n"
        for log in logs:
            log_text += log + "\n"
        
        show_message(log_text)

 def show_input_dialog(message):
     """Show input dialog"""
     ctx = uno.getComponentContext()
     sm = ctx.getServiceManager()
     toolkit = sm.createInstanceWithContext(
         "com.sun.star.awt.Toolkit", ctx)

     # Create a dialog model
     dialog_model = toolkit.createUnoControlDialogModel()
     dialog_model.Width = 200
     dialog_model.Height = 100
     dialog_model.Title = "LLM Writer"

     # Add a text field
     text_field_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
     text_field_model.Name = "TextField"
     text_field_model.PositionX = 10
     text_field_model.PositionY = 20
     text_field_model.Width = 180
     text_field_model.Height = 10
     dialog_model.insertByName("TextField", text_field_model)

     # Add a label
     label_model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
     label_model.Name = "Label"
     label_model.PositionX = 10
     label_model.PositionY = 5
     label_model.Width = 180
     label_model.Height = 10
     label_model.Label = message
     dialog_model.insertByName("Label", label_model)

     # Add an OK button
     ok_button_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
     ok_button_model.Name = "OKButton"
     ok_button_model.PositionX = 60
     ok_button_model.PositionY = 40
     ok_button_model.Width = 80
     ok_button_model.Height = 15
     ok_button_model.Label = "OK"
     dialog_model.insertByName("OKButton", ok_button_model)

     # Create the dialog
     dialog = toolkit.createUnoControlDialog(dialog_model)

     # Event handler for the OK button
     def button_action_handler(event):
         dialog.endExecute(1)

     ok_button = dialog.getControl("OKButton")
     ok_button.addActionListener(lambda x: button_action_handler(x))

     # Show the dialog
     dialog.execute()

     # Get the text from the text field
     text_field = dialog.getControl("TextField")
     return text_field.Text if dialog.getResult() == 1 else None



# Export the macros properly
init_db()
g_exportedScripts = (autocomplete, transform_text, show_logs)

