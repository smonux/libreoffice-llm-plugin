 =======
 # LLM Writer Macro for LibreOffice/OpenOffice

 This macro provides AI-powered text autocompletion and transformation capabilities within LibreOffice/OpenOffice using OpenAI's API.

 ## Features

 - **Autocomplete**: Continue writing naturally from the cursor position
 - **Text Transformation**: Modify selected text based on instructions
 - **Customizable**: Configure API settings and prompt templates
 - **Logging**: View API call history and debug information

 [Short demonstration](/images/llm_macro.gif)

 ## Installation

 1. Save the `llm_writer_macro.py` file to your LibreOffice macro directory:
    - Linux: `~/.config/libreoffice/4/user/Scripts/python/`
    - Windows: `%APPDATA%\LibreOffice\4\user\Scripts\python\`
    - macOS: `~/Library/Application Support/LibreOffice/4/user/Scripts/python/`

 2. Restart LibreOffice/OpenOffice

 3. Set up your OpenAI API key:
    - Go to Tools > Macros > Organize Macros > Python
    - Run the `modify_config` macro
    - Enter your OpenAI API key and other settings

 ## Usage

 ### Assigning to Toolbar

 1. Go to Tools > Customize
 2. Select the toolbar you want to add the macro to
 3. Click "Add..."
 4. In the Category list, select "Python"
 5. Choose the desired macro (autocomplete, transform_text, etc.)
 6. Click "Add" and "Close"

 ### Assigning Keyboard Shortcut

 1. Go to Tools > Customize > Keyboard
 2. Select a function key combination
 3. In the Category list, select "Python"
 4. Choose the desired macro
 5. Click "Modify" and "OK"

 ### Using the Macros

 - **Autocomplete**: Place cursor where you want text to continue and run the macro
 - **Transform Text**: Select text to modify, run the macro, and enter instructions. 
 In absence of instructions, the llm will run the explicit or implied instruction 
 of the text (for example, translate this: xxx).
 The original text may be kept or replaced by the generation.
 - **View Logs**: Run the `show_logs` macro to see API call history
 - **Modify Config**: Run the `modify_config` macro to change API settings

 ## Configuration

 The following parameters can be configured:

 - OPENAI_ENDPOINT: API endpoint URL (may be used with openrouter or any other 
 compatible platform).
 - OPENAI_API_KEY: Your OpenAI API key
 - MODEL: GPT model to use (e.g., gpt-4o, gpt-4o-mini, etc..)
 - MAX_GENERATION_WORDS: Maximum words to generate (aprox)
 - AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS: Prompt template for autocomplete
 - CONTEXT_PREVIOUS_CHARS: Number of previous characters to use as context
 - CONTEXT_NEXT_CHARS: Number of following characters to use as context
 - TEMPERATURE: Creativity level (0.0 to 2.0)

 ## Requirements

 - LibreOffice/OpenOffice with Python support
 - OpenAI-compatible API key (OpenAI, OpenRouter, Ollama, vLLM, etc...) 
 - Internet connection (unless running server locally)

 ## Troubleshooting

 - Check API logs using the `show_logs` macro
 - Verify your API key is correct
 - Ensure you have an active internet connection
 - Make sure the macro file is in the correct directory

 ## Configuration and Log Files

The macro uses two files stored in the user's home directory:

1. **llm_writer_params.json**  
   Location:  
     - Linux/macOS: `~/llm_writer_params.json`  
     - Windows: `%USERPROFILE%\llm_writer_params.json`  
   Purpose: Stores all configuration parameters including API key, model settings, and prompt templates.  
   Format: JSON file that can be manually edited if needed.

2. **llm_writer_api_logs.log**  
   Location:  
     - Linux/macOS: `~/llm_writer_api_logs.log`  
     - Windows: `%USERPROFILE%\llm_writer_api_logs.log`  
   Purpose: Logs all API calls made by the macro including requests, responses, and timestamps.  
   Format: Plain text file that can be viewed with any text editor.

These files are automatically created when the macro is first run. The configuration file can be modified either through the macro's configuration dialog or by directly editing the JSON file.

 ## License

 This project is licensed under the MIT License - see the LICENSE file for details.
