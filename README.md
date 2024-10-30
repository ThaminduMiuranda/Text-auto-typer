# Word-to-Google-Docs Typer

This project extracts text and formatting from a Word document and simulates typing it into Google Docs as if it were done by a human, complete with occasional typing mistakes and corrections. This simulation is ideal for scenarios where you want text transferred with human-like imperfections.

## Features

- **Text Extraction**: Reads text from a Microsoft Word document, including font styles, font sizes, bold/italic attributes, and indentation.
- **Human-like Typing Simulation**: Types each character with variable delays, randomly inserting mistakes and corrections for a natural typing effect.
- **Style Prompts**: Displays style information for manual adjustments in Google Docs to match the original formatting.

## Prerequisites

- **Python** 3.6 or higher
- **Microsoft Word** installed (for Windows only, required for `win32com.client` functionality)
- **Google Docs** open and ready for typing

### Required Python Libraries

Install the necessary Python libraries:

```bash
pip install pywin32 pyautogui python-docx
```

## Usage

### Step 1: Prepare Word Document

Ensure that the Word document you want to use is saved and available. Note the file path for use in the script.

### Step 2: Run the Script

1. Run the Python script and Open **Google Docs** and place the cursor in the document withing the time delay where you want to start typing.

  ```python
file_path = "absolute/path/to/your/word_document.docx"
time.sleep(10)
text_with_styles = extract_text_with_styles(file_path)

# Ensure Google Docs cursor is active
type_with_styles(text_with_styles)
```

2. The script will begin typing the contents of the Word document with simulated human typing.

### Script Breakdown

- **`extract_text_with_styles(file_path)`**: Reads text and basic styles from the Word document.
- **`type_with_styles(text_with_styles)`**: Simulates typing the text with human-like delays, errors, and corrections.
- **`apply_style_in_google_docs()`**: Prints style settings for the paragraph, which you can adjust manually in Google Docs as needed.

## Example

In your script, replace `file_path` with the path to your Word document:

```python
file_path = "path/to/your/word_document.docx"
text_with_styles = extract_text_with_styles(file_path)
type_with_styles(text_with_styles)
```

## Notes

- **Manual Styling**: Due to limitations in Google Docs control, styling changes (e.g., font, size, and indentation) must be applied manually based on prompts.
- **Focus Requirement**: Ensure that Google Docs is active and the cursor is positioned correctly before running the script, as `pyautogui` will type wherever the cursor is.
- **Compatibility**: The script only works on Windows systems with Microsoft Word installed due to the use of `win32com.client`.

## License

This project is licensed under the MIT License.
