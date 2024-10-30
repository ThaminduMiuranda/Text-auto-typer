import win32com.client as win32


def extract_text_with_styles(file_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(file_path)

    text_with_styles = []

    for paragraph in doc.Paragraphs:
        text_info = {
            "text": paragraph.Range.Text.strip(),
            "font_name": paragraph.Range.Font.Name,
            "font_size": paragraph.Range.Font.Size,
            "bold": paragraph.Range.Font.Bold,
            "italic": paragraph.Range.Font.Italic,
            "indentation": paragraph.Format.LeftIndent,
        }
        text_with_styles.append(text_info)

    doc.Close()
    word.Quit()

    return text_with_styles


import pyautogui
import time
import random


def type_with_styles(text_with_styles):
    for text_info in text_with_styles:
        # Set Google Docs style based on extracted information
        apply_style_in_google_docs(
            font_name=text_info["font_name"],
            font_size=text_info["font_size"],
            bold=text_info["bold"],
            italic=text_info["italic"],
            indentation=text_info["indentation"]
        )

        # Type each character with a delay and occasional mistakes
        for char in text_info["text"]:

            if random.random() < 0.05:
                pyautogui.write(random.choice("abcdefghijklmnopqrstuvwxyz"))
                pyautogui.press("backspace")

            pyautogui.write(char)

        pyautogui.press("enter")  # Move to the next paragraph


def apply_style_in_google_docs(font_name, font_size, bold, italic, indentation):
    # Adjust font settings in Google Docs manually before running the script
    # This code assumes you manually adjust indentation or font settings based on prompts
    print(f"Set font to: {font_name}, size: {font_size}, bold: {bold}, italic: {italic}, indentation: {indentation}")

file_path = "..\input.docx"
time.sleep(10)
text_with_styles = extract_text_with_styles(file_path)

# Ensure Google Docs cursor is active
type_with_styles(text_with_styles)
