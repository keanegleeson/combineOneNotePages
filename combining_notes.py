#script to automate the process of extracting my notes for annual summariet


import time
import pyautogui
# from docx import Document


#give me a sec to move to OneNote
time.sleep(5)

# Define constant coordinates
file_button_coords = (25, 45)
export_button_coords = (66, 318)
word_export_button_coords = (658, 235)
save_button_coords = (576, 478)

# Define the initial coordinates and file name prefix
x_coordinate = 100  # Adjust this value as needed
y_coordinate = 200  # Initial y-coordinate
file_name_prefix = 'Page_'  # Prefix for the Word file names



# Assuming you have OneNote open and the first page is displayed.

# Iterate through the pages (assuming you have 52 pages)
for page_number in range(1, 3):
    # Click the "File" button
    pyautogui.click(file_button_coords)

    # Click the "Export" button once in the File view
    pyautogui.click(export_button_coords)

    # Click the button for exporting to Word
    pyautogui.click(word_export_button_coords)

    # Click the button to export the page to the Word doc
    pyautogui.click(save_button_coords)

    # Wait for the save dialog to appear (adjust the sleep duration as needed)
    time.sleep(2)

    # Type the file name and path
    file_name = f'{file_name_prefix}{page_number}.docx'
    # pyautogui.write(file_name)

    # Press Enter to save the file
    pyautogui.press('enter')

    # Wait for the save process to complete (adjust the sleep duration)
    time.sleep(5)

    # Use Ctrl+Page Down to navigate to the next page
    pyautogui.hotkey('ctrl', 'pagedown')

    time.sleep(5)



# Combine the exported Word files into one merged document
# def combine_word_documents(files):
#     merged_document = Document()

#     for file in files:
#         sub_doc = Document(file)
#         for element in sub_doc.element.body:
#             merged_document.element.body.append(element)

#     merged_document.save('2022_merged.docx')

# # Combine the Word files into a single merged document
# combine_word_documents(word_files)
