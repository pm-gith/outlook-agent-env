from summarizer import process_json_file, save_to_word
import os, glob

#Find latest file from OneDrive folder
FOLDER_PATH = r"C:\Users\prachimishra\OneDrive - LucidMotors\Desktop\email-summaris"
list_of_files = glob.glob(os.path.join(FOLDER_PATH, "*.json")) #get all .json files
filename = max(list_of_files, key=os.path.getctime)

#Select a specific file
# filename = r"C:\Users\prachimishra\OneDrive - LucidMotors\Desktop\email-summaris\summary_2025-07-23_18-34.json"

print("Loading email summary from {filename}")
summary = process_json_file(filename)

for section, items in summary.items():
    print(f"\n=== {section} ===")
    for item in items:
        print(item)

output_doc = os.path.join(FOLDER_PATH, "summary_for_copilot.docx")
# Check if the file already exists, and if so, remove it
if os.path.exists(output_doc):
    os.remove(output_doc)  # Deletes the file if it exists
# Save the summary to a Word document
save_to_word(summary, output_doc)
print(f"File ready for Copilot summarization: {output_doc}")