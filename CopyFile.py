import shutil


from docx import Document


def FindWord(path, text_search):
 doc = Document(path)
 for paragraph in doc.paragraphs:
     if text_search in paragraph.text:
         return True
 return False


old_file = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\Results\\Folders-123_01\\AABAI_Bulk_Limit_01\DOC\\AABAI_Bulk_Limit_01.docx"

new_path = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\New Word Folder"


text_to_find = "Execution Status	: Passed"


if FindWord(old_file, text_to_find):
 print("Word Found")
else:
 print("Word Not Found")
