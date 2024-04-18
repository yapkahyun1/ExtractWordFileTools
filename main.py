import shutil
import sys
import os
import wx
from datetime import datetime
from docx import Document

def PromptWindows(path):
    app = wx.App(False)
    frame = wx.Frame(None, wx.ID_ANY, "Test Case Results Word Docs Extractor")
    panel = wx.Panel(frame, wx.ID_ANY)

    radio = wx.RadioButton(panel, label="Passed Results", style=wx.RB_GROUP)
    radio2 = wx.RadioButton(panel, label="Failed Results")

    def onBtn(event):
        currentTimeAndDate = datetime.now()
        StrDateTime = currentTimeAndDate.strftime("%d%m%Y_%H%M%S")

        if radio.GetValue() == True:
            Pass_Folder_Name = "Passed Results" + "_" + str(StrDateTime)
            new_pass_Path = os.path.join(path, Pass_Folder_Name)
            os.mkdir(new_pass_Path)
            text_to_find = "Execution Status	: Passed"
            ExtractDocFile(path, text_to_find, new_pass_Path)
            dlg = wx.MessageDialog(frame, "All Passed Results Word Document has been extracted completed\n" + "Result saved to this location: \n" + str(new_pass_Path),
                                   "Word Document Extraction",
                                   wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

        elif radio2.GetValue() == True:
            Fail_Folder_Name = "Failed Results" + "_" + str(StrDateTime)
            new_fail_path = os.path.join(path, Fail_Folder_Name)
            os.mkdir(new_fail_path)
            text_to_find = "Execution Status	: Failed"
            ExtractDocFile(path, text_to_find, new_fail_path)
            dlg = wx.MessageDialog(frame, "All Passed Results Word Document has been extracted completed\n" + "Result saved to this location: \n" + str(new_fail_path),
                                   "Word Document Extraction",
                                   wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

    btn = wx.Button(panel, label="Extract Results")
    btn.SetMinSize((150, 35))  # Minimum width: 150, Height: automatic
    btn.SetMaxSize((200, 10))  # Maximum width: 300, Height: automatic
    btn.SetBackgroundColour("#B6D0E2")
    btn.Bind(wx.EVT_BUTTON, onBtn)
    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(radio, 0, wx.ALL, 10)
    sizer.Add(radio2, 0, wx.ALL, 10)
    sizer.Add(btn, 0, wx.ALL, 5)
    panel.SetSizer(sizer)

    frame.Show()
    app.MainLoop()



def ExtractDocFile(CurrDirectory, text_to_find, newPath):
    for root, dirs, files in os.walk(CurrDirectory):
        if "DOC" in dirs:
            doc_folderPath = os.path.join(root, "DOC")
            if any(file.endswith(".docx") for file in os.listdir(doc_folderPath)):
                for file in os.listdir(doc_folderPath):
                    FilePath = os.path.join(doc_folderPath, file)
                    if file.endswith(".docx") and os.path.isfile(FilePath):
                        doc = Document(FilePath)
                        text_found = False
                        for paragraph in doc.paragraphs:
                            if text_to_find in paragraph.text:
                                shutil.copy(FilePath, newPath)
                                print("Done Copying")
                                text_found = True
                                break  # Exit the loop once text is found
                        if not text_found:
                            print("No Passed Result Found for file:", FilePath)
            else:
                print("No Docs File Found in:", doc_folderPath)






# This section is getting where the python script is being located
# First, sys.argv[0] will get the name of the script being run which is main.py
# Second, os.path.abspath will get the absolute path of where the script is being located or put
# For example, if script is being put into the desktop it will show like this: /Users/<Name of User?>/Desktop/<Script Name>
# Lastly, os.path.dirname will get just only the path of the where the script is located
# For example, if path of the script is /Users/<Name of User>/Desktop/examplescript.py
# It will only extract just the path /Users/<Name of User>/Desktop
DirectoryOfScript = os.path.dirname(os.path.abspath(sys.argv[0]))
print("Directory of the Generate Summary Report Script is here: " + DirectoryOfScript)
#Once it got the proper directory neeed to change it. Must be done this way on Mac, For windows not needed
os.chdir(DirectoryOfScript)





# Set the current directory
startCurrDirectory = DirectoryOfScript
#startCurrDirectory = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\Results\\Folders-123_01"
PromptWindows(startCurrDirectory)




