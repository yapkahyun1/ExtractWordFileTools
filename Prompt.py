import os

import wx



from datetime import datetime


# class MyForm(wx.Frame):
#
#     #----------------------------------------------------------------------
#     def __init__(self):
#         wx.Frame.__init__(self, None, wx.ID_ANY, "Test Case Results Word Docs Extractor")
#         panel = wx.Panel(self, wx.ID_ANY)
#
#         self.radio = wx.RadioButton(panel, label="Passed Results", style = wx.RB_GROUP)
#         self.radio2 = wx.RadioButton(panel, label="Failed Results")
#
#         btn = wx.Button(panel, label="Extract Results")
#         btn.SetMinSize((150, 35))  # Minimum width: 150, Height: automatic
#         btn.SetMaxSize((200, 10))  # Maximum width: 300, Height: automatic
#         btn.SetBackgroundColour("#B6D0E2")
#         btn.Bind(wx.EVT_BUTTON, self.onBtn)
#         sizer = wx.BoxSizer(wx.VERTICAL)
#         sizer.Add(self.radio, 0, wx.ALL, 10)
#         sizer.Add(self.radio2, 0, wx.ALL, 10)
#         sizer.Add(btn, 0, wx.ALL, 5)
#         panel.SetSizer(sizer)
#
#     #----------------------------------------------------------------------
#     def onBtn(self, event, path):
#         currentTimeAndDate = datetime.now()
#         StrDateTime = currentTimeAndDate.strftime("%d%m%Y_%H%M%S")
#
#         NewDirecname = "Passed Results_"+ str(StrDateTime)
#
#         new_path = os.path.join(path, NewDirecname)
#
#
#
#         if self.radio.GetValue() == True:
#             print("First Radio")
#             os.mkdir(new_path)
#         elif self.radio2.GetValue() == True:
#             print("Second Radio")
#             print("Generated")
#         # print( "First radioBtn = ", self.radio.GetValue())
#         # print( "Second radioBtn = ", self.radio2.GetValue())

# # Run the program
#
#
#
#
# app = wx.PySimpleApp()
# frame = MyForm().Show()
# app.MainLoop()
#
#
#

#
#

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
            new_pass_path = os.path.join(path,Pass_Folder_Name)
            os.mkdir(new_pass_path)
            dlg = wx.MessageDialog(frame, "All Passed Results Word Document has been extracted completed",
                                   "Word Document Extraction",
                                   wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

        elif radio2.GetValue() == True:
            Fail_Folder_Name = "Failed Results" + "_" + str(StrDateTime)
            new_fail_path = os.path.join(path,Fail_Folder_Name)
            os.mkdir(new_fail_path)
            dlg = wx.MessageDialog(frame, "All Failed Results Word Document has been extracted completed",
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

old_file = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\Results"
PromptWindows(old_file)











import shutil

import sys

from datetime import datetime
import os
import wx

import subprocess
import platform


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
            new_pass_path = os.path.join(path,Pass_Folder_Name)
            NewPassFolder = os.mkdir(new_pass_path)
            text_search = "Execution Status	: Passed"
            ExtractDocFile(CurrDirectory, text_search, str(NewPassFolder))
            dlg = wx.MessageDialog(frame, "All Passed Results Word Document has been extracted completed",
                                   "Word Document Extraction",
                                   wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()

        elif radio2.GetValue() == True:
            Fail_Folder_Name = "Failed Results" + "_" + str(StrDateTime)
            new_fail_path = os.path.join(path,Fail_Folder_Name)
            os.mkdir(new_fail_path)
            text_search = "Execution Status	: Failed"
            ExtractDocFile(CurrDirectory, text_search)
            dlg = wx.MessageDialog(frame, "All Failed Results Word Document has been extracted completed",
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




def FindWord(path, text_search):
    doc = Document(path)
    for paragraph in doc.paragraphs:
        if text_search in paragraph.text:
            return True
    return False



def ExtractDocFile(CurrDirectory, text_to_find, newPath):
    for root, dirs, files in os.walk(CurrDirectory):

        # It will check if the "HTM" folders exist or not before performing extraction, meaning it will execute the individual test cases of HTML File instead of the summary html reprot files.
        if "DOC" in dirs:
            doc_folderPath = os.path.join(root, "DOC")
            print("Doc Folder Path: " + doc_folderPath)
            # Check if there is any html file exist within the HTM folder and then perform extraction of the html files
            if any(file.endswith(".docx") for file in os.listdir(doc_folderPath)):
                # If HTML File exist, access the html file by building a path to the html first
                for file in os.listdir(doc_folderPath):
                    FilePath = os.path.join(doc_folderPath, file)
                    print("Doc File Path: " + FilePath)
                    # This will check if the html file exists or the file is there
                    if file.endswith(".docx") and os.path.isfile(FilePath):
                        docFilePath = os.path.join(doc_folderPath, file)
                        print("docFilePath: " + docFilePath)
                        if FindWord(docFilePath, text_to_find):
                            shutil.copy(docFilePath, newPath)
                            print("Done Copying")
                        else:
                            print("No Passed Result Found")

            else:
                print("No Docs File Found")





# This section is getting where the python script is being located
# First, sys.argv[0] will get the name of the script being run which is main.py
# Second, os.path.abspath will get the absolute path of where the script is being located or put
# For example, if script is being put into the desktop it will show like this: /Users/<Name of User?>/Desktop/<Script Name>
# Lastly, os.path.dirname will get just only the path of the where the script is located
# For example, if path of the script is /Users/<Name of User>/Desktop/examplescript.py
# It will only extract just the path /Users/<Name of User>/Desktop
DirectoryOfScript = os.path.dirname(os.path.abspath(sys.argv[0]))
print("Directory of the Generate Summary Report Script is here: " + DirectoryOfScript)
# Once it got the proper directory neeed to change it. Must be done this way on Mac, For windows not needed
os.chdir(DirectoryOfScript)

# CurrDirectory = DirectoryOfScript
CurrDirectory = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\Results"
# oldFile = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\Results\\Folders-123_01\\AABAI_Bulk_Limit_01\DOC\\AABAI_Bulk_Limit_01.docx"
newpath = "C:\\Users\\HP\\PycharmProjects\\ExtractWordFile\\New Word Folder"

PromptWindows(CurrDirectory)