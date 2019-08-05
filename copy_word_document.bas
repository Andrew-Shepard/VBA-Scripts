Attribute VB_Name = "Module1"
        Set wb = Documents.Open(ActiveDocument.Path & "\FILENAME.docx")
            Selection.WholeStory
            Selection.Copy
            Windows(1).Activate
            Selection.EndKey Unit:=wdStory, Extend:=wdMove
            Selection.Paste
        wb.Close False
