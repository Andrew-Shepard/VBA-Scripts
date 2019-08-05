Attribute VB_Name = "copy_word_document"
        Set wb = Documents.Open(ActiveDocument.Path & "\FILENAME.docx")
            Selection.WholeStory
            Selection.Copy
            Windows(1).Activate
            Selection.EndKey Unit:=wdStory, Extend:=wdMove
            Selection.Paste
        wb.Close False
