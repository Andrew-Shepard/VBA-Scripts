Attribute VB_Name = "Module1"
    Set wb = Documents.Open(ActiveDocument.Path & "\FILENAME.docx")
        With ActiveDocument.Sections(1)
            .Footers(wdHeaderFooterPrimary).Range.Copy
        End With
            Windows(1).Activate
        wb.Close False
        
        For y = 1 To ActiveDocument.Sections.Count
            With ActiveDocument.Sections(y)
                .Footers(wdHeaderFooterPrimary).Range.Paste
            End With
    Next
