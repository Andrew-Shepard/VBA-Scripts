Attribute VB_Name = "fill_from_txt"
Dim ReadData As String
Dim myarray() As String
Dim data As String
Open ActiveDocument.Path & "\text.txt" For Input As #1

    Do Until EOF(1)
        Line Input #1, ReadData
        data = data & ReadData
    Loop
    
    If Not Left(data, 1) = "*" Then
        myarray = Split(data, "|")
    End If
Close #1

Dim z As String
i = 1
    For Each f In myarray
        z = CStr(i)
        .Variables(z).Value = f
        i = i + 1
     Next f
     .Fields.Update
     
With ActiveDocument
    strFileName = "EXAMPLE" & myarray(0)
    strPath = .Path
    .SaveAs2 FileName:=strPath & "\" & strFileName, FileFormat:=wdFormatDocumentDefault
    Documents.Open (strPath & "\" & strFileName)
    ActiveDocument.Quit SaveChanges:=wdDoNotSaveChanges
End With
