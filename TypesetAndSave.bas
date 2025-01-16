Sub TypesetAndSave()
    Dim newPath As String
    
    If ActiveDocument.ReadOnly Then
        newPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1) & "_formatted.docx"
        ActiveDocument.SaveAs FileName:=newPath
    End If
    
    Selection.WholeStory
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 11
    
    If ActiveDocument.Paragraphs(8).Range.Characters.Count < 4 Then
        ActiveDocument.Paragraphs(8).Range.Delete
    End If
    
    If ActiveDocument.Paragraphs(8).Range.Characters.Count < 4 Then
        ActiveDocument.Paragraphs(8).Range.Delete
    End If
    
    Dim pdfPath As String
    pdfPath = Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1) & ".pdf"
    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True
End Sub
