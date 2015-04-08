Option Explicit
Option Private Module

Sub ExportAllPDFs()
   
    '----------------------------------------------------------------
    'Converts all the PDF files that their paths are in column A of
    'the worksheet "Convert PDF Files" into a different file format,
    'based on the value in column B (extension).
   
    'By Christos Samaras
    'Date: 18/07/2013
    'http://www.myengineeringworld.net
    '----------------------------------------------------------------

    Dim LastRow As Long
    Dim i As Integer
   
    shPaths.Activate
   
    'Find the last row.
    With shPaths
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
   
    'Check that there are available file paths.
    If LastRow < 2 Then
        shPaths.Range("A2").Select
        MsgBox "There are no file paths to convert!", vbInformation, "File paths missing"
        Exit Sub
    End If
       
    'Checking for errors before conversion.
    For i = 2 To LastRow
   
        'Check if the file extensions are not empty.
        If Cells(i, 2).Value = "" Then
            shPaths.Cells(i, 2).Select
            MsgBox "Please select an output format from the dropdown list!", vbCritical, "File paths missing"
            Exit Sub
        End If
       
        'Check if the file exists.
        If Dir(shPaths.Cells(i, 1).Value) = "" Then
            shPaths.Cells(i, 1).Select
            MsgBox "The file path is not valid!", vbCritical, "File path error"
            Exit Sub
        End If
   
        'Check if the input file is a PDF file.
        If LCase(Right(shPaths.Cells(i, 1).Value, 3)) <> "pdf" Then
            shPaths.Cells(i, 1).Select
            MsgBox "The file is not a pdf file!", vbCritical, "No pdf file"
            Exit Sub
        End If
       
    Next i
   
    'For each cell in the range "A2:A" & last row convert the pdf file
    'into different format according to the "B2:B" & last row value.
    For i = 2 To LastRow
        SavePDFAs Cells(i, 1).Value, Cells(i, 2).Value
    Next i
   
    'Adjust the two columns.
    Columns("A:B").EntireColumn.AutoFit
    
    'Inform the user that conversion finished.
    MsgBox "All files were converted successfully!", vbInformation, "Finished"
   
End Sub

Private Sub SavePDFAs(PDFPath As String, FileExtension As String)
   
    '---------------------------------------------------------------------------------------
    'Saves a PDF file as other format using Adobe Professional.
   
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
   
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).
   
    'By Christos Samaras
    'Date: 30/03/2013
    'http://www.myengineeringworld.net
    '---------------------------------------------------------------------------------------
   
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
       
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select
   
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
       
        'Format is correct and no errors.
       
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
       
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
    Else
       
        'Something went wrong, so close the PDF file and the application.
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit

    End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
       
End Sub
