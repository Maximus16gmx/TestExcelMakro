Sub ExportRangeToPDF(rangeAddress As String, recipient As String)
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim worksheet As Worksheet
    Dim rangeToPrint As Range
    Dim tempFilePath As String
    Dim tempFileName As String
    Dim tempFileFullPath As String
    Dim currentDateTime As String

    On Error GoTo ErrorHandler
    
    ' Set the worksheet and range to print
    Set worksheet = ThisWorkbook.Sheets(1)
    Set rangeToPrint = worksheet.Range(rangeAddress)
    
    ' Create a temporary file path with the current date and time
    currentDateTime = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    tempFilePath = Environ$("temp") & "\"
    tempFileName = "Export_" & currentDateTime & ".pdf"
    tempFileFullPath = tempFilePath & tempFileName
    
    ' Export the range to PDF
    rangeToPrint.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempFileFullPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Create an Outlook application and email
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    ' Configure the email
    With outlookMail
        .To = recipient
        .Subject = "Automated Email with Excel Range"
        .Body = "Hello," & vbCrLf & vbCrLf & _
                "Please find the attached PDF file with the specified range from the Excel file." & vbCrLf & vbCrLf & _
                "Best regards," & vbCrLf & _
                "Your Name"
        .Attachments.Add tempFileFullPath
        .Display ' Display the email so the user can send it manually
    End With
    
    ' Clean up objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing

    ' Optional: Delete the temporary PDF file after use
    ' Kill tempFileFullPath

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Erstellen der Outlook-Anwendung. Fehlernummer: " & Err.Number & vbCrLf & "Beschreibung: " & Err.Description, vbCritical
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub

Sub CreatePrintButton(rangeAddress As String, buttonCaption As String, targetCell As Range)
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Create the button at the specified cell position
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 100, 30)
    With btn
        .Caption = buttonCaption
        .OnAction = "'PrintRange """ & rangeAddress & """'"
    End With
End Sub

Sub PrintRange(rangeAddress As String)
    Dim worksheet As Worksheet
    Dim rangeToPrint As Range
    Dim printerName As String
    Dim printerDialog As Boolean
    
    ' Set the worksheet and range to print
    Set worksheet = ThisWorkbook.Sheets(1)
    Set rangeToPrint = worksheet.Range(rangeAddress)
    
    ' Display the print dialog to select the printer
    printerDialog = Application.Dialogs(xlDialogPrinterSetup).Show
    
    ' Check if the user clicked "OK" or "Cancel"
    If printerDialog Then
        ' Get the selected printer
        printerName = Application.ActivePrinter
        
        ' Print the range
        rangeToPrint.PrintOut ActivePrinter:=printerName
    Else
        MsgBox "Druck abgebrochen", vbInformation
    End If
End Sub

Sub CreateButton(rangeAddress As String, recipient As String, buttonCaption As String, targetCell As Range)
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Create the button at the specified cell position
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 100, 30)
    With btn
        .Caption = buttonCaption
        .OnAction = "'ExportRangeToPDF """ & rangeAddress & """, """ & recipient & """'"
    End With
End Sub

Sub CreateAllButtons()
    ' Button 1: Bereich A1:F15 an empfaenger1@example.com
    ' Position in Zelle B3
    CreateButton "A1:F15", "empfaenger1@example.com", "Send Bereich 1", ThisWorkbook.Sheets(1).Range("B3")
    
    ' Button 2: Bereich G1:H10 an empfaenger2@example.com
    ' Position in Zelle E5
    CreateButton "G1:H10", "empfaenger2@example.com", "Send Bereich 2", ThisWorkbook.Sheets(1).Range("E5")
    
    ' Druck-Button: Bereich A1:F15 drucken
    ' Position in Zelle C3
    CreatePrintButton "A1:F15", "Print Bereich 1", ThisWorkbook.Sheets(1).Range("C3")
    
    ' Weitere Buttons können hier hinzugefügt werden
End Sub
