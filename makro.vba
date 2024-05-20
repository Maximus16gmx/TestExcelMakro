Sub ExportiereBereichZuMail(rangeAddress As String, recipient As String, sheetNumber As Integer)
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim Worksheet As Worksheet
    Dim rangeToPrint As Range
    Dim tempFilePath As String
    Dim tempFileName As String
    Dim tempFileFullPath As String
    Dim currentDateTime As String

    ' Set the worksheet and range to print
    Set Worksheet = ThisWorkbook.Sheets(sheetNumber)
    Set rangeToPrint = Worksheet.Range(rangeAddress)
    
    ' Create a temporary file path with the current date and time
    currentDateTime = Format(Now, "yyyy-mm-dd_hh-nn")
    tempFilePath = Environ$("temp") & "\"
    tempFileName = "Bestellung_" & currentDateTime & ".pdf"
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
End Sub

Sub DruckeBereich(rangeAddress As String, sheetNumber As Integer)
    Dim Worksheet As Worksheet
    Dim rangeToPrint As Range
    Dim printerName As String
    Dim printerDialog As Boolean
    
    ' Set the worksheet and range to print
    Set Worksheet = ThisWorkbook.Sheets(sheetNumber)
    Set rangeToPrint = Worksheet.Range(rangeAddress)
    
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


Sub ErstelleDruckButton(rangeAddress As String, buttonCaption As String, targetCell As Range, sheetNumber As Integer)
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetNumber)
    
    ' Create the button at the specified cell position
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 60, 30)
    With btn
        .Caption = buttonCaption
        .OnAction = "'DruckeBereich """ & rangeAddress & """, """ & sheetNumber & """'"
    End With
End Sub

Sub ErstelleButton(rangeAddress As String, recipient As String, buttonCaption As String, targetCell As Range, sheetNumber As Integer)
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetNumber)
    
    ' Create the button at the specified cell position
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 60, 30)
    With btn
        .Caption = buttonCaption
        .OnAction = "'ExportiereBereichZuMail """ & rangeAddress & """, """ & recipient & """, """ & sheetNumber & """'"
    End With
End Sub

Sub ErstelleAlleButtons()
    ' Button 1: Bereich A5:C15 an empfaenger1@example.com
    ' Position in Zelle B3
    ErstelleButton "C5:J15", "kamaka99@gmx.de", "Senden", ThisWorkbook.Sheets(1).Range("K15"), 1
    ErstelleDruckButton "C5:J15", "Drucken", ThisWorkbook.Sheets(1).Range("K18"), 1
    
    
    ' Weitere Buttons können hier hinzugefügt werden
End Sub
