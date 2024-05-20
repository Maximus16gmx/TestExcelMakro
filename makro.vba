Sub ExportiereBereichZuMail(rangeAddress As String, recipient As String)
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim Worksheet As Worksheet
    Dim rangeToSend As Range
    Dim tempFilePath As String
    Dim tempFileName As String
    Dim tempFileFullPath As String

    ' Set the worksheet and range to send
    Set Worksheet = ThisWorkbook.Sheets(1)
    Set rangeToSend = Worksheet.Range(rangeAddress)
    
    ' Create a temporary file path
    tempFilePath = Environ$("temp") & "\"
    tempFileName = "Bestellung.pdf"
    tempFileFullPath = tempFilePath & tempFileName
    
    ' Export the range to PDF
    rangeToSend.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempFileFullPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    ' Create an Outlook application and email
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    ' Configure the email
    With outlookMail
        .To = recipient
        .Subject = "Bestellung"
        .Body = "Hallo Leo," & vbNewLine & vbNewLine & "im Anhang findest du die Bestellung." & vbNewLine & vbNewLine & "Grüße" & vbNewLine & "Packer Team"
        .Attachments.Add tempFileFullPath
        .Display ' Display the email so the user can send it manually
    End With
    
    ' Clean up objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    ' Optional: Delete the temporary PDF file after use
    Kill tempFileFullPath
    
End Sub

Sub ErstelleButton(rangeAddress As String, recipient As String, buttonCaption As String, targetCell As Range)
    Dim btn As Button
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Create the button at the specified cell position
    Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, 100, 30)
    With btn
        .Caption = buttonCaption
        .OnAction = "'ExportiereBereichZuMail """ & rangeAddress & """, """ & recipient & """'"
    End With
End Sub

Sub ErstelleAlleButtons()
    ' Button 1: Bereich A5:C15 an empfaenger1@example.com
    ' Position in Zelle B3
    ErstelleButton "C5:J15", "kamaka99@gmx.de", "Send Bereich 1", ThisWorkbook.Sheets(1).Range("K15")
    
    ErstelleButton "C16:J20", "kamaka99@gmx.de", "Send Bereich 2", ThisWorkbook.Sheets(1).Range("K20")
    
    ' Weitere Buttons können hier hinzugefügt werden
End Sub

