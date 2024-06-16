Sub BlattAnzeigen(BlattName As String)
    ThisWorkbook.Sheets(BlattName).Visible = xlSheetVisible
    ThisWorkbook.Sheets(BlattName).Select
End Sub

Sub ZurueckZuStart()
    Dim BlattName As String
    BlattName = ActiveSheet.Name
    ThisWorkbook.Sheets(BlattName).Visible = xlSheetVeryHidden
    ThisWorkbook.Sheets("Start").Select
End Sub
Sub WeiterZuTabelle1()
    BlattAnzeigen ("Tabelle1")
End Sub

Sub WeiterZuTabelle3()
    BlattAnzeigen ("Tabelle3")
End Sub

Sub WeiterZuTabelle4()
    BlattAnzeigen ("Tabelle4")
End Sub








Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    Dim BlattName As String
    BlattName = Sh.Name
    
    ' Überprüfen, ob das Blatt ein bestimmtes Blatt ist, das nicht ausgeblendet werden soll
    If BlattName <> "Start" And BlattName <> "Start2" Then
        Sh.Visible = xlSheetVeryHidden
    End If
End Sub


    
