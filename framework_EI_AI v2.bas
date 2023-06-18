Attribute VB_Name = "Module1"
Sub RemoveEmptyColumns(ws As Worksheet)
    Dim Column As Integer

    For Column = 25 To 1 Step -1
        If WorksheetFunction.CountA(ws.Columns(Column)) = 0 Then
            ws.Columns(Column).Delete
        End If
    Next Column
End Sub

Sub AutoWidthMErgedCells(col As Integer, mergedcount As Integer, ws As Worksheet)
    Dim SourceRange As Range
    Dim DestinationRange As Range
    Dim i As Integer
    
    Set SourceRange = ws.Columns(col)
    Set DestinationRange = ws.Range("Z1")
    
    SourceRange.Copy
    DestinationRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).PasteSpecial Paste:=xlPasteValues
    ws.Columns("Z").AutoFit
    ws.Columns(col).ColumnWidth = ws.Columns("Z").ColumnWidth
    For i = 1 To mergedcount - 1 Step 1
        ws.Columns(col + i).ColumnWidth = 0
    Next i

    ws.Columns("Z").Delete
End Sub

Sub ProcessRows(sheet As Worksheet)
    Dim DHCell As Range
    Dim PRTCell As Range
    Dim SMKCell As Range
    Dim DEPCell As Range
    Dim NameCell As Range
    Dim TitleCell As Range
    Dim currentCell As Range

    RemoveEmptyColumns sheet

    ' Rechercher les cellules nécessaires
    Set PRTCell = sheet.UsedRange.Find(What:="PRT", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set DHCell = sheet.UsedRange.Find(What:="DH", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set SMKCell = sheet.UsedRange.Find(What:="SMK", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set NameCell = sheet.UsedRange.Find(What:="NAME", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set DEPCell = sheet.UsedRange.Find(What:="DEP", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set TitleCell = sheet.UsedRange.Find(What:="ROOMS", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

    ' Traitement des lignes
    Set currentCell = NameCell
    Do Until IsEmpty(currentCell)
        sheet.Rows(currentCell.Row).RowHeight = 36
        Set currentCell = currentCell.Offset(1, 0)
    Loop

    sheet.Rows(currentCell.Row - 1).AutoFit
    sheet.Rows(NameCell.Row).RowHeight = 14.25

    sheet.UsedRange.Columns.AutoFit

    ' Suppression de la colonne SMK
    If Not (DHCell Is Nothing Or PRTCell Is Nothing Or SMKCell Is Nothing Or DEPCell Is Nothing Or TitleCell Is Nothing) Then
        sheet.Columns(SMKCell.Column).Delete Shift:=xlToLeft
        DHCell.Value = "ROOM"
        PRTCell.Value = "SIGNATURE"

        AutoWidthMErgedCells DHCell.Column, 1, sheet
        sheet.Columns(DHCell.Column).ColumnWidth = sheet.Columns(DHCell.Column).ColumnWidth + 2

        AutoWidthMErgedCells PRTCell.Column, 1, sheet
        sheet.Columns(PRTCell.Column).ColumnWidth = sheet.Columns(PRTCell.Column).ColumnWidth + 2

        AutoWidthMErgedCells DEPCell.Column, 1, sheet
        AutoWidthMErgedCells NameCell.Column, 3, sheet

        sheet.Range(TitleCell.Offset(0, -1), TitleCell).Merge
    End If
End Sub


Sub ProcessAllWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProcessRows ws
    Next ws
End Sub

