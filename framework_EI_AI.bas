Attribute VB_Name = "Module1"
Sub RemoveEmptyColumns(ws As Worksheet)
 
    For Column = 25 To 1 Step -1
        ' Vérifie si toutes les cellules de la colonne sont vides
        If WorksheetFunction.CountA(ws.Columns(Column)) = 0 Then
            ' Supprime la colonne
            ws.Columns(Column).Delete
        End If
    Next Column

End Sub

Sub AutoWidthMErgedCells(col As Integer, mergedcount As Integer, ws As Worksheet)


    
    Set SourceRange = ws.Columns(col) ' Copy the entire column A
    
    ' Set the destination range
    Set destinationRange = ws.Range("Z1") ' Replace with your destination cell
    
    ' Copy and paste values only
    SourceRange.Copy
    destinationRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count).PasteSpecial Paste:=xlPasteValues
    ws.Columns("Z").AutoFit
    ws.Columns(col).ColumnWidth = ws.Columns("Z").ColumnWidth
    For i = 1 To mergedcount - 1 Step 1
        ws.Columns(col + i).ColumnWidth = 0
    Next i

    ws.Columns("Z").Delete
    
End Sub

Sub exec(sheet As Worksheet)

    Dim DHCell As Range
    Dim PRTCell As Range
    Dim SMKCell As Range
    Dim DEPCell As Range
    Dim NameCell As Range
    Dim TitleCell As Range
    Dim SNTCell As Range
    
    Dim columnToDelete As Range
    Dim currentCell As Range

    
    RemoveEmptyColumns sheet
    Set TitleCell = sheet.UsedRange.Find(What:="ROOMS", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set PRTCell = sheet.UsedRange.Find(What:="PRT", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set DHCell = sheet.UsedRange.Find(What:="DH", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set SMKCell = sheet.UsedRange.Find(What:="SMK", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set NameCell = sheet.UsedRange.Find(What:="NAME", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set DEPCell = sheet.UsedRange.Find(What:="DEP", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    Set SNTCell = sheet.UsedRange.Find(What:="SNT", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    Set currentCell = NameCell
    Do Until IsEmpty(currentCell)
        ' Process the current row (cell value)
        sheet.Rows(currentCell.Row).RowHeight = 36
        
        Set currentCell = currentCell.Offset(1, 0)
    Loop
        sheet.Rows(currentCell.Row - 1).AutoFit
        sheet.Rows(NameCell.Row).RowHeight = 14.25
    
    sheet.UsedRange.Columns.AutoFit 'Autofit width column
    If Not (DHCell Is Nothing Or DHCell Is Nothing Or SMKCell Is Nothing Or SNTCell Is Nothing Or DEPCell Is Nothing) Then
        Set columnToDelete = sheet.Columns(SMKCell.Column)
        columnToDelete.Delete Shift:=xlToLeft
    
    
    DHCell.Value = "ROOM"
    PRTCell.Value = "SIGNATURE"
    
    AutoWidthMErgedCells DHCell.Column, 1, sheet
    sheet.Columns(DHCell.Column).ColumnWidth = sheet.Columns(DHCell.Column).ColumnWidth + 2
    
    AutoWidthMErgedCells PRTCell.Column, 1, sheet
    sheet.Columns(PRTCell.Column).ColumnWidth = sheet.Columns(PRTCell.Column).ColumnWidth + 2
    
    AutoWidthMErgedCells DEPCell.Column, 1, sheet
    AutoWidthMErgedCells NameCell.Column, 3, sheet
    'AutoWidthMErgedCells SNTCell.Column, 2
    sheet.Range(TitleCell.Offset(0, -1), TitleCell).Merge
    End If

End Sub

Sub Script_Framework()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        exec ws
    Next ws
End Sub





