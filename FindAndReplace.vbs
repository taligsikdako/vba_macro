Sub MultiFindAndReplace()
    
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim findReplaceRange As Range
    Dim findValue As String
    Dim replaceValue As String
    Dim cell As Range
    Dim findCell As Range
    Dim lastRow As Long
    
    ' Set the worksheet you are working on
    Set ws = ThisWorkbook.Sheets("Data") ' Change Sheet1 to your sheet name
    
    ' Define the range where find and replace values are listed
    ' Assuming find values are in column B and replace values are in column C
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Set findReplaceRange = ws.Range("B5:C" & lastRow)
    
    ' Define the range where you want to perform the search and replace
    Set searchRange = ws.Range("F5:F100") ' Change this range to your target range
    
    ' Loop through each pair of find and replace values
    For Each findCell In findReplaceRange.Columns(1).Cells
        findValue = findCell.Value
        replaceValue = findCell.Offset(0, 1).Value
        
        ' Loop through each cell in the search range
        For Each cell In searchRange
            If cell.Value = findValue Then
                cell.Value = replaceValue
            End If
        Next cell
    Next findCell
    
    MsgBox "Multiple Find and Replace Complete"
    
End Sub