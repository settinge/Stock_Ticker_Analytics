Attribute VB_Name = "Module1"

Sub all_sheets()
'Prevent Computer Screen from running
  Application.ScreenUpdating = False
  
Sheets.Add.Name = "Combined_Da"
Sheets("Combined_Data").Move Before:=Sheets(1)
Set combined_sheet = Worksheets("Combined_Data")



For Each ws In Worksheets
lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
Next ws

combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
combined_sheet.Columns("A:G").AutoFit
      
  
End Sub





Sub stock()
'Prevent Computer Screen from running
  Application.ScreenUpdating = False

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 17).Value = "Greatest % Increase"
Cells(3, 17).Value = "Greatest % Decrease"
Cells(4, 17).Value = "Greatest Total Value"
Cells(1, 18).Value = "Ticker"
Cells(1, 19).Value = "Value"



Dim ticker_changes As Long
ticker_changes = 0


Dim no_changes As Long
no_changes = 0


Dim vol As Long
vol = 0

Dim result As Integer
result = 2

Dim min_percent As Long
min_percent = 0

Dim max_percent As Long
max_percent = 0

Dim max_vol As Long
max_vol = 0

Dim str_address As String

Dim ticker_min_max As Long
'ticker_min_max = 0


'RowCount = Cells(Rows.Count, “A”).End(xlUp).Row
For i = 2 To 5000
'For j = 1 To 6



If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
ticker_changes = ticker_changes + 1

Cells(ticker_changes + 1, 9).Value = Cells(i, 1).Value
Cells(ticker_changes + 1, 14).Value = Cells(i, 6).Value
Cells(i, 10).Value = Cells(i, 13).Value - Cells(i, 14).Value
vol = vol + Cells(i, 7).Value
Cells(result, 12).Value = vol
result = result + 1
vol = 0
no_changes = 0

Else: no_changes = no_changes + 1
vol = vol + Cells(i, 7).Value

End If



If no_changes = 1 Then
Cells(ticker_changes + 2, 13).Value = Cells(i, 3).Value
End If












'Cells(i, 10).Value = Cells(i, 13).Value - Cells(i, 14).Value


If Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.ColorIndex = 3


Else: Cells(i, 10).Interior.ColorIndex = 4
End If



With Sheets("Combined_Data")




'min_max = WorksheetFunction.Application.Max(Cells(2, 11), Cells(2, 21))

'WorksheetFunction.Application.Max(Cells(2, 11), Cells(2, 21)).Value = ActiveCell.Address




'Ticker = Ticker - 8


'If Cells(i, 11) > Cells(i + 1, 11) Then
'Cells(i, 11).Value = ticker_min_max

'Else: Cells(i + 1, 11).Value = ticker_min_max

'End If

'Cells(2, 19).Value = ticker_min_max






'min_percent = WorksheetFunction.Application.Min(Range("k2:k" & 5000))



'Cells(3, 19).Value = min_percent







'max_percent = WorksheetFunction.Application.Max(Range("k2:k" & 5000))
'Cells(2, 19).Value = max_percent

'max_vol = WorksheetFunction.Application.Max(Range("l2:l" & 5000))
'Cells(4, 19).Value = max_vol


End With



For j = 2 To 5000

On Error Resume Next



Cells(j, 11).Value = ((Cells(j, 13).Value - Cells(j, 14).Value) / Cells(j, 13).Value) * 100




Next j




Next i

'Allow Computer Screen to refresh (not necessary in most cases)
  Application.ScreenUpdating = True
End Sub







Sub find_min_ticker()
    Dim match As Range
    Dim findMe As String
    Dim findOffset As String

    findMe = WorksheetFunction.Application.Min(Range("k2:k" & 5000))


    Set match = Cells.find(findMe)
    findOffset = match.Offset(, -2).Value
    Cells(3, 18).Value = findOffset
    
    MsgBox "The adjacent word to """ & findMe & """ is """ & findOffset & """."
End Sub

Sub find_max_ticker()


    Dim match As Range
    Dim findMe As String
    Dim findOffset As String

    findMe = WorksheetFunction.Application.Max(Range("k2:k" & 5000))


    Set match = Cells.find(findMe)
    findOffset = match.Offset(, -2).Value
    Cells(2, 18).Value = findOffset
    
    MsgBox "The adjacent word to """ & findMe & """ is """ & findOffset & """."
End Sub

Sub find_max_volume()



    Dim match As Range
    Dim findMe As String
    Dim findOffset As String

    findMe = WorksheetFunction.Application.Max(Range("l2:l" & 5000))


    Set match = Cells.find(findMe)
    findOffset = match.Offset(, -3).Value
    Cells(4, 18).Value = findOffset
    
    MsgBox "The adjacent word to """ & findMe & """ is """ & findOffset & """."
End Sub

Sub min_max_values()

min_percent = WorksheetFunction.Application.Min(Range("k2:k" & 5000))

Cells(3, 19).Value = min_percent







max_percent = WorksheetFunction.Application.Max(Range("k2:k" & 5000))
Cells(2, 19).Value = max_percent

max_vol = WorksheetFunction.Application.Max(Range("l2:l" & 5000))
Cells(4, 19).Value = max_vol
End Sub


