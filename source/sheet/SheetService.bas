Attribute VB_Name = "SheetService"

'namespace=source/sheet

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Get a reference to target table or return nothing
''
Public Function getTable(tableName As String) As ListObject
  On Error Resume Next
  Set getTable = Application.Range(tableName).ListObject
  On Error GoTo 0
End Function

''
' Creates table at target sheet and with desired columns
''
Public Function createSheet(sheetName As String, tableName As String, title As String, columns() As Variant) As Worksheet
  Dim newTable As ListObject
  Dim dataRange As Range
  Dim activeWb As Workbook
  Dim columTitle As Variant
  Dim columnCounter As Long
  Dim headerRow As Long

  Application.ScreenUpdating = False

  headerRow = 2
  columnCounter = 0

  Set activeWb = ActiveWorkbook
  Set createSheet = activeWb.Sheets.Add(After:=activeWb.Sheets(activeWb.Sheets.Count))

  createSheet.Name = sheetName
  createSheet.Cells(1, 2).Value = "|   " & title

  For Each columTitle In columns
    columnCounter = columnCounter + 1
    createSheet.Cells(headerRow, columnCounter).Value = columTitle
  Next columTitle

  Set dataRange = createSheet.Range(createSheet.Cells(headerRow, 1), createSheet.Cells(headerRow + 1, columnCounter))
  createSheet.ListObjects.Add(xlSrcRange, dataRange, , xlYes).Name = tableName

  Set newTable = createSheet.ListObjects(1)
  SheetStyler.applySheetStyle createSheet
  SheetStyler.applyTableStyle newTable

  Application.ScreenUpdating = True
End Function

''
' Returns the row with matching ID or creates a new one
''
Public Function getRow(table As ListObject, idColumn As String, idValue As String) As Range
  Dim testValues() As Variant
  Dim i As Long

  testValues = table.ListColumns(idColumn).Range.Value

  For i = 1 To table.ListRows.Count + 1
    If testValues(i, 1) = idValue Then
      Set getRow = table.ListRows(i - 1).Range
      Exit Function
    End If
  Next i

  Set getRow = table.ListRows.Add.Range
End Function

''
' Clears all rows that matches an id criteria
''
Public Function deleteRowsById(table As ListObject, idColumn As String, idValue As Variant) As Range
  Dim idIndex As Long

  Application.DisplayAlerts = False

  idIndex = table.ListColumns(idColumn).Index

  With table
    .AutoFilter.ShowAllData
    .Range.AutoFilter Field:=idIndex, Criteria1:=idValue
    On Error Resume Next
    .DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
    On Error GoTo 0
    .AutoFilter.ShowAllData
  End With

  Application.DisplayAlerts = True
End Function
