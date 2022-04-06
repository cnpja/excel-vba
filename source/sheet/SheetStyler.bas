Attribute VB_Name = "SheetStyler"

'namespace=source/sheet

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Apply custom styles to target table
''
Public Function applyTableStyle(table As ListObject)
  Dim styleExist As Boolean

  styleExist = styleExists("CNPJA_TABLE_STYLE")
  If Not styleExist Then createTableStyle

  table.TableStyle = "CNPJA_TABLE_STYLE"
  table.ListRows(1).Delete
  table.Range.Interior.Pattern = xlNone
End Function

''
' Checks if target table style exists
''
Public Function styleExists(styleName As String) As Boolean
  On Error Resume Next
  styleExists = Len(ActiveWorkbook.TableStyles(styleName).Name) > 0
  On Error GoTo 0
End Function

''
' Creates a custom table style available to the workbook
''
Public Function createTableStyle()
  ActiveWorkbook.TableStyles.Add "CNPJA_TABLE_STYLE"

  With ActiveWorkbook.TableStyles("CNPJA_TABLE_STYLE")
    .ShowAsAvailableTableStyle = True

    With .TableStyleElements(xlWholeTable)
      .Font.Color = RGB(38, 38, 38)
    End With

    With .TableStyleElements(xlHeaderRow)
      .Font.Color = RGB(255, 255, 255)
      .Font.Bold = True
    End With

    With .TableStyleElements(xlHeaderRow)
      .Interior.Color = RGB(32, 48, 60)
    End With

    With .TableStyleElements(xlRowStripe1)
      .Borders(xlEdgeTop).Color = RGB(242, 242, 242)
      .Borders(xlEdgeBottom).Color = RGB(242, 242, 242)
      .Borders(xlEdgeLeft).Color = RGB(242, 242, 242)
      .Borders(xlEdgeRight).Color = RGB(242, 242, 242)
      .Borders(xlInsideHorizontal).Color = RGB(242, 242, 242)
    End With

    With .TableStyleElements(xlRowStripe2)
      .Borders(xlEdgeTop).Color = RGB(242, 242, 242)
      .Borders(xlEdgeBottom).Color = RGB(242, 242, 242)
      .Borders(xlEdgeLeft).Color = RGB(242, 242, 242)
      .Borders(xlEdgeRight).Color = RGB(242, 242, 242)
      .Borders(xlInsideHorizontal).Color = RGB(242, 242, 242)
    End With

    With .TableStyleElements(xlRowStripe2)
      .Interior.Color = RGB(242, 242, 242)
    End With
  End With
End Function

''
' Apply custom styles to target worksheet
''
Public Function applySheetStyle(sheet As Worksheet)
  Dim logo As Object

  With ActiveWindow
    .DisplayGridlines = False
    .SplitColumn = 2
    .SplitRow = 2
    .FreezePanes = True
  End With

  With sheet.Cells
    .Font.Name = "Lato"
    .Font.Size = 10.5
    .RowHeight = 20
    .ColumnWidth = 13
    .VerticalAlignment = xlVAlignCenter
    .IndentLevel = 1
  End With

  With sheet.Rows(1)
    .Interior.Color = RGB(28, 43, 55)
    .RowHeight = 40
    .Font.Bold = True
    .Font.Color = RGB(199, 229, 252)
    .Font.Size = 15
    .IndentLevel = 0
  End With

  With sheet.Rows(2)
    .Interior.Color = RGB(32, 48, 60)
    .Font.Bold = True
    .RowHeight = 45
    .HorizontalAlignment = xlHAlignCenter
    .WrapText = True
  End With

  With sheet.columns(1)
    .ColumnWidth = 19
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
  End With

  With sheet.columns(2)
    .ColumnWidth = 35
  End With

  On Error Resume Next
  ThisWorkbook.Sheets(1).Shapes("CNPJA_LOGO").Copy
  waitFor 0.3
  sheet.Paste
  On Error GoTo 0

  Set logo = sheet.Shapes(1)
  logo.Top = 13.5
  logo.Left = 19.5
End Function

''
' Halts for desired amount of seconds
''
Private Function waitFor(seconds As Single)
  Dim elapsed As Single
  elapsed = timer + seconds

  Do While timer < elapsed
    DoEvents
  Loop
End Function
