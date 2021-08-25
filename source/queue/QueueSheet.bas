Attribute VB_Name = "QueueSheet"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Acquires reference to target table creating if necessary
''
Public Function getTable() As ListObject
  Dim columns() As Variant
  Dim newSheet As Worksheet

  Set getTable = SheetService.getTable("CNPJA_FILA")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "ID", _
    "Situação", _
    "Tipo", _
    "Consulta", _
    "Custo", _
    "Mensagem", _
    "Horário de Processamento" _
  )

  Set newSheet = SheetService.createSheet( _
    "CNPJá!", _
    "CNPJA_FILA", _
    "Fila de Consultas", _
    columns _
  )

  newSheet.Tab.Color = 6332672
  newSheet.Cells(1, 3).Value = newSheet.Cells(1, 2).Value
  newSheet.Cells(1, 2).Value = ""

  With ActiveWindow
    .FreezePanes = False
    .SplitColumn = 4
    .SplitRow = 2
    .FreezePanes = True
  End With

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("ID").Range
    .ColumnWidth = 6.3
  End With

  With getTable.ListColumns("Situação").Range
    .Font.Bold = True
    .ColumnWidth = 12
    .HorizontalAlignment = xlHAlignCenter
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Pendente"""
    .FormatConditions(1).Font.Color = RGB(160, 160, 160)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Processando"""
    .FormatConditions(2).Font.Color = RGB(230, 180, 0)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Pausado"""
    .FormatConditions(3).Font.Color = RGB(70, 115, 195)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Sucesso"""
    .FormatConditions(4).Font.Color = RGB(0, 160, 95)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Incorreto"""
    .FormatConditions(5).Font.Color = RGB(245, 150, 50)
    .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Falha"""
    .FormatConditions(6).Font.Color = RGB(255, 50, 50)
  End With

  With getTable.ListColumns("Tipo").Range
    .ColumnWidth = 7
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Consulta").Range
    .ColumnWidth = 27.3
  End With

  With getTable.ListColumns("Custo").Range
    .ColumnWidth = 9
    .NumberFormat = "_-" & ChrW(&H20AA) & " * #.0_-;-" & ChrW(&H20AA) & " * #.0_-;_-" & ChrW(&H20AA) & " * ""-""_-;_-@_-"
    .IndentLevel = 0
  End With

  With getTable.ListColumns("Mensagem").Range
    .ColumnWidth = 40
  End With

  With getTable.ListColumns("Horário de Processamento").Range
    .ColumnWidth = 19
    .HorizontalAlignment = xlHAlignCenter
  End With

  newSheet.Rows(1).HorizontalAlignment = xlHAlignLeft
  newSheet.Rows(2).HorizontalAlignment = xlHAlignCenter
End Function

''
' Populates the queue table with data to be acquired
''
Public Function loadData(dataType As String, Data As Range)
  Dim table As ListObject
  Dim query As Range
  Dim newRow As Range
  Dim lastId As Long
  Dim firstRow As Boolean

  setupEnvironment
  Set table = getTable()
  firstRow = True

  For Each query In Data
    If query.Value <> "" Then
      Set newRow = table.ListRows.Add.Range
      lastId = Application.WorksheetFunction.Max(table.ListColumns("ID").Range)

      If firstRow Then
        Application.GoTo newRow
        firstRow = False
      End If

      newRow(table.ListColumns("Consulta").Index).NumberFormat = "@"

      newRow(table.ListColumns("ID").Index) = lastId + 1
      newRow(table.ListColumns("Tipo").Index) = dataType
      newRow(table.ListColumns("Consulta").Index) = query.Value
      newRow(table.ListColumns("Situação").Index) = "Pendente"
      newRow(table.ListColumns("Custo").Index) = 0
      newRow(table.ListColumns("Mensagem").Index) = ""
    End If
  Next query
End Function

''
' Ensures everyting required to call the API is in order:
' - Disable formula auto fill
' - Disable number as text and inconsistent formula warnings
' - Ask user his API key if not configured yet
' - Create tables if not present
''
Private Function setupEnvironment()
  Dim table As ListObject

  Application.AutoCorrect.AutoFillFormulasInLists = False
  Application.ErrorCheckingOptions.InconsistentTableFormula = False
  Application.ErrorCheckingOptions.NumberAsText = False

  Application.ScreenUpdating = False

  Set table = getTable()

  OfficeSheet.getTable
  MemberSheet.getTable
  PhoneSheet.getTable
  EmailSheet.getTable
  ActivitySheet.getTable
  SimplesSheet.getTable
  CccSheet.getTable

  Application.GoTo table.Range
  Application.ScreenUpdating = True

  RibbonController.activate
  CnpjaService.readMe
End Function
