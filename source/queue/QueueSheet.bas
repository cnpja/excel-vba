Attribute VB_Name = "QueueSheet"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

Private tableReference As ListObject

''
' Acquires reference to target table creating if necessary
''
Public Function getTable(Optional skipCreation As Boolean) As ListObject
  Dim columns() As Variant
  Dim newSheet As Worksheet

  If tableReference Is Nothing Then
    Set tableReference = SheetService.getTable("CNPJA_FILA")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Or skipCreation Then Exit Function

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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("ID").Range
    .ColumnWidth = 6.3
  End With

  With tableReference.ListColumns("Situação").Range
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

  With tableReference.ListColumns("Tipo").Range
    .ColumnWidth = 7
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Consulta").Range
    .ColumnWidth = 27.3
  End With

  With tableReference.ListColumns("Custo").Range
    .ColumnWidth = 9
    .NumberFormat = "_-" & ChrW(&H20AA) & " * #.0_-;-" & ChrW(&H20AA) & " * #.0_-;_-" & ChrW(&H20AA) & " * ""-""_-;_-@_-"
    .IndentLevel = 0
  End With

  With tableReference.ListColumns("Mensagem").Range
    .ColumnWidth = 40
  End With

  With tableReference.ListColumns("Horário de Processamento").Range
    .ColumnWidth = 19
    .HorizontalAlignment = xlHAlignCenter
  End With

  newSheet.Rows(1).HorizontalAlignment = xlHAlignLeft
  newSheet.Rows(2).HorizontalAlignment = xlHAlignCenter
  Set getTable = tableReference
End Function

''
' Populates the queue table with data to be acquired
''
Public Function loadData(dataType As String, Data As Range)
  Dim query As Range
  Dim newRow As Range
  Dim lastId As Long

  For Each query In Data
    If query.Value <> "" Then
      Set newRow = tableReference.ListRows.Add.Range
      lastId = Application.WorksheetFunction.Max(tableReference.ListColumns("ID").Range)

      newRow(tableReference.ListColumns("Consulta").Index).NumberFormat = "@"

      newRow(tableReference.ListColumns("ID").Index) = lastId + 1
      newRow(tableReference.ListColumns("Tipo").Index) = dataType
      newRow(tableReference.ListColumns("Consulta").Index) = query.Value
      newRow(tableReference.ListColumns("Situação").Index) = "Pendente"
      newRow(tableReference.ListColumns("Custo").Index) = 0
      newRow(tableReference.ListColumns("Mensagem").Index) = ""
    End If
  Next query
End Function
