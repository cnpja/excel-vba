Attribute VB_Name = "SimplesSheet"

'namespace=source/simples

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

  Set getTable = SheetService.getTable("CNPJA_SIMPLES")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Simples Nacional Optante", _
    "Simples Nacional Inclusão", _
    "SIMEI Optante", _
    "SIMEI Inclusão", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDCB0), _
    "CNPJA_SIMPLES", _
    "Simples Nacional", _
    columns _
  )

  With newSheet.Cells(1, 3)
    .Value = "  " & ChrW(&H26A0) & "  Requer ativação via menu"
    .Font.Size = 10.5
    .Font.Color = RGB(234, 237, 55)
  End With

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Simples Nacional Optante").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Simples Nacional Inclusão").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("SIMEI Optante").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("SIMEI Inclusão").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  getTable.ListColumns("Última Atualização").Range.ColumnWidth = 19
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim table As ListObject
  Dim row As Range

  If Not (Response.Data("company").Exists("simples")) Then Exit Function

  Set table = getTable()
  Set row = SheetService.getRow(table, "Estabelecimento", Response.Data("taxId"))

  UtilService.createTaxIdLink row(table.ListColumns("Estabelecimento").Index), Response.Data("taxId")
  row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
  row(table.ListColumns("Simples Nacional Optante").Index) = UtilService.booleanToString(Response.Data("company")("simples")("optant"))
  row(table.ListColumns("Simples Nacional Inclusão").Index) = Response.Data("company")("simples")("since")
  row(table.ListColumns("SIMEI Optante").Index) = UtilService.booleanToString(Response.Data("company")("simei")("optant"))
  row(table.ListColumns("SIMEI Inclusão").Index) = Response.Data("company")("simei")("since")
  row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))
End Function
