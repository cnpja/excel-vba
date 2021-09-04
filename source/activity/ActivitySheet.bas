Attribute VB_Name = "ActivitySheet"

'namespace=source/activity

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

  Set getTable = SheetService.getTable("CNPJA_ATIVIDADES")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Principal", _
    "Atividade Econômica ID", _
    "Atividade Econômica", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83C) & ChrW(&HDF3F), _
    "CNPJA_ATIVIDADES", _
    "Atividades Econômicas", _
    columns _
  )

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Principal").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Econômica ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Econômica").Range
    .ColumnWidth = 35
  End With

  getTable.ListColumns("Última Atualização").Range.ColumnWidth = 19
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim table As ListObject
  Dim row As Range
  Dim activity As Dictionary
  
  Set table = getTable()
  SheetService.deleteRowsById table, "Estabelecimento", Response.Data("taxId")
 
  Set row = table.ListRows.Add.Range
  row(table.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
  row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
  row(table.ListColumns("Principal").Index) = "Sim"
  row(table.ListColumns("Atividade Econômica ID").Index) = Response.Data("mainActivity")("id")
  row(table.ListColumns("Atividade Econômica").Index) = Response.Data("mainActivity")("text")
  row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  For Each activity In Response.Data("sideActivities")
    Set row = table.ListRows.Add.Range

    row(table.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("Principal").Index) = "Não"
    row(table.ListColumns("Atividade Econômica ID").Index) = activity("id")
    row(table.ListColumns("Atividade Econômica").Index) = activity("text")
    row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next activity
End Function
