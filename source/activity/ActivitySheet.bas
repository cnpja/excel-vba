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
    "Raz�o Social", _
    "Principal", _
    "Atividade Econ�mica ID", _
    "Atividade Econ�mica", _
    "�ltima Atualiza��o" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83C) & ChrW(&HDF3F), _
    "CNPJA_ATIVIDADES", _
    "Atividades Econ�micas", _
    columns _
  )

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Principal").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Econ�mica ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Econ�mica").Range
    .ColumnWidth = 35
  End With

  getTable.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
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
  row(table.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
  row(table.ListColumns("Principal").Index) = "Sim"
  row(table.ListColumns("Atividade Econ�mica ID").Index) = Response.Data("mainActivity")("id")
  row(table.ListColumns("Atividade Econ�mica").Index) = Response.Data("mainActivity")("text")
  row(table.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  For Each activity In Response.Data("sideActivities")
    Set row = table.ListRows.Add.Range

    row(table.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(table.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("Principal").Index) = "N�o"
    row(table.ListColumns("Atividade Econ�mica ID").Index) = activity("id")
    row(table.ListColumns("Atividade Econ�mica").Index) = activity("text")
    row(table.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next activity
End Function
