Attribute VB_Name = "ActivitySheet"

'namespace=source/activity

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

Private tableReference As ListObject

''
' Acquires reference to target table creating if necessary
''
Public Function getTable() As ListObject
  Dim columns() As Variant
  Dim newSheet As Worksheet

  If tableReference Is Nothing Then
    Set tableReference = SheetService.getTable("CNPJA_ATIVIDADES")
  End If

  Set getTable = tableReference
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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Principal").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Atividade Econômica ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Atividade Econômica").Range
    .ColumnWidth = 35
  End With

  tableReference.ListColumns("Última Atualização").Range.ColumnWidth = 19
  Set getTable = tableReference
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim row As Range
  Dim activity As Dictionary

  SheetService.deleteRowsById tableReference, "Estabelecimento", Response.Data("taxId")
 
  Set row = tableReference.ListRows.Add.Range
  row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
  row(tableReference.ListColumns("Razão Social").Index) = Response.Data("company")("name")
  row(tableReference.ListColumns("Principal").Index) = "Sim"
  row(tableReference.ListColumns("Atividade Econômica ID").Index) = Response.Data("mainActivity")("id")
  row(tableReference.ListColumns("Atividade Econômica").Index) = Response.Data("mainActivity")("text")
  row(tableReference.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  For Each activity In Response.Data("sideActivities")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("Principal").Index) = "Não"
    row(tableReference.ListColumns("Atividade Econômica ID").Index) = activity("id")
    row(tableReference.ListColumns("Atividade Econômica").Index) = activity("text")
    row(tableReference.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next activity
End Function
