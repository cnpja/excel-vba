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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Principal").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Atividade Econ�mica ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Atividade Econ�mica").Range
    .ColumnWidth = 35
  End With

  tableReference.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
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
  row(tableReference.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
  row(tableReference.ListColumns("Principal").Index) = "Sim"
  row(tableReference.ListColumns("Atividade Econ�mica ID").Index) = Response.Data("mainActivity")("id")
  row(tableReference.ListColumns("Atividade Econ�mica").Index) = Response.Data("mainActivity")("text")
  row(tableReference.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  For Each activity In Response.Data("sideActivities")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("Principal").Index) = "N�o"
    row(tableReference.ListColumns("Atividade Econ�mica ID").Index) = activity("id")
    row(tableReference.ListColumns("Atividade Econ�mica").Index) = activity("text")
    row(tableReference.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next activity
End Function
