Attribute VB_Name = "PhoneSheet"

'namespace=source/phone

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

  Set getTable = SheetService.getTable("CNPJA_TELEFONES")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "DDD", _
    "Número", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDCDE), _
    "CNPJA_TELEFONES", _
    "Telefones", _
    columns _
  )

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("DDD").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Número").Range
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
  Dim phone As Dictionary
  
  Set table = getTable()
  SheetService.deleteRowsById table, "Estabelecimento", Response.Data("taxId")
 
  For Each phone In Response.Data("phones")
    Set row = table.ListRows.Add.Range

    UtilService.createTaxIdLink row(table.ListColumns("Estabelecimento").Index), Response.Data("taxId")
    row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("DDD").Index) = phone("area")
    row(table.ListColumns("Número").Index) = phone("number")
    row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next phone
End Function
