Attribute VB_Name = "EmailSheet"

'namespace=source/email

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

  Set getTable = SheetService.getTable("CNPJA_EMAILS")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Raz�o Social", _
    "Endere�o", _
    "Dom�nio", _
    "�ltima Atualiza��o" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDCE7), _
    "CNPJA_EMAILS", _
    "E-mails", _
    columns _
  )

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Endere�o").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Dom�nio").Range
    .ColumnWidth = 20
  End With

  getTable.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim table As ListObject
  Dim row As Range
  Dim email As Dictionary
  
  Set table = getTable()
  SheetService.deleteRowsById table, "Estabelecimento", Response.Data("taxId")
 
  For Each email In Response.Data("emails")
    Set row = table.ListRows.Add.Range

    row(table.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(table.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("Endere�o").Index) = email("address")
    row(table.ListColumns("Dom�nio").Index) = email("domain")
    row(table.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next email
End Function
