Attribute VB_Name = "EmailSheet"

'namespace=source/email

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

  If tableReference is Nothing Then
    Set tableReference = SheetService.getTable("CNPJA_EMAILS")
  End If

  Set getTable = tableReference
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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Endere�o").Range
    .ColumnWidth = 35
  End With

  With tableReference.ListColumns("Dom�nio").Range
    .ColumnWidth = 20
  End With

  tableReference.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
  Set getTable = tableReference
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim row As Range
  Dim email As Dictionary

  SheetService.deleteRowsById tableReference, "Estabelecimento", Response.Data("taxId")
 
  For Each email In Response.Data("emails")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("Endere�o").Index) = email("address")
    row(tableReference.ListColumns("Dom�nio").Index) = email("domain")
    row(tableReference.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next email
End Function
