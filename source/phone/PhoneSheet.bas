Attribute VB_Name = "PhoneSheet"

'namespace=source/phone

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
    Set tableReference = SheetService.getTable("CNPJA_TELEFONES")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Raz�o Social", _
    "DDD", _
    "N�mero", _
    "�ltima Atualiza��o" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDCDE), _
    "CNPJA_TELEFONES", _
    "Telefones", _
    columns _
  )

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("DDD").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("N�mero").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  tableReference.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
  Set getTable = tableReference
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim row As Range
  Dim phone As Dictionary
  
  SheetService.deleteRowsById tableReference, "Estabelecimento", Response.Data("taxId")
 
  For Each phone In Response.Data("phones")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("DDD").Index) = phone("area")
    row(tableReference.ListColumns("N�mero").Index) = phone("number")
    row(tableReference.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next phone
End Function
