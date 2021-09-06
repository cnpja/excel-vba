Attribute VB_Name = "CccSheet"

'namespace=source/ccc

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
    Set tableReference = SheetService.getTable("CNPJA_CCC")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Estado", _
    "Inscrição Estadual", _
    "Habilitada", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83C) & ChrW(&HDF0E), _
    "CNPJA_CCC", _
    "Inscrições Estaduais", _
    columns _
  )

  With newSheet.Cells(1, 3)
    .Value = "  " & ChrW(&H26A0) & "  Requer ativação via menu"
    .Font.Size = 10.5
    .Font.Color = RGB(234, 237, 55)
  End With

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Estado").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Inscrição Estadual").Range
    .ColumnWidth = 19
    .NumberFormat = "@"
  End With

  With tableReference.ListColumns("Habilitada").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  tableReference.ListColumns("Última Atualização").Range.ColumnWidth = 19
  Set getTable = tableReference
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim row As Range
  Dim registration As Dictionary

  If Not (Response.Data.Exists("registrations")) Then Exit Function

  SheetService.deleteRowsById tableReference, "Estabelecimento", Response.Data("taxId")
 
  For Each registration In Response.Data("registrations")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("Estado").Index) = registration("state")
    row(tableReference.ListColumns("Inscrição Estadual").Index) = registration("number")
    row(tableReference.ListColumns("Habilitada").Index) = UtilService.booleanToString(registration("enabled"))
    row(tableReference.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))
  Next registration
End Function
