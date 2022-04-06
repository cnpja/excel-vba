Attribute VB_Name = "SimplesSheet"

'namespace=source/simples

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
    Set tableReference = SheetService.getTable("CNPJA_SIMPLES")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Recibo", _
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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Recibo").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Simples Nacional Optante").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Simples Nacional Inclusão").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("SIMEI Optante").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("SIMEI Inclusão").Range
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
  Dim link As Dictionary

  If Not (Response.Data("company").Exists("simples")) Then Exit Function

  Set row = SheetService.getRow(tableReference, "Estabelecimento", Response.Data("taxId"))

  row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
  row(tableReference.ListColumns("Razão Social").Index) = Response.Data("company")("name")
  row(tableReference.ListColumns("Simples Nacional Optante").Index) = UtilService.booleanToString(Response.Data("company")("simples")("optant"))
  row(tableReference.ListColumns("Simples Nacional Inclusão").Index) = Response.Data("company")("simples")("since")
  row(tableReference.ListColumns("SIMEI Optante").Index) = UtilService.booleanToString(Response.Data("company")("simei")("optant"))
  row(tableReference.ListColumns("SIMEI Inclusão").Index) = Response.Data("company")("simei")("since")
  row(tableReference.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  For Each link In Response.Data("links")
    Select Case link("type")
      Case "SIMPLES_CERTIFICATE"
        UtilService.createLink row(tableReference.ListColumns("Recibo").Index), link("url"), ChrW(&HD83D) & ChrW(&HDCE5) & " PDF"
    End Select
  Next link
End Function
