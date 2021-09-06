Attribute VB_Name = "MemberSheet"

'namespace=source/member

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
    Set tableReference = SheetService.getTable("CNPJA_SOCIOS")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Data de Entrada", _
    "Tipo", _
    "Nome", _
    "CPF / CNPJ", _
    "Faixa Etária", _
    "País M49", _
    "País", _
    "Qualificação ID", _
    "Qualificação", _
    "Representante Nome", _
    "Representante CPF", _
    "Representante Qualificação ID", _
    "Representante Qualificação", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDC64), _
    "CNPJA_SOCIOS", _
    "Sócios e Administradores", _
    columns _
  )

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Data de Entrada").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Tipo").Range
    .ColumnWidth = 15
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Nome").Range
    .ColumnWidth = 35
  End With

  With tableReference.ListColumns("CPF / CNPJ").Range
    .ColumnWidth = 19
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Faixa Etária").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("País M49").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Qualificação ID").Range
    .ColumnWidth = 12
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Qualificação").Range
    .ColumnWidth = 30
  End With

  With tableReference.ListColumns("Representante Nome").Range
    .ColumnWidth = 35
  End With

  With tableReference.ListColumns("Representante CPF").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Representante Qualificação ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Representante Qualificação").Range
    .ColumnWidth = 30
  End With

  tableReference.ListColumns("Última Atualização").Range.ColumnWidth = 19
  Set getTable = tableReference
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim row As Range
  Dim member As Dictionary
  
  SheetService.deleteRowsById tableReference, "Estabelecimento", Response.Data("taxId")
 
  For Each member In Response.Data("company")("members")
    Set row = tableReference.ListRows.Add.Range

    row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
    row(tableReference.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(tableReference.ListColumns("Data de Entrada").Index) = member("since")
    row(tableReference.ListColumns("Nome").Index) = member("person")("name")
    row(tableReference.ListColumns("CPF / CNPJ").Index) = member("person")("taxId")
    row(tableReference.ListColumns("Faixa Etária").Index) = member("person")("age")
    row(tableReference.ListColumns("Qualificação ID").Index) = member("role")("id")
    row(tableReference.ListColumns("Qualificação").Index) = member("role")("text")
    row(tableReference.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

    If member("person")("type") = "NATURAL" Then
      row(tableReference.ListColumns("Tipo").Index) = "Pessoa Física"
    ElseIf member("person")("type") = "LEGAL" Then
      row(tableReference.ListColumns("Tipo").Index) = "Pessoa Jurídica"
    ElseIf member("person")("type") = "FOREIGN" Then
      row(tableReference.ListColumns("Tipo").Index) = "Pessoa Jurídica Estrangeira"
    End If

    If member.Exists("country") Then
      row(tableReference.ListColumns("País M49").Index) = member("country")("id")
      row(tableReference.ListColumns("País").Index) = member("country")("name")
    End If

    If member.Exists("agent") Then
      row(tableReference.ListColumns("Representante Nome").Index) = member("agent")("person")("name")
      row(tableReference.ListColumns("Representante CPF").Index) = member("agent")("person")("taxId")
      row(tableReference.ListColumns("Representante Qualificação ID").Index) = member("agent")("role")("id")
      row(tableReference.ListColumns("Representante Qualificação").Index) = member("agent")("role")("text")
    End If
  Next member
End Function
