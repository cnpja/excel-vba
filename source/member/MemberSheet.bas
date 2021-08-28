Attribute VB_Name = "MemberSheet"

'namespace=source/member

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

  Set getTable = SheetService.getTable("CNPJA_SOCIOS")
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

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Data de Entrada").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Tipo").Range
    .ColumnWidth = 15
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Nome").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("CPF / CNPJ").Range
    .ColumnWidth = 19
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Faixa Etária").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("País M49").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Qualificação ID").Range
    .ColumnWidth = 12
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Qualificação").Range
    .ColumnWidth = 30
  End With

  With getTable.ListColumns("Representante Nome").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Representante CPF").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Representante Qualificação ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Representante Qualificação").Range
    .ColumnWidth = 30
  End With

  getTable.ListColumns("Última Atualização").Range.ColumnWidth = 19
End Function

''
' Load API response data into the table
''
Public Function loadData(Response As WebResponse)
  Dim table As ListObject
  Dim row As Range
  Dim member As Dictionary
  
  Set table = getTable()
  SheetService.deleteRowsById table, "Estabelecimento", Response.Data("taxId")
 
  For Each member In Response.Data("company")("members")
    Set row = table.ListRows.Add.Range

    UtilService.createTaxIdLink row(table.ListColumns("Estabelecimento").Index), Response.Data("taxId")
    row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("Data de Entrada").Index) = member("since")
    row(table.ListColumns("Nome").Index) = member("person")("name")
    row(table.ListColumns("CPF / CNPJ").Index) = member("person")("taxId")
    row(table.ListColumns("Faixa Etária").Index) = member("person")("age")
    row(table.ListColumns("Qualificação ID").Index) = member("role")("id")
    row(table.ListColumns("Qualificação").Index) = member("role")("text")
    row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

    If member("person")("type") = "NATURAL" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa Física"
    ElseIf member("person")("type") = "LEGAL" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa Jurídica"
    ElseIf member("person")("type") = "FOREIGN" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa Jurídica Estrangeira"
    End If

    If member.Exists("country") Then
      row(table.ListColumns("País M49").Index) = member("country")("id")
      row(table.ListColumns("País").Index) = member("country")("name")
    End If

    If member.Exists("agent") Then
      row(table.ListColumns("Representante Nome").Index) = member("agent")("person")("name")
      row(table.ListColumns("Representante CPF").Index) = member("agent")("person")("taxId")
      row(table.ListColumns("Representante Qualificação ID").Index) = member("agent")("role")("id")
      row(table.ListColumns("Representante Qualificação").Index) = member("agent")("role")("text")
    End If
  Next member
End Function
