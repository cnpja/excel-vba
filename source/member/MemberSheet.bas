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
    "Raz�o Social", _
    "Data de Entrada", _
    "Tipo", _
    "Nome", _
    "CPF / CNPJ", _
    "Faixa Et�ria", _
    "Pa�s M49", _
    "Pa�s", _
    "Qualifica��o ID", _
    "Qualifica��o", _
    "Representante Nome", _
    "Representante CPF", _
    "Representante Qualifica��o ID", _
    "Representante Qualifica��o", _
    "�ltima Atualiza��o" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83D) & ChrW(&HDC64), _
    "CNPJA_SOCIOS", _
    "S�cios e Administradores", _
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

  With getTable.ListColumns("Faixa Et�ria").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Pa�s M49").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Qualifica��o ID").Range
    .ColumnWidth = 12
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Qualifica��o").Range
    .ColumnWidth = 30
  End With

  With getTable.ListColumns("Representante Nome").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Representante CPF").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Representante Qualifica��o ID").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Representante Qualifica��o").Range
    .ColumnWidth = 30
  End With

  getTable.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
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
    row(table.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
    row(table.ListColumns("Data de Entrada").Index) = member("since")
    row(table.ListColumns("Nome").Index) = member("person")("name")
    row(table.ListColumns("CPF / CNPJ").Index) = member("person")("taxId")
    row(table.ListColumns("Faixa Et�ria").Index) = member("person")("age")
    row(table.ListColumns("Qualifica��o ID").Index) = member("role")("id")
    row(table.ListColumns("Qualifica��o").Index) = member("role")("text")
    row(table.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))

    If member("person")("type") = "NATURAL" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa F�sica"
    ElseIf member("person")("type") = "LEGAL" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa Jur�dica"
    ElseIf member("person")("type") = "FOREIGN" Then
      row(table.ListColumns("Tipo").Index) = "Pessoa Jur�dica Estrangeira"
    End If

    If member.Exists("country") Then
      row(table.ListColumns("Pa�s M49").Index) = member("country")("id")
      row(table.ListColumns("Pa�s").Index) = member("country")("name")
    End If

    If member.Exists("agent") Then
      row(table.ListColumns("Representante Nome").Index) = member("agent")("person")("name")
      row(table.ListColumns("Representante CPF").Index) = member("agent")("person")("taxId")
      row(table.ListColumns("Representante Qualifica��o ID").Index) = member("agent")("role")("id")
      row(table.ListColumns("Representante Qualifica��o").Index) = member("agent")("role")("text")
    End If
  Next member
End Function
