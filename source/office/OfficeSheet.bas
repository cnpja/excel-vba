Attribute VB_Name = "OfficeSheet"

'namespace=source/office

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

  Set getTable = SheetService.getTable("CNPJA_ESTABELECIMENTOS")
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Razão Social", _
    "Porte ID", "Porte", _
    "Capital Social", _
    "Natureza Jurídica ID", "Natureza Jurídica", _
    "Ente Federativo Responsável", _
    "Sócios", _
    "Nome Fantasia", _
    "Data de Abertura", _
    "Matriz", _
    "Situação ID", "Situação", "Situação Data", _
    "Telefones", "E-mails", _
    "Município IBGE", "Logradouro", "Número", "Complemento", "Bairro", "Cidade", "Estado", "CEP", "País", _
    "Atividade Principal ID", "Atividade Principal", "Atividades Secundárias", _
    "Inscrições Estaduais", _
    "Situação Motivo ID", "Situação Motivo", _
    "Situação Especial ID", "Situação Especial", "Situação Especial Data", _
    "Última Atualização" _
  )

  Set newSheet = SheetService.createSheet( _
    ChrW(&HD83C) & ChrW(&HDFE6), _
    "CNPJA_ESTABELECIMENTOS", _
    "Estabelecimentos", _
    columns _
  )

  Set getTable = newSheet.ListObjects(1)

  With getTable.ListColumns("Porte ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Porte").Range
    .ColumnWidth = 15
  End With

  With getTable.ListColumns("Capital Social").Range
    .HorizontalAlignment = xlHAlignCenter
    .Style = "Currency"
    .ColumnWidth = 21
  End With

  With getTable.ListColumns("Natureza Jurídica ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Natureza Jurídica").Range
    .ColumnWidth = 30
  End With

  With getTable.ListColumns("Ente Federativo Responsável").Range
    .ColumnWidth = 12
  End With

  With getTable.ListColumns("Sócios").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Nome Fantasia").Range
  
  End With

  With getTable.ListColumns("Data de Abertura").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Matriz").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situação ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situação").Range
    .ColumnWidth = 10
  End With

  With getTable.ListColumns("Situação Data").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Telefones").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("E-mails").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Município IBGE").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Logradouro").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Número").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Complemento").Range
  
  End With

  With getTable.ListColumns("Bairro").Range
  
  End With

  With getTable.ListColumns("Cidade").Range
  
  End With

  With getTable.ListColumns("Estado").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("CEP").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("País").Range
  
  End With

  With getTable.ListColumns("Atividade Principal ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Principal").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Atividades Secundárias").Range
    .ColumnWidth = 11
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Inscrições Estaduais").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situação Motivo ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situação Motivo").Range
    .ColumnWidth = 20
  End With

  With getTable.ListColumns("Situação Especial ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situação Especial").Range
    .ColumnWidth = 20
  End With

  With getTable.ListColumns("Situação Especial Data").Range
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

  Set table = getTable()
  Set row = SheetService.getRow(table, "Estabelecimento", Response.Data("taxId"))

  row(table.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
  row(table.ListColumns("Razão Social").Index) = Response.Data("company")("name")
  row(table.ListColumns("Porte ID").Index) = Response.Data("company")("size")("id")
  row(table.ListColumns("Porte").Index) = Response.Data("company")("size")("text")
  row(table.ListColumns("Capital Social").Index) = Response.Data("company")("equity")
  row(table.ListColumns("Natureza Jurídica ID").Index) = Response.Data("company")("nature")("id")
  row(table.ListColumns("Natureza Jurídica").Index) = Response.Data("company")("nature")("text")
  row(table.ListColumns("Ente Federativo Responsável").Index) = Response.Data("company")("jurisdiction")
  row(table.ListColumns("Nome Fantasia").Index) = Response.Data("alias")
  row(table.ListColumns("Data de Abertura").Index) = Response.Data("founded")
  row(table.ListColumns("Matriz").Index) = UtilService.booleanToString(Response.Data("head"))
  row(table.ListColumns("Situação ID").Index) = Response.Data("status")("id")
  row(table.ListColumns("Situação").Index) = Response.Data("status")("text")
  row(table.ListColumns("Situação Data").Index) = Response.Data("statusDate")
  row(table.ListColumns("Município IBGE").Index) = Response.Data("address")("municipality")
  row(table.ListColumns("Logradouro").Index) = Response.Data("address")("street")
  row(table.ListColumns("Número").Index) = Response.Data("address")("number")
  row(table.ListColumns("Complemento").Index) = Response.Data("address")("details")
  row(table.ListColumns("Bairro").Index) = Response.Data("address")("district")
  row(table.ListColumns("Cidade").Index) = Response.Data("address")("city")
  row(table.ListColumns("Estado").Index) = Response.Data("address")("state")
  row(table.ListColumns("CEP").Index) = Response.Data("address")("zip")
  row(table.ListColumns("País").Index) = Response.Data("address")("country")("name")
  row(table.ListColumns("Atividade Principal ID").Index) = Response.Data("mainActivity")("id")
  row(table.ListColumns("Atividade Principal").Index) = Response.Data("mainActivity")("text")
  row(table.ListColumns("Última Atualização").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  If Response.Data.Exists("reason") Then
    row(table.ListColumns("Situação Motivo").Index) = Response.Data("reason")("text")
    row(table.ListColumns("Situação Motivo ID").Index) = Response.Data("reason")("id")
  End If

  If Response.Data.Exists("special") Then
    row(table.ListColumns("Situação Especial").Index) = Response.Data("special")("text")
    row(table.ListColumns("Situação Especial ID").Index) = Response.Data("special")("id")
    row(table.ListColumns("Situação Especial Data").Index) = Response.Data("specialDate")
  End If

  row(table.ListColumns("Sócios").Index) = Response.Data("company")("members").Count
  row(table.ListColumns("Telefones").Index) = Response.Data("phones").Count
  row(table.ListColumns("E-mails").Index) = Response.Data("emails").Count
  row(table.ListColumns("Atividades Secundárias").Index) = Response.Data("sideActivities").Count

  If Response.Data.Exists("registrations") Then
    row(table.ListColumns("Inscrições Estaduais").Index) = Response.Data("registrations").Count
  End If
End Function
