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
    "Raz�o Social", _
    "Porte ID", "Porte", _
    "Capital Social", _
    "Natureza Jur�dica ID", "Natureza Jur�dica", _
    "Ente Federativo Respons�vel", _
    "S�cios", _
    "Nome Fantasia", _
    "Data de Abertura", _
    "Matriz", _
    "Situa��o ID", "Situa��o", "Situa��o Data", _
    "Telefones", "E-mails", _
    "Munic�pio IBGE", "Logradouro", "N�mero", "Complemento", "Bairro", "Cidade", "Estado", "CEP", "Pa�s", _
    "Atividade Principal ID", "Atividade Principal", "Atividades Secund�rias", _
    "Inscri��es Estaduais", _
    "Situa��o Motivo ID", "Situa��o Motivo", _
    "Situa��o Especial ID", "Situa��o Especial", "Situa��o Especial Data", _
    "�ltima Atualiza��o" _
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

  With getTable.ListColumns("Natureza Jur�dica ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Natureza Jur�dica").Range
    .ColumnWidth = 30
  End With

  With getTable.ListColumns("Ente Federativo Respons�vel").Range
    .ColumnWidth = 12
  End With

  With getTable.ListColumns("S�cios").Range
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

  With getTable.ListColumns("Situa��o ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situa��o").Range
    .ColumnWidth = 10
  End With

  With getTable.ListColumns("Situa��o Data").Range
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

  With getTable.ListColumns("Munic�pio IBGE").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Logradouro").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("N�mero").Range
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

  With getTable.ListColumns("Pa�s").Range
  
  End With

  With getTable.ListColumns("Atividade Principal ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Atividade Principal").Range
    .ColumnWidth = 35
  End With

  With getTable.ListColumns("Atividades Secund�rias").Range
    .ColumnWidth = 11
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Inscri��es Estaduais").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situa��o Motivo ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situa��o Motivo").Range
    .ColumnWidth = 20
  End With

  With getTable.ListColumns("Situa��o Especial ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With getTable.ListColumns("Situa��o Especial").Range
    .ColumnWidth = 20
  End With

  With getTable.ListColumns("Situa��o Especial Data").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  getTable.ListColumns("�ltima Atualiza��o").Range.ColumnWidth = 19
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
  row(table.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
  row(table.ListColumns("Porte ID").Index) = Response.Data("company")("size")("id")
  row(table.ListColumns("Porte").Index) = Response.Data("company")("size")("text")
  row(table.ListColumns("Capital Social").Index) = Response.Data("company")("equity")
  row(table.ListColumns("Natureza Jur�dica ID").Index) = Response.Data("company")("nature")("id")
  row(table.ListColumns("Natureza Jur�dica").Index) = Response.Data("company")("nature")("text")
  row(table.ListColumns("Ente Federativo Respons�vel").Index) = Response.Data("company")("jurisdiction")
  row(table.ListColumns("Nome Fantasia").Index) = Response.Data("alias")
  row(table.ListColumns("Data de Abertura").Index) = Response.Data("founded")
  row(table.ListColumns("Matriz").Index) = UtilService.booleanToString(Response.Data("head"))
  row(table.ListColumns("Situa��o ID").Index) = Response.Data("status")("id")
  row(table.ListColumns("Situa��o").Index) = Response.Data("status")("text")
  row(table.ListColumns("Situa��o Data").Index) = Response.Data("statusDate")
  row(table.ListColumns("Munic�pio IBGE").Index) = Response.Data("address")("municipality")
  row(table.ListColumns("Logradouro").Index) = Response.Data("address")("street")
  row(table.ListColumns("N�mero").Index) = Response.Data("address")("number")
  row(table.ListColumns("Complemento").Index) = Response.Data("address")("details")
  row(table.ListColumns("Bairro").Index) = Response.Data("address")("district")
  row(table.ListColumns("Cidade").Index) = Response.Data("address")("city")
  row(table.ListColumns("Estado").Index) = Response.Data("address")("state")
  row(table.ListColumns("CEP").Index) = Response.Data("address")("zip")
  row(table.ListColumns("Pa�s").Index) = Response.Data("address")("country")("name")
  row(table.ListColumns("Atividade Principal ID").Index) = Response.Data("mainActivity")("id")
  row(table.ListColumns("Atividade Principal").Index) = Response.Data("mainActivity")("text")
  row(table.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  If Response.Data.Exists("reason") Then
    row(table.ListColumns("Situa��o Motivo").Index) = Response.Data("reason")("text")
    row(table.ListColumns("Situa��o Motivo ID").Index) = Response.Data("reason")("id")
  End If

  If Response.Data.Exists("special") Then
    row(table.ListColumns("Situa��o Especial").Index) = Response.Data("special")("text")
    row(table.ListColumns("Situa��o Especial ID").Index) = Response.Data("special")("id")
    row(table.ListColumns("Situa��o Especial Data").Index) = Response.Data("specialDate")
  End If

  row(table.ListColumns("S�cios").Index) = Response.Data("company")("members").Count
  row(table.ListColumns("Telefones").Index) = Response.Data("phones").Count
  row(table.ListColumns("E-mails").Index) = Response.Data("emails").Count
  row(table.ListColumns("Atividades Secund�rias").Index) = Response.Data("sideActivities").Count

  If Response.Data.Exists("registrations") Then
    row(table.ListColumns("Inscri��es Estaduais").Index) = Response.Data("registrations").Count
  End If
End Function
