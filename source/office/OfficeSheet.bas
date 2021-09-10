Attribute VB_Name = "OfficeSheet"

'namespace=source/office

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
    Set tableReference = SheetService.getTable("CNPJA_ESTABELECIMENTOS")
  End If

  Set getTable = tableReference
  If Not getTable Is Nothing Then Exit Function

  columns = Array( _
    "Estabelecimento", _
    "Raz�o Social", _
    "Recibo", _
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
    "Munic�pio IBGE", "Mapa A�reo", "Vis�o da Rua", _
    "Logradouro", "N�mero", "Complemento", "Bairro", "Cidade", "Estado", "CEP", "Pa�s", _
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

  Set tableReference = newSheet.ListObjects(1)

  With tableReference.ListColumns("Recibo").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Mapa A�reo").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Vis�o da Rua").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Porte ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Porte").Range
    .ColumnWidth = 15
  End With

  With tableReference.ListColumns("Capital Social").Range
    .HorizontalAlignment = xlHAlignCenter
    .Style = "Currency"
    .ColumnWidth = 21
  End With

  With tableReference.ListColumns("Natureza Jur�dica ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Natureza Jur�dica").Range
    .ColumnWidth = 30
  End With

  With tableReference.ListColumns("Ente Federativo Respons�vel").Range
    .ColumnWidth = 12
  End With

  With tableReference.ListColumns("S�cios").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Nome Fantasia").Range
  
  End With

  With tableReference.ListColumns("Data de Abertura").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Matriz").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Situa��o ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Situa��o").Range
    .ColumnWidth = 10
  End With

  With tableReference.ListColumns("Situa��o Data").Range
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Telefones").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("E-mails").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Munic�pio IBGE").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Logradouro").Range
    .ColumnWidth = 35
  End With

  With tableReference.ListColumns("N�mero").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Complemento").Range
  
  End With

  With tableReference.ListColumns("Bairro").Range
  
  End With

  With tableReference.ListColumns("Cidade").Range
  
  End With

  With tableReference.ListColumns("Estado").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("CEP").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Pa�s").Range
  
  End With

  With tableReference.ListColumns("Atividade Principal ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Atividade Principal").Range
    .ColumnWidth = 35
  End With

  With tableReference.ListColumns("Atividades Secund�rias").Range
    .ColumnWidth = 11
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Inscri��es Estaduais").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Situa��o Motivo ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Situa��o Motivo").Range
    .ColumnWidth = 20
  End With

  With tableReference.ListColumns("Situa��o Especial ID").Range
    .ColumnWidth = 10
    .HorizontalAlignment = xlHAlignCenter
  End With

  With tableReference.ListColumns("Situa��o Especial").Range
    .ColumnWidth = 20
  End With

  With tableReference.ListColumns("Situa��o Especial Data").Range
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
  Dim link As Dictionary

  Set row = SheetService.getRow(tableReference, "Estabelecimento", Response.Data("taxId"))

  row(tableReference.ListColumns("Estabelecimento").Index) = Response.Data("taxId")
  row(tableReference.ListColumns("Raz�o Social").Index) = Response.Data("company")("name")
  row(tableReference.ListColumns("Porte ID").Index) = Response.Data("company")("size")("id")
  row(tableReference.ListColumns("Porte").Index) = Response.Data("company")("size")("text")
  row(tableReference.ListColumns("Capital Social").Index) = Response.Data("company")("equity")
  row(tableReference.ListColumns("Natureza Jur�dica ID").Index) = Response.Data("company")("nature")("id")
  row(tableReference.ListColumns("Natureza Jur�dica").Index) = Response.Data("company")("nature")("text")
  row(tableReference.ListColumns("Ente Federativo Respons�vel").Index) = Response.Data("company")("jurisdiction")
  row(tableReference.ListColumns("Nome Fantasia").Index) = Response.Data("alias")
  row(tableReference.ListColumns("Data de Abertura").Index) = Response.Data("founded")
  row(tableReference.ListColumns("Matriz").Index) = UtilService.booleanToString(Response.Data("head"))
  row(tableReference.ListColumns("Situa��o ID").Index) = Response.Data("status")("id")
  row(tableReference.ListColumns("Situa��o").Index) = Response.Data("status")("text")
  row(tableReference.ListColumns("Situa��o Data").Index) = Response.Data("statusDate")
  row(tableReference.ListColumns("Munic�pio IBGE").Index) = Response.Data("address")("municipality")
  row(tableReference.ListColumns("Logradouro").Index) = Response.Data("address")("street")
  row(tableReference.ListColumns("N�mero").Index) = Response.Data("address")("number")
  row(tableReference.ListColumns("Complemento").Index) = Response.Data("address")("details")
  row(tableReference.ListColumns("Bairro").Index) = Response.Data("address")("district")
  row(tableReference.ListColumns("Cidade").Index) = Response.Data("address")("city")
  row(tableReference.ListColumns("Estado").Index) = Response.Data("address")("state")
  row(tableReference.ListColumns("CEP").Index) = Response.Data("address")("zip")
  row(tableReference.ListColumns("Pa�s").Index) = Response.Data("address")("country")("name")
  row(tableReference.ListColumns("Atividade Principal ID").Index) = Response.Data("mainActivity")("id")
  row(tableReference.ListColumns("Atividade Principal").Index) = Response.Data("mainActivity")("text")
  row(tableReference.ListColumns("�ltima Atualiza��o").Index) = WebHelpers.ParseIso(Response.Data("updated"))

  If Response.Data.Exists("reason") Then
    row(tableReference.ListColumns("Situa��o Motivo").Index) = Response.Data("reason")("text")
    row(tableReference.ListColumns("Situa��o Motivo ID").Index) = Response.Data("reason")("id")
  End If

  If Response.Data.Exists("special") Then
    row(tableReference.ListColumns("Situa��o Especial").Index) = Response.Data("special")("text")
    row(tableReference.ListColumns("Situa��o Especial ID").Index) = Response.Data("special")("id")
    row(tableReference.ListColumns("Situa��o Especial Data").Index) = Response.Data("specialDate")
  End If

  row(tableReference.ListColumns("S�cios").Index) = Response.Data("company")("members").Count
  row(tableReference.ListColumns("Telefones").Index) = Response.Data("phones").Count
  row(tableReference.ListColumns("E-mails").Index) = Response.Data("emails").Count
  row(tableReference.ListColumns("Atividades Secund�rias").Index) = Response.Data("sideActivities").Count

  If Response.Data.Exists("registrations") Then
    row(tableReference.ListColumns("Inscri��es Estaduais").Index) = Response.Data("registrations").Count
  End If

  For Each link In Response.Data("links")
    Select Case link("type")
      Case "RFB_CERTIFICATE"
        UtilService.createLink row(tableReference.ListColumns("Recibo").Index), link("url"), ChrW(&HD83D) & ChrW(&HDCE5) & " PDF"
      Case "OFFICE_MAP"
        UtilService.createLink row(tableReference.ListColumns("Mapa A�reo").Index), link("url"), ChrW(&HD83D) & ChrW(&HDCE5) & " PNG"
      Case "OFFICE_STREET"
        UtilService.createLink row(tableReference.ListColumns("Vis�o da Rua").Index), link("url"), ChrW(&HD83D) & ChrW(&HDCE5) & " PNG"
    End Select
  Next link
End Function
