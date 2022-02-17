Attribute VB_Name = "CnpjaService"

'namespace=source/cnpja

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

Private cnpjaClient As WebAsyncWrapper
Private latestVersion As String

''
' Build the base adapter for requests to CNPJá! API
''
Private Property Get Cnpja() As WebAsyncWrapper
  Dim httpClient As New WebClient

  If cnpjaClient Is Nothing Then
    Set cnpjaClient = New WebAsyncWrapper
    Set httpClient.Authenticator = New CnpjaAuthenticator
    httpClient.BaseUrl = "https://api.cnpja.com/"
    httpClient.TimeoutMs = 59000
    Set cnpjaClient.Client = httpClient
  End If

  Set Cnpja = cnpjaClient
End Property

''
' [Sync] Read self data
''
Public Function readMe() As WebResponse
  Dim Request As New WebRequest
  Request.Resource = "me"
  Set readMe = Cnpja.Client.Execute(Request)
  ConfigService.setKey "ACCOUNT", "NAME", readMe.Data("name")
End Function

''
' [Sync] Read self credits
''
Public Function readMeCredit() As WebResponse
  Dim Request As New WebRequest
  Dim credits As Single

  Request.Resource = "me/credit"
  Set readMeCredit = Cnpja.Client.Execute(Request)

  credits = readMeCredit.Data("perpetual") + readMeCredit.Data("transient")
  ConfigService.setKey "ACCOUNT", "CREDITS", CStr(credits)
End Function

''
' [Sync] Generate self dashboard link
''
Public Function readMeDashboard() As WebResponse
  Dim Request As New WebRequest

  Request.Resource = "me/dashboard"
  Set readMeDashboard = Cnpja.Client.Execute(Request)
End Function

''
' [Sync] Create ticket
''
Public Function createMeTicket(message As String) As WebResponse
  Dim Request As New WebRequest
  Dim Body As Dictionary
  Dim Name As String

  Name = ConfigService.getKey("ACCOUNT", "NAME")
  Set Body = New Dictionary
  Body.Add "subject", "[Excel] Atendimento " & Name
  Body.Add "body", message

  Set Request.Body = Body
  Request.Method = WebMethod.HttpPost
  Request.Resource = "me/ticket"

  Set createMeTicket = Cnpja.Client.Execute(Request)
End Function

''
' [Async] Read an office by its tax id (CNPJ)
''
Public Function readOfficeByTaxId(requestId As Long, taxId As String)
  Dim Request As New WebRequest
  Dim maxAge As String
  Dim simplesPressed As String
  Dim cccPressed As String
  Dim registrations As String
  Dim states() As Variant
  Dim state As Variant

  Request.Resource = "office/" & taxId

  ' Add `maxAge`
  maxAge = ConfigService.getKey("OFFICE", "MAX_AGE")
  If maxAge <> Empty Then Request.AddQuerystringParam "maxAge", maxAge

  ' Add `simples`
  simplesPressed = ConfigService.getKey("RIBBON", "tb-office-simples")
  If simplesPressed = "True" Then Request.AddQuerystringParam "simples", "true"

  ' Add `registrations`
  registrations = ""
  states = Array("ac", "al", "ap", "am", "ba", "ce", "df", "es", "go", "ma", "mg", "mt", "ms", _
    "pa", "pb", "pr", "pe", "pi", "rj", "rn", "rs", "ro", "rr", "sc", "sp", "se", "to")

  For Each state In states
    cccPressed = ConfigService.getKey("RIBBON", "cb-office-ccc-" & state)

    If cccPressed = "True" Then
      registrations = registrations & UCase(state) & ","
    End If
  Next state

  If registrations <> "" Then
    registrations = Left(registrations, Len(registrations) - 1)
    Request.AddQuerystringParam "registrations", registrations
  End If

  ' Add `links`
  Request.AddQuerystringParam "links", "RFB_CERTIFICATE,SIMPLES_CERTIFICATE,OFFICE_MAP,OFFICE_STREET"

  ' Add `sync` in order to acquire request cost after fulfillment
  Request.AddQuerystringParam "sync", "true"

  Cnpja.ExecuteAsync Request, "callback", requestId
End Function

''
' [Sync] Read current tool version
''
Public Function getCurrentVersion() as String
  getCurrentVersion = "1.2.0"
End Function

''
' [Sync] Read latest tool version, cache result to prevent unnecessary requests
''
Public Function getLatestVersion() as String
  Dim req As Object
  Dim url As String

  If latestVersion = "" Then
    url = "https://raw.githubusercontent.com/cnpja/excel-vba/master/.cnpja-version"

    Set req = CreateObject("MSXML2.XMLHTTP")
    req.Open "GET", url, False
    req.Send

    latestVersion = req.ResponseText
  End If

  getLatestVersion = latestVersion
End Function
