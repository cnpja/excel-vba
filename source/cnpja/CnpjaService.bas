Attribute VB_Name = "CnpjaService"

'namespace=source/cnpja

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

Private cnpjaClient As WebAsyncWrapper

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
  Request.Resource = "office/" & taxId
  buildOfficeMaxAge Request
  buildOfficeEmbeds Request
  Request.AddQuerystringParam "sync", "true"
  Cnpja.ExecuteAsync Request, "callback", requestId
End Function

''
' Build the max age query param
''
Private Function buildOfficeMaxAge(Request As WebRequest)
  Dim maxAge As String
  maxAge = ConfigService.getKey("OFFICE", "MAX_AGE")
  If maxAge = Empty Then maxAge = "30"
  Request.AddQuerystringParam "maxAge", maxAge
End Function

''
' Build the office embeds query params which may contain CCC and/or SIMPLES
''
Private Function buildOfficeEmbeds(Request As WebRequest)
  Dim embeds As String
  Dim embedSimples As String
  Dim states() As Variant
  Dim state As Variant
  Dim cccPressed As String
  Dim cccStates As String

  embeds = ""
  cccStates = ""

  embedSimples = ConfigService.getKey("RIBBON", "tb-office-simples")
  If embedSimples = "True" Then embeds = embeds & "SIMPLES,"

  states = Array("ac", "al", "ap", "am", "ba", "ce", "df", "es", "go", "ma", "mg", "mt", "ms", _
    "pa", "pb", "pr", "pe", "pi", "rj", "rn", "rs", "ro", "rr", "sc", "sp", "se", "to")

  For Each state In states
    cccPressed = ConfigService.getKey("RIBBON", "cb-office-ccc-" & state)

    If cccPressed = "True" Then
      cccStates = cccStates & UCase(state) & ","
    End If
  Next state

  If cccStates <> "" Then
    cccStates = Left(cccStates, Len(cccStates) - 1)
    Request.AddQuerystringParam "cccStates", cccStates
    embeds = embeds & "CCC,"
  End If

  If embeds <> "" Then
    embeds = Left(embeds, Len(embeds) - 1)
    Request.AddQuerystringParam "embeds", embeds
  End If
End Function
