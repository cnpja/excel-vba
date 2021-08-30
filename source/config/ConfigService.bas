Attribute VB_Name = "ConfigService"

'namespace=source/config

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Ask the user to update the API key tied to machine and returns the new input
''
Public Function setApiKey() As String
  Dim msg As String
  Dim title As String
  Dim placeholder As String
  Dim readMe As WebResponse

  title = "CNPJá! Autenticação"
  placeholder = "********-****-****-****-************-********-****-****-****-************"

  msg = "Informe sua Chave de API para conectar a nossa plataforma." & vbCrLf & vbCrLf & _
    "Caso ainda não possua, crie gratuitamente em:" & vbCrLf & _
    "www.cnpja.com/me"

  setApiKey = InputBox(msg, title, placeholder)
  If setApiKey = Empty Then End

  ConfigService.setKey "ACCOUNT", "API_KEY", setApiKey
  Set readMe = CnpjaService.readMe()

  If readMe.StatusCode <> 200 Then
    ConfigService.delSection "ACCOUNT"
    QueueService.pauseRequests

    MsgBox "Chave de API inválida!" & vbCrLf & vbCrLf & _
      "Acesse sua conta para visualizar a correta:" & vbCrLf & _
      "www.cnpja.com/me", vbExclamation, "CNPJá! Autenticação"

    End
  End If

  CnpjaService.readMeCredit
End Function

''
' Retrieves a key from settings storage
''
Public Function getKey(section As String, Key As String)
  getKey = GetSetting("CNPJA", section, Key)
End Function

''
' Create or update a key from settings storage
''
Public Function setKey(section As String, Key As String, Value As String)
  SaveSetting "CNPJA", section, Key, Value
  RibbonController.refresh
End Function

''
' Removes a key from settings storage
''
Public Function delKey(section As String, Key As String)
  On Error Resume Next
  DeleteSetting "CNPJA", section, Key
  On Error GoTo 0
End Function

''
' Removes a section from settings storage
''
Public Function delSection(section As String)
  On Error Resume Next
  DeleteSetting "CNPJA", section
  On Error GoTo 0
End Function

''
' Clear all stored settings
''
Public Function resetSettings()
  On Error Resume Next
  DeleteSetting "CNPJA", "ACCOUNT"
  DeleteSetting "CNPJA", "OFFICE"
  DeleteSetting "CNPJA", "QUEUE"
  DeleteSetting "CNPJA", "RIBBON"
  On Error GoTo 0
End Function
