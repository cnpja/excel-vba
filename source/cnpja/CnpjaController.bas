Attribute VB_Name = "CnpjaController"

'namespace=source/cnpja

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Getter of e-mail
''
Public Sub getAccountName(ByRef control As Office.IRibbonControl, ByRef label)
  Dim email As String
  label = ConfigService.getKey("ACCOUNT", "NAME")
  If label = Empty Then label = "Desconectado"
End Sub

''
' Getter of credits
''
Public Sub getAccountCredits(ByRef control As Office.IRibbonControl, ByRef label)
  Dim credits As String
  credits = ConfigService.getKey("ACCOUNT", "CREDITS")
  If credits = Empty Then credits = "0"
  label = Format(credits, "#,##0") & " " & ChrW(&H20AA)
End Sub

''
' Handler of b-account-name
''
Public Sub openMe(ByRef control As Office.IRibbonControl)
  UtilService.openUrl "https://www.cnpja.com/me"
End Sub

''
' Handler of b-account-credits
''
Public Sub openPlans(ByRef control As Office.IRibbonControl)
  UtilService.openUrl "https://www.cnpja.com/plans"
End Sub

''
' Handler of b-account-api-key
''
Public Sub setApiKey(ByRef control As Office.IRibbonControl)
  ConfigService.setApiKey
End Sub

''
' Handler of b-help-docs
''
Public Sub openDocs(ByRef control As Office.IRibbonControl)
  UtilService.openUrl "https://docs.cnpja.com/excel/usage"
End Sub

''
' Handler of b-help-feedback
''
Public Sub openFeedback(ByRef control As Office.IRibbonControl)
  UtilService.openUrl "https://forms.gle/M6EJkPMuHb9Bb5Ej7"
End Sub

''
' Handler of b-help-status
''
Public Sub openStatus(ByRef control As Office.IRibbonControl)
  UtilService.openUrl "https://status.cnpja.com"
End Sub
