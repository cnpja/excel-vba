Attribute VB_Name = "OfficeController"

'namespace=source/office

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Handler of b-office-query
''
Public Sub queryOffice(ByRef control As Office.IRibbonControl)
  Dim inputData As Range
  Set inputData = QueueService.askInputData("CNPJ")
  QueueService.setupEnvironment
  QueueSheet.loadData "CNPJ", inputData
  QueueService.startRequests
End Sub

''
' Getter of cb-office-max-age-*
''
Public Sub getOfficeMaxAge(ByRef control As Office.IRibbonControl, ByRef pressed)
  Dim maxAge As String

  maxAge = ConfigService.getKey("OFFICE", "MAX_AGE")

  If maxAge = Empty Then
    maxAge = "030"
    ConfigService.setKey "OFFICE", "MAX_AGE", maxAge
  End If

  If control.Id = "cb-office-max-age-" & maxAge Then
    pressed = True
  Else
    pressed = False
  End If
End Sub

''
' Setter of cb-office-max-age-*
''
Public Sub setOfficeMaxAge(ByRef control As Office.IRibbonControl, pressed As Boolean)
  ConfigService.setKey "OFFICE", "MAX_AGE", Right(control.Id, 3)
End Sub
