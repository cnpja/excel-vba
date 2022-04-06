Attribute VB_Name = "RibbonController"

'namespace=source/ribbon

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

Private cnpjaRibbon As IRibbonUI

#If VBA7 Then
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

''
' Reacquires ribbon by using memmory pointer in case of failure
''
#If VBA7 Then
Function getRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function getRibbon(ByVal lRibbonPointer As Long) As Object
#End If
  Dim objRibbon As Object
  CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
  Set getRibbon = objRibbon
  Set objRibbon = Nothing
End Function

''
' Hook when the ribbon loads to save a reference to ribbon control
''
Public Sub onLoad(ribbon As IRibbonUI)
  Set cnpjaRibbon = ribbon
  ConfigService.setKey "RIBBON", "POINTER", ObjPtr(ribbon)
End Sub

''
' Validates that the ribbon object still available, if not reacquire it
''
Private Sub validate()
  Dim ribbonPointer As Variant

  If cnpjaRibbon Is Nothing Then
    ribbonPointer = ConfigService.getKey("RIBBON", "POINTER")
    Set cnpjaRibbon = getRibbon(ribbonPointer)
  End If
End Sub

''
' Trigger a refresh of elements state
''
Public Sub refresh()
  validate
  cnpjaRibbon.Invalidate
End Sub

''
' Activates the ribbon tab
''
Public Sub activate()
  validate
  On Error Resume Next
  cnpjaRibbon.ActivateTab "t-cnpja"
  On Error GoTo 0
End Sub

''
' Generic state getter
''
Public Sub getToggle(ByRef control As Office.IRibbonControl, ByRef pressed)
  Dim togglePressed As String
  togglePressed = ConfigService.getKey("RIBBON", control.Id)

  If togglePressed = "True" Then
    pressed = True
  Else
    pressed = False
  End If
End Sub

''
' Generic state setter
''
Public Sub setToggle(ByRef control As Office.IRibbonControl, pressed As Boolean)
  If pressed Then
    ConfigService.setKey "RIBBON", control.Id, "True"
  Else
    ConfigService.setKey "RIBBON", control.Id, "False"
  End If
End Sub
