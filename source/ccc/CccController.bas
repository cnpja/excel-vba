Attribute VB_Name = "CccController"

'namespace=source/ccc

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Getter of tb-office-ccc
''
Public Sub getCcc(ByRef control As Office.IRibbonControl, ByRef pressed)
  Dim states() As Variant
  Dim state As Variant
  Dim togglePressed As String

  states = Array("ac", "al", "ap", "am", "ba", "ce", "df", "es", "go", "ma", "mg", "mt", "ms", _
    "pa", "pb", "pr", "pe", "pi", "rj", "rn", "rs", "ro", "rr", "sc", "sp", "se", "to")

  For Each state In states
    togglePressed = ConfigService.getKey("RIBBON", "cb-office-ccc-" & state)

    If togglePressed = "True" Then
      pressed = True
      Exit Sub
    End If
  Next state

  pressed = False
End Sub

''
' Setter of tb-office-ccc
''
Public Sub setCcc(ByRef control As Office.IRibbonControl, pressed As Boolean)
  If pressed Then
    setMultipleCccState ("True")
  Else
    setMultipleCccState ("False")
  End If
End Sub

''
' Handler of b-office-ccc-all
''
Public Sub setCccAll(ByRef control As Office.IRibbonControl)
  setMultipleCccState ("True")
End Sub

''
' Handler of b-office-ccc-none
''
Public Sub setCccNone(ByRef control As Office.IRibbonControl)
  setMultipleCccState ("False")
End Sub

''
' Sets the toggle state of all CCC button to desired value
''
Private Function setMultipleCccState(toggleState As String)
  Dim states() As Variant
  Dim state As Variant

  states = Array("ac", "al", "ap", "am", "ba", "ce", "df", "es", "go", "ma", "mg", "mt", "ms", _
    "pa", "pb", "pr", "pe", "pi", "rj", "rn", "rs", "ro", "rr", "sc", "sp", "se", "to")

  For Each state In states
    ConfigService.setKey "RIBBON", "cb-office-ccc-" & state, toggleState
  Next state
End Function
