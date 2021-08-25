Attribute VB_Name = "UtilService"

'namespace=source/util

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Open target website in default browser
''
Public Function openUrl(url As String)
  On Error GoTo WorkbookFollow
  Call CreateObject("Shell.Application").ShellExecute(url)
  On Error GoTo 0
  Exit Function
WorkbookFollow:
  ActiveWorkbook.FollowHyperlink (url)
End Function

''
' Builds a hyperlink to target table, and sets the reference to a cell that matches
' lookup with column name
''
Public Function createLink(cell As Range, table As String, column As String, lookupValue As Variant, displayFormula As Variant)
  cell.Formula = _
    "=IFERROR(HYPERLINK(""#""&CELL(""address""," & _
    "INDEX(" & table & "[" & column & "]," & _
    "MATCH(""" & lookupValue & """," & table & "[" & column & "],0)" & _
    "))," & displayFormula & "),0)"

  With cell
    .Font.Name = "Lato"
    .Font.Size = 10.5
    .Font.Bold = True
    .Font.Underline = xlUnderlineStyleNone
    .Font.Color = RGB(0, 161, 96)
  End With
End Function

''
' Shortcut for office link creation by tax id
''
Public Function createTaxIdLink(cell As Range, taxId As Variant)
  createLink cell, "CNPJA_ESTABELECIMENTOS", "Estabelecimento", taxId, """" & taxId & """"
End Function

''
' Shortcut for office link creation by tax id
''
Public Function createCountLink(cell As Range, taxId As Variant, target As String) As String
  createLink cell, target, "Estabelecimento", taxId, "COUNTIF(" & target & "[Estabelecimento],[@Estabelecimento])"
End Function

''
' Extract non-digits from target string
''
Function stringToDigits(inputString As String) As String
  Dim extractedDigits As String
  Dim i As Integer

  extractedDigits = ""

  For i = 1 To Len(inputString)
    If Mid(inputString, i, 1) >= "0" And Mid(inputString, i, 1) <= "9" Then
      extractedDigits = extractedDigits + Mid(inputString, i, 1)
    End If
  Next

  stringToDigits = extractedDigits
End Function

''
' Resets font style of target range
''
Function resetStyle(targetRange As Range)
  With targetRange
    .Font.Name = "Lato"
    .Font.Size = 10.5
    .Font.Underline = xlUnderlineStyleNone
    .Font.Color = RGB(38, 38, 38)
  End With
End Function

''
' Converts a boolean into "Sim" or "Não"
''
Function booleanToString(inputBoolean As Variant) As String
  If inputBoolean Then
    booleanToString = "Sim"
  Else
    booleanToString = "Não"
  End If
End Function
