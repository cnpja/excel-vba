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
' Builds a hyperlink to URL at target cell and apply custom styling
''
Public Function createLink(cell As Range, url As String, label As String)
  With cell
    .Hyperlinks.Add cell, url, , "Abre o arquivo em seu navegador, sujeito a cobran�a de cr�ditos por clique.", label
    .Font.Name = "Lato"
    .Font.Size = 10.5
    .Font.Bold = True
    .Font.Underline = xlUnderlineStyleNone
    .Font.Color = RGB(0, 161, 96)
  End With
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
' Converts a boolean into "Sim" or "N�o"
''
Function booleanToString(inputBoolean As Variant) As String
  If inputBoolean Then
    booleanToString = "Sim"
  Else
    booleanToString = "N�o"
  End If
End Function
