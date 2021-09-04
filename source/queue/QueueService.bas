Attribute VB_Name = "QueueService"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Count pending items
''
Public Function countPending() As Long
  countPending = 0
  Dim queueTable As ListObject
  Set queueTable = SheetService.getTable("CNPJA_FILA")
  If queueTable Is Nothing Then Exit Function
  countPending = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Pendente")
End Function

''
' Count paused items
''
Public Function countPaused() As Long
  countPaused = 0
  Dim queueTable As ListObject
  Set queueTable = SheetService.getTable("CNPJA_FILA")
  If queueTable Is Nothing Then Exit Function
  countPaused = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Pausado")
End Function

''
' Count processing items
''
Public Function countProcessing() As Long
  countProcessing = 0
  Dim queueTable As ListObject
  Set queueTable = SheetService.getTable("CNPJA_FILA")
  If queueTable Is Nothing Then Exit Function
  countProcessing = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Processando")
End Function

''
' Count processing, pending and paused items
''
Public Function countOpen() As Long
  countOpen = countPending() + countPaused() + countProcessing()
End Function

''
' Count error items
''
Public Function countError() As Long
  countError = 0
  Dim queueTable As ListObject
  Set queueTable = SheetService.getTable("CNPJA_FILA")
  If queueTable Is Nothing Then Exit Function
  countError = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Falha")
End Function

''
' Asks the user to select a range containg data list (1 column, multiple rows)
''
Public Function askInputData(displayName As String) As Range
  Dim inputData As Range
  Dim defaultRange As Range
  Dim prompt As String
  Dim columnCount As Long

  prompt = "Selecione o intervalo contendo a lista de " & displayName & "s a consultar:"

  If TypeName(selection) = "Range" Then
    Set defaultRange = selection
  Else
    Set defaultRange = ActiveCell
  End If

  On Error Resume Next
  Set inputData = Application.InputBox(prompt, "CNPJá! Consulta", defaultRange.Address, Type:=8)
  On Error GoTo 0

  If inputData Is Nothing Then End
  columnCount = inputData.columns.Count

  If columnCount > 1 Then
    MsgBox "O intervalo selecionado deve conter apenas uma coluna! " & vbCrLf & _
      "Selecionadas: " & columnCount & vbCrLf & vbCrLf & _
      "Se a lista que deseja consultar estiver dividida em múltiplas colunas, " & _
      "será necessário que primeiro organize-a em linhas", vbExclamation
    End
  End If

  Set askInputData = inputData
End Function

''
' Pauses all pending requests
''
Public Function pauseRequests()
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable()

  Application.DisplayAlerts = False
  queueTable.ListColumns("Situação").Range.Replace "Pendente", "Pausado"
  queueTable.ListColumns("Situação").Range.Replace "Processando", "Pausado"
  queueTable.ListColumns("Mensagem").Range.Replace "Em andamento, aguarde...", ""
  Application.DisplayAlerts = True

  ConfigService.setKey "QUEUE", "RUNNING", "False"
  RibbonController.refresh
  DoEvents
  End
End Function

''
' Retries failed requests
''
Public Function retryRequests()
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable()
  queueTable.ListColumns("Situação").Range.Replace "Falha", "Pendente"
  RibbonController.refresh
  startRequests
End Function

''
' Trigger fetching procedure for pending items
''
Public Function startRequests()
  Dim maxConcurrency As Integer
  Dim queueTable As ListObject
  Dim queueRow As ListRow
  Dim requestId As Range
  Dim requestQuery As Range
  Dim requestType As Range
  Dim requestStatus As Range
  Dim requestDate As Range
  Dim requestMessage As Range
  Dim taxId As String

  maxConcurrency = CInt(ConfigService.getKey("QUEUE", "CONCURRENCY"))
  If countProcessing() >= maxConcurrency Then Exit Function

  If countPending() = 0 Then
    ConfigService.setKey "QUEUE", "RUNNING", "False"
    Exit Function
  End If

  ConfigService.setKey "QUEUE", "RUNNING", "True"
  Set queueTable = QueueSheet.getTable()

  For Each queueRow In queueTable.ListRows
    Set requestId = queueRow.Range.Cells(1, queueTable.ListColumns("ID").Index)
    Set requestQuery = queueRow.Range.Cells(1, queueTable.ListColumns("Consulta").Index)
    Set requestType = queueRow.Range.Cells(1, queueTable.ListColumns("Tipo").Index)
    Set requestStatus = queueRow.Range.Cells(1, queueTable.ListColumns("Situação").Index)
    Set requestDate = queueRow.Range.Cells(1, queueTable.ListColumns("Horário de Processamento").Index)
    Set requestMessage = queueRow.Range.Cells(1, queueTable.ListColumns("Mensagem").Index)

    If requestStatus.Value = "Pendente" Then
      If countProcessing() >= maxConcurrency Then Exit For

      requestDate.Value = Now
      requestStatus.Formula = "=IF((NOW()-[@[Horário de Processamento]])*1440>1,""Falha"",""Processando"")"
      requestMessage.Formula = "=IF((NOW()-[@[Horário de Processamento]])*1440>1,""Tempo de processamento excedido"",""Em andamento, aguarde..."")"
      UtilService.resetStyle requestMessage

      Select Case requestType.Value
        Case "CNPJ"
          taxId = UtilService.stringToDigits(requestQuery.Value)
          CnpjaService.readOfficeByTaxId requestId.Value, taxId

        Case Else
          requestStatus.Value = "Incorreto"
          requestMessage.Value = "Consulta não suportada"
      End Select
    End If
  Next queueRow

  CnpjaService.readMeCredit
End Function

''
' Fulfills an async request by filling the matching queue item with response data
''
Public Function fulfillRequest(Response As WebResponse, requestIdValue As Long)
  Dim queueTable As ListObject
  Dim queueRow As ListRow
  Dim requestId As Range
  Dim requestType As Range
  Dim requestStatus As Range
  Dim requestCost As Range
  Dim requestMessage As Range

  disableUpdates

  Set queueTable = QueueSheet.getTable()
  
  For Each queueRow In queueTable.ListRows
    Set requestId = queueRow.Range.Cells(1, queueTable.ListColumns("ID").Index)
    Set requestType = queueRow.Range.Cells(1, queueTable.ListColumns("Tipo").Index)
    Set requestStatus = queueRow.Range.Cells(1, queueTable.ListColumns("Situação").Index)
    Set requestCost = queueRow.Range.Cells(1, queueTable.ListColumns("Custo").Index)
    Set requestMessage = queueRow.Range.Cells(1, queueTable.ListColumns("Mensagem").Index)

    If requestId.Value = requestIdValue Then
      requestCost.Value = FindInKeyValues(Response.Headers, "cnpja-request-cost")

      Select Case Response.StatusCode
        Case 200
          requestStatus.Value = "Sucesso"
          requestMessage.Value = ""

          Select Case requestType.Value
            Case "CNPJ"
              OfficeSheet.loadData Response
              MemberSheet.loadData Response
              PhoneSheet.loadData Response
              EmailSheet.loadData Response
              ActivitySheet.loadData Response
              SimplesSheet.loadData Response
              CccSheet.loadData Response
          End Select

        Case 400
          requestStatus.Value = "Incorreto"
          requestMessage.Value = requestType.Value & " inválido"

        Case 401
          requestStatus.Value = "Falha"
          requestMessage.Value = "Falha de autenticação"

        Case 404
          requestStatus.Value = "Incorreto"
          requestMessage.Value = requestType.Value & " inexistente"

        Case 429
          requestStatus.Value = "Falha"
          If InStr(Response.Data("message"), "credits") > 0 Then
            requestMessage.Value = "Créditos insuficientes"
          Else
            requestMessage.Value = "Limite por minuto excedido"
          End If

        Case 500
          requestStatus.Value = "Falha"
          requestMessage.Value = "Um problema inesperado ocorreu"

        Case 503
          requestStatus.Value = "Falha"
          requestMessage.Value = "Plataforma indisponível no momento"

        Case 504
          requestStatus.Value = "Falha"
          requestMessage.Value = "Tempo de processamento excedido"
      End Select

      Exit For
    End If
  Next queueRow

  enableUpdates
  startRequests
End Function

''
' Disables Excel update operations
''
Private Function disableUpdates()
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
End Function

''
' Re-enables Excel update operations
''
Private Function enableUpdates()
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
  DoEvents
End Function
