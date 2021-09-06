Attribute VB_Name = "QueueService"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

#If Win64 Then
  Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private appHwnd As LongPtr
#Else
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
  Private appHwnd As Long
#End If

''
' Count pending items
''
Public Function countPending() As Long
  SetForegroundWindow(appHwnd)
  countPending = 0
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable(True)
  If queueTable Is Nothing Then Exit Function
  countPending = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Pendente")
End Function

''
' Count paused items
''
Public Function countPaused() As Long
  SetForegroundWindow(appHwnd)
  countPaused = 0
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable(True)
  If queueTable Is Nothing Then Exit Function
  countPaused = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Pausado")
End Function

''
' Count processing items
''
Public Function countProcessing() As Long
  SetForegroundWindow(appHwnd)
  countProcessing = 0
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable(True)
  If queueTable Is Nothing Then Exit Function
  countProcessing = Application.WorksheetFunction.CountIf(queueTable.ListColumns("Situação").Range, "Processando")
End Function

''
' Count error items
''
Public Function countError() As Long
  SetForegroundWindow(appHwnd)
  countError = 0
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable(True)
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
' Ensures everyting required to call the API is in order:
' - Disable formula auto fill
' - Disable number as text and inconsistent formula warnings
' - Ask user his API key if not configured yet
' - Create tables if not present
''
Public Function setupEnvironment()
  appHwnd = Application.hwnd
  ConfigService.delKey "QUEUE", "FULFILLING"

  Application.AutoCorrect.AutoFillFormulasInLists = False
  Application.ErrorCheckingOptions.InconsistentTableFormula = False
  Application.ErrorCheckingOptions.NumberAsText = False
  Application.ScreenUpdating = False

  QueueSheet.getTable
  OfficeSheet.getTable
  MemberSheet.getTable
  PhoneSheet.getTable
  EmailSheet.getTable
  ActivitySheet.getTable
  SimplesSheet.getTable
  CccSheet.getTable

  Application.ScreenUpdating = True

  RibbonController.activate
  CnpjaService.readMe
End Function

''
' Begin the procered of consuming request queue
''
Public Function startRequests()
  Dim queueTable As ListObject
  
  setupEnvironment
  Set queueTable = QueueSheet.getTable()
  queueTable.ListColumns("Situação").Range.Replace "Pausado", "Pendente"

  ConfigService.setKey "QUEUE", "RUNNING", "True"
  ConfigService.delKey "QUEUE", "FULFILLING"

  CnpjaService.readMeCredit
  processRequestsWithHealthCheck
End Function

''
' Pauses all pending requests
''
Public Function pauseRequests()
  SetForegroundWindow(appHwnd)
  Dim queueTable As ListObject
  Set queueTable = QueueSheet.getTable()

  ConfigService.delKey "QUEUE", "RUNNING"

  Application.DisplayAlerts = False
  queueTable.ListColumns("Situação").Range.Replace "Pendente", "Pausado"
  queueTable.ListColumns("Situação").Range.Replace "Processando", "Pausado"
  queueTable.ListColumns("Mensagem").Range.Replace "Em andamento, aguarde...", ""
  Application.DisplayAlerts = True

  CnpjaService.readMeCredit
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
  startRequests
End Function

''
' Runs finalizing operations after all requests been processed
''
Public Function completeRequests()
  ConfigService.delKey "QUEUE", "RUNNING"
  CnpjaService.readMeCredit
End Function

''
' Schedules a procedure that periodically checks for failed requests
''
Public Function processRequestsWithHealthCheck()
  SetForegroundWindow(appHwnd)
  Application.Calculate

  If (countPending() + countProcessing()) = 0 Then
    completeRequests
  ElseIf ConfigService.getKey("QUEUE", "RUNNING") = "True" Then
    processRequests
    Application.onTime Now + TimeValue("00:00:15"), "QueueService.processRequestsWithHealthCheck"
  End If
End Function

''
' Trigger fetching procedure for pending items
''
Public Function processRequests()
  SetForegroundWindow(appHwnd)
  Dim concurrency As String
  Dim maxConcurrency As Integer
  Dim queueTable As ListObject
  Dim queueRowId As Long
  Dim queueRow As ListRow
  Dim requestId As Range
  Dim requestQuery As Range
  Dim requestType As Range
  Dim requestStatus As Range
  Dim requestDate As Range
  Dim requestMessage As Range
  Dim taxId As String

  concurrency = ConfigService.getKey("QUEUE", "CONCURRENCY")
  If concurrency = "" Then maxConcurrency = 1 Else maxConcurrency = CInt(concurrency)

  Set queueTable = QueueSheet.getTable()

  Do While countProcessing() < maxConcurrency
    Set requestStatus = queueTable.ListColumns("Situação").Range.Find("Pendente", LookIn:=xlValues)
    If requestStatus Is Nothing Then Exit Do

    queueRowId = queueTable.ListRows(requestStatus.Row - queueTable.HeaderRowRange.Row).Index
    Set queueRow = queueTable.ListRows(queueRowId)

    Set requestId = queueRow.Range.Cells(1, queueTable.ListColumns("ID").Index)
    Set requestQuery = queueRow.Range.Cells(1, queueTable.ListColumns("Consulta").Index)
    Set requestType = queueRow.Range.Cells(1, queueTable.ListColumns("Tipo").Index)
    Set requestDate = queueRow.Range.Cells(1, queueTable.ListColumns("Horário de Processamento").Index)
    Set requestMessage = queueRow.Range.Cells(1, queueTable.ListColumns("Mensagem").Index)

    requestDate.Value = Now
    requestStatus.Formula = "=IF((NOW()-[@[Horário de Processamento]])*1440>0.75,""Falha"",""Processando"")"
    requestMessage.Formula = "=IF((NOW()-[@[Horário de Processamento]])*1440>0.75,""Tempo de processamento excedido"",""Em andamento, aguarde..."")"
    UtilService.resetStyle requestMessage

    Select Case requestType.Value
      Case "CNPJ"
        taxId = UtilService.stringToDigits(requestQuery.Value)
        CnpjaService.readOfficeByTaxId requestId.Value, taxId

      Case Else
        requestStatus.Value = "Incorreto"
        requestMessage.Value = "Consulta não suportada"
    End Select
  Loop
End Function

''
' Fulfills an async request by filling the matching queue item with response data
''
Public Function fulfillRequest(Response As WebResponse, requestIdValue As Long)
  SetForegroundWindow(appHwnd)
  Dim queueTable As ListObject
  Dim queueRowId As Long
  Dim queueRow As ListRow
  Dim requestId As Range
  Dim requestType As Range
  Dim requestStatus As Range
  Dim requestCost As Range
  Dim requestMessage As Range
  Dim isFulfilling As String

  On Error GoTo FulfillmentFailure

  Do
    DoEvents
    isFulfilling = ConfigService.getKey("QUEUE", "FULFILLING")
  Loop While isFulfilling = "True"

  ConfigService.setKey "QUEUE", "FULFILLING", "True"
  disableUpdates

  Set queueTable = QueueSheet.getTable()

  Set requestId = queueTable.ListColumns("ID").Range.Find(requestIdValue, LookIn:=xlValues)
  queueRowId = queueTable.ListRows(requestId.Row - queueTable.HeaderRowRange.Row).Index
  Set queueRow = queueTable.ListRows(queueRowId)

  Set requestType = queueRow.Range.Cells(1, queueTable.ListColumns("Tipo").Index)
  Set requestStatus = queueRow.Range.Cells(1, queueTable.ListColumns("Situação").Index)
  Set requestCost = queueRow.Range.Cells(1, queueTable.ListColumns("Custo").Index)
  Set requestMessage = queueRow.Range.Cells(1, queueTable.ListColumns("Mensagem").Index)

  Application.GoTo requestId
  requestCost.Value = FindInKeyValues(Response.Headers, "cnpja-request-cost")

  Select Case Response.StatusCode
    Case 200
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
      requestStatus.Value = "Sucesso"
      requestMessage.Value = ""

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

  enableUpdates
  ConfigService.delKey "QUEUE", "FULFILLING"

  If (countPending() + countProcessing()) = 0 Then
    completeRequests
  Else
    processRequests
  End If

  Exit Function

FulfillmentFailure:
  ConfigService.delKey "QUEUE", "FULFILLING"
  Debug.Print "Fulfillment failed for id: " & requestIdValue
End Function

''
' Disables Excel update operations
''
Private Function disableUpdates()
  SetForegroundWindow(appHwnd)
  Application.ScreenUpdating = False
  Application.EnableEvents = False
End Function

''
' Re-enables Excel update operations
''
Private Function enableUpdates()
  SetForegroundWindow(appHwnd)
  Application.ScreenUpdating = True
  Application.EnableEvents = True
End Function
