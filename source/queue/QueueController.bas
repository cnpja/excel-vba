Attribute VB_Name = "QueueController"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Getter of b-queue-start label
''
Public Sub getQueueStartLabel(ByRef control As Office.IRibbonControl, ByRef label)
  Dim queueOpen As Long
  queueOpen = QueueService.countOpen
  label = "Iniciar (" & queueOpen & ")"
End Sub

''
' Getter of b-queue-start visible
''
Public Sub getQueueStartVisible(ByRef control As Office.IRibbonControl, ByRef visible)
  Dim queueRunning As String
  queueRunning = ConfigService.getKey("QUEUE", "RUNNING")
  If queueRunning = "True" Then visible = False Else visible = True
End Sub

''
' Handler of b-queue-start
''
Public Sub startQueue(ByRef control As Office.IRibbonControl)
  Dim queueTable As ListObject
  Dim queueOpen As Long

  Set queueTable = QueueSheet.getTable()
  queueOpen = QueueService.countOpen()

  If queueOpen = 0 Then
    MsgBox "A fila de consultas está vazia!" & vbCrLf & vbCrLf & _
      "Para adicionar novos itens utilize o botão 'Consultar CNPJs'.", _
      vbInformation, "CNPJá! Fila de Consultas"
    Exit Sub
  End If

  queueTable.ListColumns("Situação").Range.Replace "Pausado", "Pendente"
  QueueService.startRequests
End Sub

''
' Getter of b-queue-pause label
''
Public Sub getQueuePauseLabel(ByRef control As Office.IRibbonControl, ByRef label)
  Dim queueOpen As Long
  queueOpen = QueueService.countOpen
  label = "Pausar (" & queueOpen & ")"
End Sub

''
' Getter of b-queue-pause visible
''
Public Sub getQueuePauseVisible(ByRef control As Office.IRibbonControl, ByRef visible)
  Dim queueRunning As String
  queueRunning = ConfigService.getKey("QUEUE", "RUNNING")
  If queueRunning = "True" Then visible = True Else visible = False
End Sub

''
' Handler of b-queue-pause
''
Public Sub pauseQueue(ByRef control As Office.IRibbonControl)
  QueueService.pauseRequests
End Sub

''
' Getter of b-queue-retry label
''
Public Sub getQueueRetryLabel(ByRef control As Office.IRibbonControl, ByRef label)
  Dim queueError As Long
  queueError = QueueService.countError
  label = "Reprocessar Falhas (" & queueError & ")"
End Sub

''
' Handler of b-queue-retry
''
Public Sub retryQueue(ByRef control As Office.IRibbonControl)
  Dim queueError As Long
  queueError = QueueService.countError()

  If queueError = 0 Then
    MsgBox "Não existem falhas na fila de consultas!", vbInformation, "CNPJá! Fila de Consultas"
    Exit Sub
  End If

  QueueService.retryRequests
End Sub

''
' Getter of cb-queue-concurrency-*
''
Public Sub getQueueConcurrency(ByRef control As Office.IRibbonControl, ByRef pressed)
  Dim concurrency As String

  concurrency = ConfigService.getKey("QUEUE", "CONCURRENCY")

  If concurrency = Empty Then
    concurrency = "01"
    ConfigService.setKey "QUEUE", "CONCURRENCY", concurrency
  End If

  If control.Id = "cb-queue-concurrency-" & concurrency Then
    pressed = True
  Else
    pressed = False
  End If
End Sub

''
' Setter of cb-queue-concurrency-*
''
Public Sub setQueueConcurrency(ByRef control As Office.IRibbonControl, pressed As Boolean)
  ConfigService.setKey "QUEUE", "CONCURRENCY", Right(control.Id, 2)
End Sub
