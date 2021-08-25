Attribute VB_Name = "QueueController"

'namespace=source/queue

''
' @copyright CNPJA TECNOLOGIA LTDA
' @license CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
''
Option Explicit

''
' Handler of b-queue-action
''
Public Sub startOrPauseQueue(ByRef control As Office.IRibbonControl)
  Dim queueTable As ListObject
  Dim queueProcessing As Long
  Dim queueOpen As Long

  Set queueTable = QueueSheet.getTable()
  queueProcessing = QueueService.countProcessing()
  queueOpen = QueueService.countOpen()
  
  If queueProcessing > 0 Then
    QueueService.pauseRequests
    Exit Sub
  End If

  If queueOpen = 0 Then
    MsgBox "A fila de consultas est� vazia!" & vbCrLf & vbCrLf & _
      "Para adicionar novos itens utilize o bot�o 'Consultar CNPJs'.", _
      vbInformation, "CNPJ�! Fila de Consultas"
    Exit Sub
  End If

  queueTable.ListColumns("Situa��o").Range.Replace "Pausado", "Pendente"
  QueueService.startRequests
End Sub

''
' Getter of b-queue-action image
''
Public Sub getQueueActionImage(ByRef control As Office.IRibbonControl, ByRef image)
  Dim queueProcessing As Long
  queueProcessing = QueueService.countProcessing()

  If queueProcessing > 0 Then
    image = "Media14PausePreview"
  Else
    image = "AnimationStartDropdown"
  End If
End Sub

''
' Getter of b-queue-action label
''
Public Sub getQueueActionLabel(ByRef control As Office.IRibbonControl, ByRef label)
  Dim queueProcessing As Long
  Dim queueOpen As Long

  queueProcessing = QueueService.countProcessing()
  queueOpen = QueueService.countOpen()

  If queueProcessing > 0 Then
    label = "Pausar (" & countOpen & ")"
  Else
    label = "Iniciar (" & countOpen & ")"
  End If
End Sub

''
' Getter of b-queue-error label
''
Public Sub getQueueRetryLabel(ByRef control As Office.IRibbonControl, ByRef label)
  Dim queueError As Long
  queueError = QueueService.countError()
  label = "Reprocessar Falhas (" & queueError & ")"
End Sub

''
' Handler of b-queue-retry
''
Public Sub retryQueue(ByRef control As Office.IRibbonControl)
  Dim queueError As Long
  queueError = QueueService.countError()

  If queueError = 0 Then
    MsgBox "N�o existem falhas na fila de consultas!", vbInformation, "CNPJ�! Fila de Consultas"
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
