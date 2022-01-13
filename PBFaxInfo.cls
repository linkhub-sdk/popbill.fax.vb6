VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBFaxInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'2017/07/20 added state, title, result
Public state As Integer
Public result As Integer
Public title As String

Public sendState As Integer
Public convState As Integer
Public sendNum As String
Public senderName As String
Public receiveNum As String
Public receiveNumType As String
Public receiveName As String
Public sendPageCnt As Integer
Public successPageCnt As Integer
Public failPageCnt As Integer
Public refundPageCnt As Integer
Public cancelPageCnt As Integer
Public reserveDT As String
Public receiptDT As String
Public sendDT As String
Public resultDT As String
Public sendResult As Integer
Public receiptNum As String
Public requestNum As String
Public chargePageCnt As Integer
Public tiffFileSize As String
Public fileNames As Collection


Public Sub fromDictionary(dic As Variant)
    state = dic.Item("state")
    result = dic.Item("result")
    title = dic.Item("title")

    sendState = dic.Item("sendState")
    convState = dic.Item("convState")
    sendNum = dic.Item("sendNum")
    senderName = dic.Item("senderName")
    receiveNum = dic.Item("receiveNum")
    receiveNumType = dic.Item("receiveNumType")
    receiveName = dic.Item("receiveName")
    sendPageCnt = dic.Item("sendPageCnt")
    successPageCnt = dic.Item("successPageCnt")
    failPageCnt = dic.Item("failPageCnt")
    refundPageCnt = dic.Item("refundPageCnt")
    cancelPageCnt = dic.Item("cancelPageCnt")
    reserveDT = dic.Item("reserveDT")
    sendDT = dic.Item("sendDT")
    resultDT = dic.Item("resultDT")
    receiptDT = dic.Item("receiptDT")
    sendResult = dic.Item("sendResult")
    receiptNum = dic.Item("receiptNum")
    requestNum = dic.Item("requestNum")
    chargePageCnt = dic.Item("chargePageCnt")
    tiffFileSize = dic.Item("tiffFileSize")
    
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("fileNames")
    
    If (tCollection Is Nothing) = False Then
        Set fileNames = New Collection
        Dim T As Variant
                
        For Each T In tCollection
            fileNames.Add T
        Next
    End If
End Sub

