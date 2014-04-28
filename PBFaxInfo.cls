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


Public sendState As Integer
Public convState As Integer
Public sendNum As String
Public receiveNum As String
Public receiveName As String
Public sendPageCnt As Integer
Public successPageCnt As Integer
Public failPageCnt As Integer
Public refundPageCnt As Integer
Public cancelPageCnt As Integer
Public reserveDT As String
Public sendDT As String
Public resultDT As String
Public sendResult As Integer

Public Sub fromDictionary(dic As Variant)

    sendState = dic.Item("sendState")
    convState = dic.Item("convState")
    sendNum = dic.Item("sendNum")
    receiveNum = dic.Item("receiveNum")
    receiveName = dic.Item("receiveName")
    sendPageCnt = dic.Item("sendPageCnt")
    successPageCnt = dic.Item("successPageCnt")
    failPageCnt = dic.Item("failPageCnt")
    refundPageCnt = dic.Item("refundPageCnt")
    cancelPageCnt = dic.Item("cancelPageCnt")
    reserveDT = dic.Item("reserveDT")
    sendDT = dic.Item("sendDT")
    resultDT = dic.Item("resultDT")
    sendResult = dic.Item("sendResult")
  
End Sub
