VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public receiverNum As String    '수신자 번호
Public receiverName As String  '수신자 명칭
Public interOPRefKey As String  '파트너 지정키


Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "rcv", receiverNum
    toDictionary.Add "rcvnm", receiverName
    toDictionary.Add "interOPRefKey", interOPRefKey
    
End Function
