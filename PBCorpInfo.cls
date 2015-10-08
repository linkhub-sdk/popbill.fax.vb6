VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCorpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CEOName            As String
Public CorpName           As String
Public Addr               As String
Public BizType            As String
Public BizClass           As String

Public Sub fromDictionary(dic As Variant)
    CEOName = dic.Item("ceoname")
    CorpName = dic.Item("corpName")
    Addr = dic.Item("addr")
    BizType = dic.Item("bizType")
    BizClass = dic.Item("bizClass")
End Sub





