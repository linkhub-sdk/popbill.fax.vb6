VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBFAXService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Popbill API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dll. most of windows system has it
' on system32 by default.
'
' http://www.popbill.com
' Author : Kim Seongjun
' Written : 2014-04-22
' Contributor : Jeong Yohan (code@linkhubcorp.com)
' Updated : 2022-04-06
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"
Private Const ServiceURL_Static_REAL = "https://static-popbill.linkhub.co.kr"
Private Const ServiceURL_Static_TEST = "https://static-popbill-test.linkhub.co.kr"
Private Const ServiceURL_GA_REAL = "https://ga-popbill.linkhub.co.kr"
Private Const ServiceURL_GA_TEST = "https://ga-popbill-test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private Const adTypeBinary As Long = 1
Private Const adTypeText As Long = 2
Private Const adModeReadWrite As Long = 3

Private m_IsTest As Boolean
Private m_IPRestrictOnOff As Boolean
Private m_UseStaticIP As Boolean
Private m_UseGAIP As Boolean
Private m_UseLocalTimeYN As Boolean
Private m_token_Dic
Private m_Linkhub As Linkhub

Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub

Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'최근 오류코드 확인
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'최근오류메시지 확인
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'테스트 플래그
Public Property Let IsTest(ByVal value As Boolean)
    m_IsTest = value
End Property

Public Property Let IPRestrictOnOff(ByVal value As Boolean)
    m_IPRestrictOnOff = value
End Property

Public Property Let UseStaticIP(ByVal value As Boolean)
    m_UseStaticIP = value
End Property

Public Property Let UseGAIP(ByVal value As Boolean)
    m_UseGAIP = value
End Property

Public Property Let UseLocalTimeYN(ByVal value As Boolean)
    m_UseLocalTimeYN = value
End Property

'초기화
Public Sub Initialize(LinkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.LinkID = LinkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    ''팩스
    m_scope.Add "160"
    
    m_IsTest = False
    m_IPRestrictOnOff = True
    m_UseStaticIP = False
    m_UseGAIP = False
    m_UseLocalTimeYN = False
End Sub

Private Function getTargetURL() As String
    If m_UseGAIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_GA_TEST, ServiceURL_GA_REAL)
    ElseIf m_UseStaticIP Then
        getTargetURL = IIf(m_IsTest, ServiceURL_Static_TEST, ServiceURL_Static_REAL)
    Else
        getTargetURL = IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL)
    End If
End Function

Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists(CorpNum) Then
        Set m_Token = m_token_Dic.Item(CorpNum)
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = m_Linkhub.GetTime(m_UseStaticIP, m_UseGAIP, m_UseLocalTimeYN)
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope, IIf(m_IPRestrictOnOff, "", "*"), m_UseStaticIP, m_UseGAIP, m_UseLocalTimeYN)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
        If m_token_Dic.Exists(CorpNum) Then
            m_token_Dic.Remove (CorpNum)
        End If

        m_token_Dic.Add CorpNum, m_Token
        
    End If
    
    getSession_token = m_Token.session_token

End Function

Private Function httpGET(URL As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("GET", getTargetURL() + URL, False)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(URL As String, BearerToken As String, postdata As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(URL, BearerToken, "", postdata, UserID)
    
End Function

Private Function httpPOST_override(URL As String, BearerToken As String, override As String, postdata As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", getTargetURL() + URL, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_override = m_Linkhub.parse(Response)

End Function

Private Function GetOnlyFileName(ByVal FilePath As String) As String

     Dim Temp() As String
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))

End Function

Private Function httpPOST_File(URL As String, BearerToken As String, postdata As String, FilePaths As Collection, Optional UserID As String) As Variant
     
    Dim winhttp1
    Dim boundary As String
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", getTargetURL() + URL, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    Dim applicationform As String
    
    If postdata <> "" Then
        applicationform = "--" & boundary & vbCrLf & _
                          "Content-Disposition: form-data; name=""form""" & vbCrLf & _
                          "Content-Type: Application/json" & vbCrLf & vbCrLf & _
                          postdata
    
    End If
    
    Dim btFilePart As New Collection
    
    'Dim btFilePart() As Byte
    Dim filePartSize As Long
    
    filePartSize = 0
    
    Dim fileCnt As Integer
    
    For fileCnt = 1 To FilePaths.Count
        Dim FilePath As String
        
        FilePath = FilePaths.Item(fileCnt)
        
        If FilePath = "" Then Err.Raise -99999999, "POPBILL", "전송파일 경로가 잘못되었습니다."
        
        Dim FileHeader As String
        
        FileHeader = vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""file""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
               "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
               
        
        '### 이미지 바이너리를 얻는다.
        Dim binary() As Byte
        Open FilePath For Binary Access Read As #1
             ReDim binary(LOF(1))
             Get #1, , binary
        Close #1
              
        Dim btFHeader() As Byte
        Dim headerSize, fileSize As Long
        btFHeader = UniStrToUTF8(FileHeader)
        
        headerSize = UBound(btFHeader) + 1
        fileSize = UBound(binary) - LBound(binary)
        Dim btFile() As Byte
        
        ReDim btFile(headerSize + fileSize)
        
        CopyMemory btFile(0), btFHeader(0), headerSize
        CopyMemory btFile(headerSize), binary(0), fileSize
        
        filePartSize = filePartSize + UBound(btFile)
        
        btFilePart.Add btFile
        
        
    Next
    
    
    Dim tail As String
    
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    
    Dim btForm() As Byte, BtTail() As Byte, Buffer() As Byte
    Dim formSize, tailSize As Long
    
    btForm = StrConv(applicationform, vbFromUnicode)
    BtTail = StrConv(tail, vbFromUnicode)
    
    formSize = UBound(btForm) + 1
    tailSize = UBound(BtTail) + 1
    
    ReDim Buffer(formSize + filePartSize + tailSize)
    CopyMemory Buffer(0), btForm(0), formSize
    
    Dim i As Integer
    Dim offsetPost As Long
    offsetPost = formSize
    For i = 1 To btFilePart.Count
        Dim tmp() As Byte
        tmp = btFilePart.Item(i)
        
        CopyMemory Buffer(offsetPost), tmp(0), UBound(tmp)
        offsetPost = offsetPost + UBound(tmp)
    Next
    
    CopyMemory Buffer(formSize + filePartSize), BtTail(0), tailSize
    
    winhttp1.Send (Buffer)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_File = m_Linkhub.parse(Response)

End Function
'연동회원 아이디 중복 확인
Public Function CheckID(IDString As String) As PBResponse
On Error GoTo ErrHandler
    If IDString = "" Then
        Err.Raise -99999999, "POPBILL", "중복여부를 확인할 아이디가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/IDCheck?ID=" + IDString, "")
    
    Set CheckID = New PBResponse
    CheckID.code = result.Item("code")
    CheckID.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CheckID = Nothing
End Function

'연동회원 담당자 정보확인
Public Function GetContactInfo(CorpNum As String, ContactID As String, Optional UserID As String = "") As PBContactInfo
On Error GoTo ErrHandler
    Dim result As Variant
    Dim postdata As String
    
    Dim tmp
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactID
            
    postdata = m_Linkhub.toString(tmp)
    
    
    Set result = httpPOST("/Contact", getSession_token(CorpNum), postdata, UserID)
    
    Set GetContactInfo = New PBContactInfo
    GetContactInfo.fromDictionary result
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetContactInfo = Nothing
End Function

'연동회원 담당자 목록조회
Public Function ListContact(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
        
    Dim result As Variant
       
    Set result = httpGET("/IDs", getSession_token(CorpNum), UserID)
    
    Set ListContact = New Collection
    
    Dim T As Variant
    Dim info As PBContactInfo
    
    For Each T In result
        Set info = New PBContactInfo
        
        info.fromDictionary T
        
        ListContact.Add info
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set ListContact = Nothing
End Function
'담당자 추가
Public Function RegistContact(CorpNum As String, ContactInfo As PBContactInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set RegistContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "pwd", ContactInfo.pwd
    tmp.Add "Password", ContactInfo.Password
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
    tmp.Add "searchRole", ContactInfo.searchRole
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs/New", getSession_token(CorpNum), postdata, UserID)
    
    Set RegistContact = New PBResponse
    
    RegistContact.code = result.Item("code")
    RegistContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'담당자 정보 수정
Public Function UpdateContact(CorpNum As String, ContactInfo As PBContactInfo, UserID As String) As PBResponse
On Error GoTo ErrHandler
    Set UpdateContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
    tmp.Add "searchRole", ContactInfo.searchRole
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateContact = New PBResponse
    
    UpdateContact.code = result.Item("code")
    UpdateContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'회사정보 수정
Public Function UpdateCorpInfo(CorpNum As String, CorpInfo As PBCorpInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set UpdateCorpInfo = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ceoname", CorpInfo.CEOName
    tmp.Add "corpName", CorpInfo.CorpName
    tmp.Add "addr", CorpInfo.Addr
    tmp.Add "bizType", CorpInfo.BizType
    tmp.Add "bizClass", CorpInfo.BizClass
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/CorpInfo", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateCorpInfo = New PBResponse
    
    UpdateCorpInfo.code = result.Item("code")
    UpdateCorpInfo.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'회사정보 조회
Public Function GetCorpInfo(CorpNum As String, Optional UserID As String = "") As PBCorpInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/CorpInfo", getSession_token(CorpNum), UserID)
    
    Set GetCorpInfo = New PBCorpInfo
        
    GetCorpInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetCorpInfo = Nothing
End Function

'파트너 포인트 충전 URL - 2017/08/30 추가
Public Function GetPartnerURL(CorpNum As String, TOGO As String) As String
    On Error GoTo ErrHandler
    
    GetPartnerURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO, m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPartnerURL = m_LastErrCode
End Function


'회원잔액조회
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), m_UseStaticIP, m_UseGAIP)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPartnerBalance = m_LastErrCode
End Function
'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetPopbillURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPopbillURL = ""
End Function
'팝빌 로그인 기본 URL
Public Function GetAccessURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
   Dim result As Variant
   
   Set result = httpGET("/?TG=LOGIN", getSession_token(CorpNum), UserID)
   
   GetAccessURL = result.Item("url")
   Exit Function
ErrHandler:
   m_LastErrCode = Err.number
   m_LastErrMessage = Err.Description
   GetAccessURL = ""
End Function

'팝빌 연동회원 포인트 충전 URL
Public Function GetChargeURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
   Dim result As Variant
   
   Set result = httpGET("/?TG=CHRG", getSession_token(CorpNum), UserID)
   
   GetChargeURL = result.Item("url")
   Exit Function
ErrHandler:
   m_LastErrCode = Err.number
   m_LastErrMessage = Err.Description
   GetChargeURL = ""
End Function

'팝빌 연동회원 포인트 결제내역 팝업 URL
Public Function GetPaymentURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=PAYMENT", getSession_token(CorpNum), UserID)
    
    GetPaymentURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPaymentURL = ""
End Function

'팝빌 연동회원 포인트 사용내역 팝업 URL
Public Function GetUseHistoryURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=USEHISTORY", getSession_token(CorpNum), UserID)
    
    GetUseHistoryURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetUseHistoryURL = ""
End Function

'회원가입 여부
Public Function CheckIsMember(CorpNum As String, LinkID As String) As PBResponse
On Error GoTo ErrHandler
    If CorpNum = "" Then
        Err.Raise -99999999, "POPBILL", "팝빌회원의 사업자번호가 입력되지 않았습니다."
    End If

    If LinkID = "" Then
        Err.Raise -99999999, "POPBILL", "링크아이디(LinkID)가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + LinkID, "")
    
    Set CheckIsMember = New PBResponse
    CheckIsMember.code = result.Item("code")
    CheckIsMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CheckIsMember = Nothing
End Function
'회원가입
Public Function JoinMember(JoinInfo As PBJoinForm) As PBResponse
On Error GoTo ErrHandler
    Set JoinMember = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.LinkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.CorpName
    tmp.Add "CEOName", JoinInfo.CEOName
    tmp.Add "Addr", JoinInfo.Addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.BizClass
    tmp.Add "BizType", JoinInfo.BizType
    tmp.Add "ContactName", JoinInfo.ContactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.id
    tmp.Add "PWD", JoinInfo.pwd
    tmp.Add "Password", JoinInfo.Password
    
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postdata)
    
    Set JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'과금정보 조회
Public Function GetChargeInfo(CorpNum As String, Optional UserID As String = "") As PBChargeInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/FAX/ChargeInfo", getSession_token(CorpNum), UserID)
    
    Set GetChargeInfo = New PBChargeInfo
        
    GetChargeInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetChargeInfo = Nothing
End Function

'''''''''''''''''''''''''' End Of PopbillBase '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''단가확인
Public Function GetUnitCost(CorpNum As String) As Single
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/FAX/UnitCost", getSession_token(CorpNum))
    
    GetUnitCost = result.Item("unitCost")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode

End Function

'팩스 전송
Public Function SendFAX(CorpNum As String, sendNum As String, receivers As Collection, FilePaths As Collection, Optional reserveDT As String, Optional UserID As String, Optional senderName As String, Optional adsYN As Boolean = False, Optional title As String = "", Optional requestNum As String = "") As String
On Error GoTo ErrHandler
    
    If receivers Is Nothing Or receivers.Count = 0 Then Err.Raise -99999999, "POPBILL", "수신자정보 가 입력되지 않았습니다."
    If FilePaths Is Nothing Or FilePaths.Count = 0 Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."
    If FilePaths.Count > 20 Then Err.Raise -99999999, "POPBILL", "1회 전송 가능한 파일갯수는 20개입니다."
    
    Dim result As Variant
    Dim postdata As String
    Dim Form As Variant
    Dim colRecevers As New Collection
    Dim receiver As PBReceiver
    
    Set Form = CreateObject("Scripting.Dictionary")
    
    If adsYN Then Form.Add "adsYN", True
    
    Form.Add "snd", sendNum
    
    If senderName <> "" Then Form.Add "sndnm", senderName
    If reserveDT <> "" Then Form.Add "sndDT", reserveDT
    If title <> "" Then Form.Add "title", title
    If requestNum <> "" Then Form.Add "requestNum", requestNum
   
    Form.Add "fCnt", FilePaths.Count
    
    
    For Each receiver In receivers
    
        colRecevers.Add receiver.toDictionary
    
    Next
    
    Form.Add "rcvs", colRecevers
    
    postdata = m_Linkhub.toString(Form)
    
    Set result = httpPOST_File("/FAX", getSession_token(CorpNum), postdata, FilePaths, UserID)
    
    SendFAX = result.Item("receiptNum")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    SendFAX = ""
End Function

'팩스 재전송
Public Function ResendFAX(CorpNum As String, receiptNum As String, sendNum As String, senderName As String, receivers As Collection, Optional reserveDT As String, Optional UserID As String = "", Optional title As String = "", Optional requestNum As String = "") As String
On Error GoTo ErrHandler
    If Not Len(receiptNum) = 18 Then
        Err.Raise -99999999, "POPBILL", "팩스접수번호(receiptNum)가 올바르지 않습니다."
    End If
    
    Dim result As Variant
    Dim postdata As String
    Dim Form As Variant
    Dim colRecevers As New Collection
    Dim receiver As PBReceiver
    
    Set Form = CreateObject("Scripting.Dictionary")
    
    If sendNum <> "" Then Form.Add "snd", sendNum
    If senderName <> "" Then Form.Add "sndnm", senderName
    If reserveDT <> "" Then Form.Add "sndDT", reserveDT
    If title <> "" Then Form.Add "title", title
    If requestNum <> "" Then Form.Add "requestNum", requestNum
    
    '수신정보 배열이 존재하는 경우
    If Not receivers Is Nothing Then
        For Each receiver In receivers
            colRecevers.Add receiver.toDictionary
        Next
        Form.Add "rcvs", colRecevers
    End If
    
    postdata = m_Linkhub.toString(Form)
    
    Set result = httpPOST("/FAX/" & receiptNum, getSession_token(CorpNum), postdata, UserID)
    
    ResendFAX = result.Item("receiptNum")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    ResendFAX = ""
End Function


'팩스 재전송(요청번호 할당)
Public Function ResendFAXRN(CorpNum As String, OrgRequestNum As String, sendNum As String, senderName As String, receivers As Collection, Optional reserveDT As String, Optional UserID As String = "", Optional title As String = "", Optional requestNum As String = "") As String
On Error GoTo ErrHandler
    If OrgRequestNum = "" Then Err.Raise -99999999, "POPBILL", "원본팩스번호가 입력되지 않았습니다."
    
    Dim result As Variant
    Dim postdata As String
    Dim Form As Variant
    Dim colRecevers As New Collection
    Dim receiver As PBReceiver
    
    Set Form = CreateObject("Scripting.Dictionary")
    
    If sendNum <> "" Then Form.Add "snd", sendNum
    If senderName <> "" Then Form.Add "sndnm", senderName
    If reserveDT <> "" Then Form.Add "sndDT", reserveDT
    If title <> "" Then Form.Add "title", title
    If requestNum <> "" Then Form.Add "requestNum", requestNum
    
    '수신정보 배열이 존재하는 경우
    If Not receivers Is Nothing Then
        For Each receiver In receivers
            colRecevers.Add receiver.toDictionary
        Next
        Form.Add "rcvs", colRecevers
    End If
    
    postdata = m_Linkhub.toString(Form)
    
    Set result = httpPOST("/FAX/Resend/" & OrgRequestNum, getSession_token(CorpNum), postdata, UserID)
    
    ResendFAXRN = result.Item("receiptNum")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    ResendFAXRN = ""
End Function

'전송내역 확인
Public Function GetMessages(CorpNum As String, receiptNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    If Not Len(receiptNum) = 18 Then
        Err.Raise -99999999, "POPBILL", "팩스접수번호(receiptNum)가 올바르지 않습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/FAX/" & receiptNum, getSession_token(CorpNum), UserID)
    
    Set GetMessages = New Collection
    
    Dim T As Variant
    Dim sentMsg As PBFaxInfo
    
    For Each T In result
        Set sentMsg = New PBFaxInfo
        
        sentMsg.fromDictionary T
        
        GetMessages.Add sentMsg
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetMessages = Nothing

End Function

'전송내역 확인 (요청번호 할당)
Public Function GetMessagesRN(CorpNum As String, requestNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    Dim result As Variant
    If requestNum = "" Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/FAX/Get/" & requestNum, getSession_token(CorpNum), UserID)
    
    Set GetMessagesRN = New Collection
    
    Dim T As Variant
    Dim sentMsg As PBFaxInfo
    
    For Each T In result
        Set sentMsg = New PBFaxInfo
        
        sentMsg.fromDictionary T
        
        GetMessagesRN.Add sentMsg
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetMessagesRN = Nothing

End Function

'예약전송 취소
Public Function CancelReserve(CorpNum As String, receiptNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If Not Len(receiptNum) = 18 Then
        Err.Raise -99999999, "POPBILL", "팩스접수번호(receiptNum)가 올바르지 않습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/FAX/" & receiptNum & "/Cancel", getSession_token(CorpNum), UserID)
    Set CancelReserve = New PBResponse
    
    CancelReserve.code = result.Item("code")
    CancelReserve.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CancelReserve = Nothing

End Function

'예약전송 취소 (요청번호 할당)
Public Function CancelReserveRN(CorpNum As String, requestNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    If requestNum = "" Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/FAX/Cancel/" & requestNum, getSession_token(CorpNum), UserID)
    Set CancelReserveRN = New PBResponse
    
    CancelReserveRN.code = result.Item("code")
    CancelReserveRN.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CancelReserveRN = Nothing

End Function

'팩스 관련 URL
Public Function GetURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/FAX/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetURL = ""
End Function

'발신번호 등록여부 확인
Public Function CheckSenderNumber(CorpNum As String, SenderNumber As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    If SenderNumber = "" Then
        Err.Raise -99999999, "POPBILL", "발신번호가 입력되지 않았습니다."
    End If

    Set result = httpGET("/FAX/CheckSenderNumber/" + SenderNumber, getSession_token(CorpNum), UserID)
    
    Set CheckSenderNumber = New PBResponse
    
    CheckSenderNumber.code = result.Item("code")
    CheckSenderNumber.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CheckSenderNumber = Nothing
End Function

'발신번호 관리 팝업 URL
Public Function GetSenderNumberMgtURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/FAX/?TG=SENDER", getSession_token(CorpNum), UserID)
    
    GetSenderNumberMgtURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetSenderNumberMgtURL = ""
End Function


'팩스 전송내역 팝업 URL
Public Function GetSentListURL(CorpNum As String, UserID As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/FAX/?TG=BOX", getSession_token(CorpNum), UserID)
    
    GetSentListURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetSentListURL = ""
End Function


'팩스전송내역 목록조회
Public Function Search(CorpNum As String, SDate As String, EDate As String, state As Collection, ReserveYN As Boolean, SenderOnly As Boolean, Page As Integer, PerPage As Integer, Order As String, Optional QString As String = "") As PBFaxSearchList
On Error GoTo ErrHandler
    Dim result As Variant
    Dim uri As String
    Dim i As Integer
    
    uri = "/FAX/Search"
    uri = uri + "?SDate=" + SDate
    uri = uri + "&EDate=" + EDate
    
    If state.Count > 0 Then
        uri = uri + "&State="
        For i = 1 To state.Count
            If i = state.Count Then
                uri = uri + state(i)
            Else
                uri = uri + state(i) + ","
            End If
        Next
    End If
    
    If ReserveYN Then
        uri = uri + "&ReserveYN=1"
    End If
    
    If SenderOnly Then
        uri = uri + "&SenderOnly=1"
    End If
    
    uri = uri + "&Page=" + CStr(Page)
    uri = uri + "&PerPage=" + CStr(PerPage)
    uri = uri + "&Order=" + Order
    
    If QString <> "" Then uri = uri + "&QString=" + URLEncode(QString)
    
    Set result = httpGET(uri, getSession_token(CorpNum), "")
    
    Set Search = New PBFaxSearchList
    
    Search.fromDictionary result
    
    
    Exit Function
    
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set Search = Nothing
End Function

' 팩스발신번호 목록 조회
Public Function GetSenderNumberList(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler

    Dim result As Variant
    
    Set result = httpGET("/FAX/SenderNumber", getSession_token(CorpNum), UserID)
    
    Set GetSenderNumberList = New Collection
    
    Dim T As Variant
    Dim info As PBFaxSenderNumber
    
    For Each T In result
        Set info = New PBFaxSenderNumber
        info.fromDictionary T
        GetSenderNumberList.Add info
    Next
    
    Exit Function
   
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetSenderNumberList = Nothing
End Function

'팩스 미리보기 URL
Public Function GetPreviewURL(CorpNum As String, receiptNum As String, UserID As String) As String
On Error GoTo ErrHandler
    If Not Len(receiptNum) = 18 Then
        Err.Raise -99999999, "POPBILL", "접수번호(receiptNum)가 올바르지 않습니다."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/FAX/Preview/" + receiptNum, getSession_token(CorpNum), UserID)
    
    GetPreviewURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPreviewURL = ""
End Function


Private Function URLEncode(ByVal StringToEncode As String) As String
   Dim i                As Integer
   Dim iAsc             As Long
   Dim sTemp            As String
   
   Dim ByteArrayToEncode() As Byte
 
   ByteArrayToEncode = ADO_EncodeUTF8(StringToEncode)
   
   For i = 0 To UBound(ByteArrayToEncode)
      iAsc = ByteArrayToEncode(i)
      Select Case iAsc
         Case 32 'space
            sTemp = "+"
         Case 48 To 57, 65 To 90, 97 To 122
            sTemp = Chr(ByteArrayToEncode(i))
         Case Else
            Debug.Print iAsc
            sTemp = "%" & Hex(iAsc)
      End Select
      URLEncode = URLEncode & sTemp
   Next
 
End Function
'Purpose: UTF16 to UTF8 using ADO
Private Function ADO_EncodeUTF8(ByVal strUTF16 As String) As Byte()
 
   Dim objStream        As Object
   Dim Data()           As Byte
 
   Set objStream = CreateObject("ADODB.Stream")
   objStream.Charset = "utf-8"
   objStream.Mode = adModeReadWrite
   objStream.Type = adTypeText
   objStream.Open
   objStream.WriteText strUTF16
   objStream.Flush
   objStream.Position = 0
   objStream.Type = adTypeBinary
   objStream.Read 3 ' skip BOM
   Data = objStream.Read()
   objStream.Close
   ADO_EncodeUTF8 = Data
 
End Function

