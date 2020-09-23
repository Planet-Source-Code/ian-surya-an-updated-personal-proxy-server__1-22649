Attribute VB_Name = "modGlobal"
Option Explicit

Public Sub Main()
    If (App.PrevInstance = False) Then
        frmMain.Show
    End If
End Sub

Public Function IsHTTPHeader(Data As String) As Boolean
Const HEADER_HTTP = "HTTP"
Const METHOD_GET = "GET"
Const METHOD_POST = "POST"
Const METHOD_HEAD = "HEAD"
Const METHOD_PROPFIND = "PROPFIND"
Const METHOD_OPTION = "OPTIONS"
Const METHOD_CONNECT = "CONNECT"
Dim lpos As Long, Method As String

    lpos = InStr(1, Data, " ", vbTextCompare)
    If lpos <> 0 Then
        Method = UCase(Left$(Data, lpos - 1))
        Select Case Method
        Case HEADER_HTTP, METHOD_GET, METHOD_POST, METHOD_HEAD, METHOD_PROPFIND, METHOD_OPTION, METHOD_CONNECT: IsHTTPHeader = True
        End Select
    End If
End Function

Public Function GenerateAuthorizationHeader(Header As String, ProxyAuthorizationString As String) As String
Dim tmpString As String, AuthorizationString As String

    tmpString = Header
    
    If InStr(ProxyAuthorizationString, "Proxy-Authorization") <> 0 Then
        tmpString = DeleteHttpHeader(tmpString, "Proxy-Connection")
        tmpString = DeleteHttpHeader(tmpString, "Proxy-Authorization")
    
        AuthorizationString = "Proxy-Connection: Keep-Alive" & vbCrLf
    Else
        tmpString = DeleteHttpHeader(tmpString, "Connection")
        tmpString = DeleteHttpHeader(tmpString, "WWW-Authorization")
        
        AuthorizationString = "Connection: Keep-Alive" & vbCrLf
    End If
    
    AuthorizationString = AuthorizationString & ProxyAuthorizationString
    
    tmpString = Replace(tmpString, vbCrLf, vbCrLf & AuthorizationString, 1, 1)
    
    GenerateAuthorizationHeader = tmpString
End Function

Public Function DeleteHttpHeader(Header As String, HeaderCaption As String) As String
Dim lpos As Long
Dim endpos As Long
Dim HeaderData As String

    lpos = InStr(1, Header, HeaderCaption & ":", vbTextCompare)
    If lpos <> 0 Then
        endpos = InStr(lpos + 1, Header, vbCrLf, vbTextCompare)
        HeaderData = Mid$(Header, lpos, endpos - lpos)
        DeleteHttpHeader = Replace(Header, HeaderData & vbCrLf, "", 1, 1, vbTextCompare)
    Else
        DeleteHttpHeader = Header
    End If
    
End Function

Public Function AddHttpHeader(Header As String, HeaderCaption As String, HeaderValue As String) As String
Dim lpos As Long
Dim endpos As Long
Dim HeaderData As String

    lpos = InStr(1, Header, HeaderCaption & ":", vbTextCompare)
    If lpos <> 0 Then
        endpos = InStr(lpos + 1, Header, vbCrLf, vbTextCompare)
        HeaderData = Mid$(Header, lpos, endpos - lpos)
        AddHttpHeader = Replace(Header, HeaderData & vbCrLf, HeaderCaption & ": " & HeaderValue & vbCrLf, 1, 1, vbTextCompare)
    Else
        AddHttpHeader = Replace(Header, vbCrLf, vbCrLf & HeaderCaption & ": " & HeaderValue & vbCrLf, 1, 1, vbTextCompare)
    End If
    
End Function

Public Function FilterRequestHeader(Header As String) As String
Dim tmpString As String

    tmpString = Header
    If Filter_Reload Then
        tmpString = AddHttpHeader(tmpString, "Pragma", "no-cache")
        tmpString = AddHttpHeader(tmpString, "Cache-Control", "no-cache")
        'tmpString = DeleteHttpHeader(tmpString, "If-Modified-Since")
    End If
    If Filter_Disable_Cookie Then
        tmpString = DeleteHttpHeader(tmpString, "Cookie")
    End If
    If Filter_Hide_UserAgent Then
        If UserAgent <> "" Then
            tmpString = AddHttpHeader(tmpString, "User-Agent", UserAgent)
        Else
            tmpString = AddHttpHeader(tmpString, "User-Agent", C_USER_AGENT_PPS)
        End If
    End If
    'If UseAuthentication Then
        'tmpString = DeleteHttpHeader(tmpString, "Proxy-Authorization")
    'End If
    If Not UseProxy And Left(LCase(GetResource(Header)), Len("http://" & LCase(GetHttpHeader(Header, "Host")) & "/")) = "http://" & LCase(GetHttpHeader(Header, "Host")) & "/" Then
        tmpString = Replace(tmpString, "http://" & LCase(GetHttpHeader(Header, "Host")) & "/", "/")
    End If
    
    FilterRequestHeader = tmpString
End Function

Public Function GetResource(Header As String) As String
Dim lpos As Long, tmpString As String
Dim Methode As String, Resource As String, Version As String

    tmpString = Header
    tmpString = GetToken(tmpString, vbCrLf)
    tmpString = tmpString & " "
    Methode = GetToken(tmpString, " ")
    Resource = GetToken(tmpString, " ")
    Version = GetToken(tmpString, " ")
    
    GetResource = Resource
End Function

Public Function FilterResponseHeader(Header As String) As String
Dim tmpString As String

    tmpString = Header
    If Filter_Hide_Server Then
        If PersonalProxyName <> "" Then
            tmpString = AddHttpHeader(tmpString, "Server", PersonalProxyName)
        Else
            tmpString = AddHttpHeader(tmpString, "Server", C_PERSONAL_PROXY)
        End If
    End If
    If Filter_Hide_Proxy Then
        If LocalComputerName <> "" Then
            tmpString = AddHttpHeader(tmpString, "Via", LocalComputerName)
        Else
            tmpString = DeleteHttpHeader(tmpString, "Via")
        End If
    End If
    If Filter_Disable_Cookie Then
        tmpString = DeleteHttpHeader(tmpString, "Set-Cookie")
    End If
    
    FilterResponseHeader = tmpString
End Function

Public Function GetHttpHeader(Header As String, HeaderCaption As String) As String
Dim lpos As Long
Dim endpos As Long
Dim HeaderData As String

    lpos = InStr(1, Header, HeaderCaption & ":", vbTextCompare)
    If lpos <> 0 Then
        endpos = InStr(lpos + 1, Header, vbCrLf, vbTextCompare)
        HeaderData = Mid$(Header, lpos + Len(HeaderCaption) + 2, endpos - (lpos + Len(HeaderCaption) + 2))
    End If
    GetHttpHeader = HeaderData
    
End Function

Public Function IsInUserCollection(Col As Collection, Key As String) As Boolean
Dim i As Long
    
    For i = 1 To Col.Count
        If GetUser(Col(i).Key) = GetUser(Key) And GetPassword(Col(i).Key) = GetPassword(Key) Then
            IsInUserCollection = True
            Exit For
        End If
    Next i
End Function

Public Function IsInCollection(Col As Collection, Key As String) As Boolean
'*** Note : do not set the ide to break on all error but use break on unhandled error
'***        or the this checking won't work
Dim i As Long

    On Error GoTo errHandler
    
    If Col(Key).Key = Key Then
        IsInCollection = True
    End If
    Exit Function
    
errHandler:
End Function

Public Function GetToken(sData As String, Delimiter As String) As String
Dim tmpString As String, lpos As Long

    lpos = InStr(1, sData, Delimiter, vbTextCompare)
    If lpos <> 0 Then
        tmpString = Left$(sData, lpos - 1)
        sData = Mid$(sData, lpos + Len(Delimiter))
    Else
        tmpString = sData
        sData = ""
    End If
    GetToken = tmpString
End Function

Public Function FormatByte(Number As Double) As String
    FormatByte = Format(Number / 1024, "#,##0.00")
End Function

