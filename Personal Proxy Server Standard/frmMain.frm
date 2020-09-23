VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Personal Proxy Server"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLog 
      Caption         =   "Generate Log File"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   90
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin TabDlg.SSTab tabProxy 
      Height          =   4995
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8811
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraLog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Header"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraResponse"
      Tab(1).Control(1)=   "fraRequest"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Stat"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraStat"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraStat 
         Caption         =   "Connection Request Statistic"
         Height          =   4485
         Left            =   -74910
         TabIndex        =   10
         Top             =   390
         Width           =   6435
         Begin MSFlexGridLib.MSFlexGrid flxStatistic 
            Height          =   4095
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   7223
            _Version        =   393216
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame fraResponse 
         Caption         =   "Response Header"
         Height          =   1905
         Left            =   -74910
         TabIndex        =   7
         Top             =   2340
         Width           =   6465
         Begin VB.TextBox txtResponse 
            Height          =   1515
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   270
            Width           =   6195
         End
      End
      Begin VB.Frame fraRequest 
         Caption         =   "Request Header"
         Height          =   1905
         Left            =   -74910
         TabIndex        =   6
         Top             =   390
         Width           =   6465
         Begin VB.TextBox txtRequest 
            Height          =   1515
            Left            =   150
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   6195
         End
      End
      Begin VB.Frame fraLog 
         Caption         =   "Proxy Logs"
         Height          =   4485
         Left            =   90
         TabIndex        =   4
         Top             =   390
         Width           =   6435
         Begin VB.ListBox lstLog 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   4155
            ItemData        =   "frmMain.frx":035E
            Left            =   120
            List            =   "frmMain.frx":0360
            TabIndex        =   5
            Top             =   210
            Width           =   6195
         End
      End
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear"
      Height          =   405
      Left            =   2100
      TabIndex        =   2
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton cmdConfiguration 
      Caption         =   "Config"
      Height          =   405
      Left            =   1050
      TabIndex        =   1
      Top             =   0
      Width           =   1035
   End
   Begin VB.Timer tmrClient 
      Index           =   0
      Interval        =   10
      Left            =   4470
      Top             =   0
   End
   Begin VB.Timer tmrServer 
      Index           =   0
      Interval        =   10
      Left            =   4020
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckClient 
      Index           =   0
      Left            =   3150
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3570
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Start"
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.Menu mnuHeader 
      Caption         =   "Header"
      Visible         =   0   'False
      Begin VB.Menu mnuClearHeader 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "Log"
      Visible         =   0   'False
      Begin VB.Menu mnuClearLog 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuStat 
      Caption         =   "Stat"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal addr As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32" (addr As Long, addrLen As Long, addrType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private objSystray As clsSysTray

Private Const DEBUG_MODE = False

Dim BlockingClient() As Boolean
Dim BlockingServer() As Boolean

Dim ServerConnection As Collection
Dim ClientConnection As Collection

Private Sub cmdClearLog_Click()
    lstLog.Clear
    txtRequest.Text = ""
    txtResponse.Text = ""
End Sub

Private Sub cmdConfiguration_Click()
    LoadConfigurationScreen Me, netProxy
End Sub

Private Sub cmdSwitch_Click()
    If cmdSwitch.Caption = "Start" Then
        objSystray.ToolTip = "Personal Proxy Server On"
        StartProxy ListeningPort
        cmdSwitch.Caption = "Stop"
    Else
        objSystray.ToolTip = "Personal Proxy Server Off"
        StopProxy
        cmdSwitch.Caption = "Start"
    End If
End Sub

Private Sub StartProxy(LocalPort As Long)
    SendMessage "Initializing Personal Proxy Server"
    InitializeSocket sckServer(0)
    sckServer(0).LocalPort = LocalPort
    sckServer(0).Listen
    SendMessage "Listening on port " & LocalPort
End Sub

Private Sub StopProxy()
Dim Socket As Winsock

    SendMessage "Shutting down Personal Proxy Server"
    For Each Socket In sckServer
        tmrServer(Socket.Index).Enabled = False
        tmrClient(Socket.Index).Enabled = False
        DoEvents
        CloseSocket Socket.Index
    Next
    SendMessage "Personal Proxy Server stopped"
End Sub

Private Sub InitializeSocket(Socket As Winsock)
On Error Resume Next

    SendMessage "Initialize Socket " & Socket.LocalPort
    Socket.Close
    Socket.LocalPort = 0
End Sub

Private Sub SendMessage(Message As String)
    lstLog.AddItem "[" & Now & "] " & Message
    If lstLog.ListCount > 10000 Then lstLog.Clear
End Sub

Private Sub flxStatistic_Click()
Dim i As Long

    With flxStatistic
        For i = 1 To ConnectionRequest.Count
            DoEvents
            If i = .Rows Then
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = ConnectionRequest(i).IPAddress
                If ConnectionRequest(i).HostName = "" Then
                    ConnectionRequest(i).HostName = NameByAddr(ConnectionRequest(i).IPAddress)
                End If
                .TextMatrix(i, 2) = ConnectionRequest(i).HostName
            End If
            .TextMatrix(i, 3) = ConnectionRequest(i).Stat_Connect_Count
            .TextMatrix(i, 4) = FormatByte(ConnectionRequest(i).Stat_Bytes_Received)
            .TextMatrix(i, 5) = FormatByte(ConnectionRequest(i).Stat_Bytes_Sent)
        Next i
    End With
End Sub

Private Sub flxStatistic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Shift = vbCtrlMask Then
        PopupMenu mnuStat, , X + 250, Y + 1100
    End If
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        End
    End If
    
    LocalIP = sckServer(0).LocalIP
    LoadUser UserList, "UserList.txt"
    LoadUser InvalidList, "Invalid.txt"
    InitializeGrid
    
    Set netProxy = New CProxy
    
    If Len(App.Path & "\" & ConfigFileName) = 0 Then LocalComputerName = sckServer(0).LocalHostName
    LoadProxyConfiguration
    
    Set ServerConnection = New Collection
    Set ClientConnection = New Collection

    Set objSystray = New clsSysTray
    Set objSystray.SourceWindow = Me
    objSystray.ChangeIcon Me.Icon
    objSystray.ToolTip = "Proxy Server Off"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdSwitch.Caption <> "Start" Then
        objSystray.ToolTip = "Personal Proxy Server Off"
        StopProxy
        cmdSwitch.Caption = "Start"
    End If
End Sub

Private Sub Form_Resize()
    tabProxy.Width = Me.ScaleWidth
    If Me.ScaleHeight > (cmdSwitch.Height) Then
        tabProxy.Height = Me.ScaleHeight - (cmdSwitch.Height)
    End If
    If tabProxy.Width > 200 Then
        fraLog.Width = tabProxy.Width - 200
        fraStat.Width = tabProxy.Width - 200
    End If
    If tabProxy.Height > 500 Then
        fraLog.Height = tabProxy.Height - 500
        fraStat.Height = tabProxy.Height - 500
    End If
    If fraLog.Width > 200 Then
        lstLog.Width = fraLog.Width - 200
        flxStatistic.Width = fraStat.Width - 200
    End If
    If fraLog.Height > 240 Then
        lstLog.Height = fraLog.Height - 240
    End If
    If fraStat.Height > 320 Then
        flxStatistic.Height = fraStat.Height - 320
    End If
    fraRequest.Width = fraLog.Width
    fraRequest.Height = fraLog.Height \ 2
    If fraRequest.Height > 400 Then
        txtRequest.Height = fraRequest.Height - 400
    End If
    If fraRequest.Width > 300 Then
        txtRequest.Width = fraRequest.Width - 300
    End If
    fraResponse.Top = fraRequest.Top + fraRequest.Height
    fraResponse.Width = fraLog.Width
    fraResponse.Height = fraLog.Height \ 2
    If fraResponse.Height > 400 Then
        txtResponse.Height = fraResponse.Height - 400
    End If
    If fraResponse.Width > 300 Then
        txtResponse.Width = fraResponse.Width - 300
    End If
    
    With flxStatistic
        .ColWidth(0) = Abs(360 / 6165 * (.Width - 100))
        .ColWidth(1) = Abs(960 / 6165 * (.Width - 100))
        .ColWidth(2) = Abs(2010 / 6165 * (.Width - 100))
        .ColWidth(3) = Abs(915 / 6165 * (.Width - 100))
        .ColWidth(4) = Abs(960 / 6165 * (.Width - 100))
        .ColWidth(5) = Abs(960 / 6165 * (.Width - 100))
        
    End With

    If Me.WindowState = vbMinimized Then
        If cmdSwitch.Caption = "Stop" Then
            objSystray.ChangeIcon Me.Icon
        Else
            objSystray.ChangeIcon frmUserLogin.Icon
        End If
        objSystray.MinToSysTray
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Socket As Winsock

    On Error Resume Next
    For Each Socket In sckClient
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then
            Unload Socket
        End If
    Next
    
    For Each Socket In sckServer
        CloseSocket Socket.Index
        If Socket.Index <> 0 Then
            Unload Socket
        End If
    Next
        
    Set netProxy = Nothing
    Set ServerConnection = Nothing
    Set ClientConnection = Nothing

    objSystray.RemoveFromSysTray
    
    Set objSystray = Nothing
End Sub

Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLog, vbPopupMenuRightButton, X + 240, Y + 1060
    End If
End Sub

Private Sub mnuClearHeader_Click()
    txtRequest.Text = ""
    txtResponse.Text = ""
End Sub

Private Sub mnuClearLog_Click()
    lstLog.Clear
End Sub

Private Sub mnuSave_Click()
    SaveUser InvalidList, "Invalid.txt"
End Sub

Private Sub sckClient_Close(Index As Integer)
    InitializeSocket sckClient(Index)
'    ClientConnection(Index).ClearBuffer
'
'    If Index <> 0 Then
'        Do While Len(ServerConnection(Index).SendBuffer) <> 0 Or sckServer(Index).State = sckClosed
'            DoEvents
'        Loop
'        InitializeSocket sckServer(Index)
'        ServerConnection(Index).ClearBuffer
'    End If
End Sub

Private Sub sckClient_Connect(Index As Integer)
Dim vData As String
Static Blocking As Boolean

    'DoEvents
    If sckClient(Index).State = sckConnected Then
        vData = ClientConnection(Index).SendBuffer
        If Len(vData) <> 0 And Not Blocking Then
            'Blocking = True
            vData = ClientConnection(Index).SendBuffer.GetString
            SendDataTo sckClient(Index), vData
            'DoEvents
            'ClientConnection(Index).SendBuffer = Mid(ClientConnection(Index).SendBuffer, Len(vData) + 1)
            SendMessage "Connected to Server " & sckClient(Index).RemoteHostIP & ":" & sckClient(Index).RemotePort
            'Blocking = False
            If DEBUG_MODE Then Debug.Print "send to server " & vbCrLf & vData
        End If
    End If
End Sub

Private Sub sckClient_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim vData As String
Dim lpos As Long
Dim Header As String, Data As String

    If Index <> 0 And sckClient(Index).State = sckConnected Then
        SendMessage "Receive data from server " & sckClient(Index).RemoteHostIP & ":" & sckClient(Index).RemotePort & " size: " & bytesTotal & " bytes"
        
        sckClient(Index).GetData vData
        ServerConnection(Index).Append vData
        AppendLog Index, "From Server " & Index & " :" & vbCrLf & vData
        
        If ServerConnection(Index).HeaderReceived And Not ServerConnection(Index).Connected Then
            If DEBUG_MODE Then Debug.Print "received from server " & vbCrLf & ServerConnection(Index).Header
            Header = FilterResponseHeader(ServerConnection(Index).Header)
            ServerConnection(Index).SendBuffer = Header & vbCrLf & ServerConnection(Index).Data
            ServerConnection(Index).DataSent = ServerConnection(Index).DataSent + Len(ServerConnection(Index).Data)
            ServerConnection(Index).Connected = True
            SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
            If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
        End If
    End If
End Sub

Private Sub sckClient_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'InitializeSocket sckClient(Index)
    'If Index <> 0 Then
        'ServerConnection(Index).ClearBuffer
    'End If
    
    InitializeSocket sckClient(Index)
    ClientConnection(Index).ClearBuffer
    If Index <> 0 Then
        Do While Len(ServerConnection(Index).SendBuffer) <> 0 Or sckServer(Index).State = sckClosed
            DoEvents
        Loop
        InitializeSocket sckServer(Index)
        ServerConnection(Index).ClearBuffer
    End If
End Sub

Private Sub sckServer_Close(Index As Integer)
    CloseSocket Index
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim i As Long, ActiveConnection As Long, ReceivingSocket As Winsock
    If Index = 0 Then
        ActiveConnection = 0
        For i = 0 To sckServer.Count - 1
            If i <> 0 Then If sckServer(i).State <> sckClosed Then ActiveConnection = ActiveConnection + 1
        Next i
        Set ReceivingSocket = AvailableSocket
        ReceivingSocket.Accept requestID
        If ActiveConnection < MaximumConnection Then
            SendMessage "Accept connection request from client " & AvailableSocket.RemoteHostIP & ":" & ReceivingSocket.RemotePort
        Else
            ServerConnection(ReceivingSocket.Index).Rejected = True
            SendMessage "Maximum connection reached, Connection request from client " & ReceivingSocket.RemoteHostIP & ":" & ReceivingSocket.RemotePort & " rejected"
        End If
    End If
    
    Set ReceivingSocket = Nothing
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Long, lpos As Long
Dim vData As String
Static Blocking As Boolean
Dim Header As String

    If Index <> 0 And sckServer(Index).State = sckConnected Then
        sckServer(Index).GetData vData
        SendMessage "Receive data from client " & sckServer(Index).RemoteHostIP & ":" & sckServer(Index).RemotePort & " size: " & bytesTotal & " bytes"
        ClientConnection(Index).Append vData
        AppendLog Index, "From Client " & Index & " :" & vbCrLf & vData
        AddStatReceived sckServer(Index), Len(vData)
        
        If ClientConnection(Index).HeaderReceived And Not ClientConnection(Index).Connected Then
            If DEBUG_MODE Then Debug.Print "received from client " & vbCrLf & ClientConnection(Index).Header
            If ServerConnection(Index).Rejected Then
                Header = GenerateHTMLForm(ftMaxReached)
                ServerConnection(Index).SendBuffer = Header
                SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                DoEvents
                CloseSocket Index
                If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
            ElseIf Not ServerConnection(Index).AuthorizeUser Then
                ServerConnection(Index).AuthorizeUser = CheckCredential(sckServer(Index), ClientConnection(Index).Header)
                If Not ServerConnection(Index).AuthorizeUser Then
                    Header = GenerateHTMLForm(ftAuthenticate)
                    ServerConnection(Index).SendBuffer = Header
                    SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                    If DEBUG_MODE Then Debug.Print "send to client buffer " & vbCrLf & Header
                Else
                    InitializeSocket sckClient(Index)
                    Header = FilterRequestHeader(ClientConnection(Index).Header)
                    ClientConnection(Index).SendBuffer = Header & vbCrLf & ClientConnection(Index).Data
                    ClientConnection(Index).DataSent = ClientConnection(Index).DataSent + Len(ClientConnection(Index).Data)
                    ClientConnection(Index).Connected = True
                    SendRequestHeader "Socket " & Index & " :" & vbCrLf & Header
                    AddConnectionStatistic sckServer(Index)
                    If DEBUG_MODE Then Debug.Print "send to server buffer " & vbCrLf & Header
                End If
            ElseIf Left$(ClientConnection(Index).Header, 7) = "OPTIONS" Then
                Header = GenerateHTMLForm(ftNotFound)
                ServerConnection(Index).SendBuffer = Header
                SendResponseHeader "Socket " & Index & " :" & vbCrLf & Header
                DoEvents
                CloseSocket Index
            Else
                InitializeSocket sckClient(Index)
                Header = FilterRequestHeader(ClientConnection(Index).Header)
                ClientConnection(Index).SendBuffer = Header & vbCrLf & ClientConnection(Index).Data
                ClientConnection(Index).DataSent = ClientConnection(Index).DataSent + Len(ClientConnection(Index).Data)
                ClientConnection(Index).Connected = True
                SendRequestHeader "Socket " & Index & " :" & vbCrLf & Header
                AddConnectionStatistic sckServer(Index)
                If DEBUG_MODE Then Debug.Print "send to server buffer " & vbCrLf & Header
            End If
        End If
    End If
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CloseSocket Index
End Sub

Private Function AvailableSocket() As Winsock
Dim Socket As Winsock

    For Each Socket In sckServer
        DoEvents
        If Socket.State = sckClosed Then
            CloseSocket Socket.Index
            Set AvailableSocket = Socket
            Exit Function
        End If
    Next
    Set AvailableSocket = AddNewConnection
End Function

Private Function AddNewConnection() As Winsock
Dim ServerData As New CBuffer
Dim ClientData As New CBuffer
Dim NewSocket As Long
    
    NewSocket = sckServer.Count
    
    Load sckServer(NewSocket)
    Load tmrServer(NewSocket)
    ServerData.HeaderType = htResponse
    ServerData.ClearBuffer
    ServerConnection.Add ServerData, Chr(NewSocket)
    
    Load sckClient(NewSocket)
    Load tmrClient(NewSocket)
    ClientData.HeaderType = htRequest
    ClientData.ClearBuffer
    ClientConnection.Add ClientData, Chr(NewSocket)
    ClientData.AuthenticationCounter = 0

    Set AddNewConnection = sckServer(NewSocket)
End Function

Private Sub sckServer_SendComplete(Index As Integer)
'    If ServerConnection(Index).DataSent >= ServerConnection(Index).ResponseHeader.Length Then
'        If ServerConnection(Index).ResponseHeader.GetHeader("Content-Length") <> "" Then
'        'If Not LCase(ServerConnection(Index).ResponseHeader.GetHeader("Connection")) = "keep-alive" And Not LCase(ServerConnection(Index).ResponseHeader.GetHeader("Proxy-Connection")) = "keep-alive" Then
'            CloseSocket Index
'        'End If
'        ElseIf sckClient(Index).State <> sckConnected And Len(ServerConnection(Index).SendBuffer) = 0 Then
'            CloseSocket Index
'        End If
'    End If
End Sub

Private Sub tmrClient_Timer(Index As Integer)
Dim i As Long
Dim vData As String
'Static Blocking As Boolean
ReDim Preserve BlockingClient(tmrClient.Count - 1) As Boolean

    If Index <> 0 Then
        If Not BlockingClient(Index - 1) Then
            BlockingClient(Index - 1) = True
            i = Index
            Do While Len(ClientConnection(i).SendBuffer)
                DoEvents
                vData = ClientConnection(i).SendBuffer
                If Len(vData) <> 0 Then
                    If sckClient(i).State <> sckConnected And sckClient(i).State <> sckConnecting Then
                        ConnectSocket sckClient(i), ClientConnection(i)
                    ElseIf sckClient(i).State = sckConnected And Len(vData) <> 0 Then
                        vData = ClientConnection(i).SendBuffer.GetString
                        SendDataTo sckClient(i), vData
                        AppendLog Index, "To Server " & Index & " :" & vbCrLf & vData
                        If DEBUG_MODE Then Debug.Print "send to server " & vbCrLf & vData
                    End If
                End If
            Loop
            BlockingClient(Index - 1) = False
        End If
    End If
End Sub

Private Sub SendDataTo(Socket As Winsock, vData As String)
    Socket.SendData vData
    SendMessage "Sending data to " & Socket.RemoteHostIP & ":" & Socket.RemotePort & " Size:" & Len(vData)
End Sub

Private Sub ConnectSocket(Socket As Winsock, BufferConnection As CBuffer)
Dim vProxyServer As String, vProxyPort As Long

    On Error GoTo errHandler
    
    If UseProxy Then
        vProxyServer = netProxy.Server
        vProxyPort = netProxy.Port
    Else
        vProxyServer = BufferConnection.Server
        vProxyPort = BufferConnection.Port
    End If
    Socket.Connect vProxyServer, vProxyPort
    DoEvents
    SendMessage "Connecting to server " & vProxyServer & ":" & vProxyPort
    Exit Sub
errHandler:
End Sub

Private Sub tmrServer_Timer(Index As Integer)
Dim i As Long
Dim vData As String
'Static Blocking As Boolean
ReDim Preserve BlockingServer(tmrServer.Count - 1) As Boolean

    If Index <> 0 Then
        If Not BlockingServer(Index - 1) Then
            BlockingServer(Index - 1) = True
            i = Index
            DoEvents
            If sckServer(i).State = sckConnected Then
                vData = ServerConnection(i).SendBuffer
                If Len(vData) <> 0 Then
                    vData = ServerConnection(i).SendBuffer.GetString
                    SendDataTo sckServer(i), vData
                    AddStatSent sckServer(i), Len(vData)
                    AppendLog Index, "To Client " & Index & " :" & vbCrLf & vData
                    If DEBUG_MODE Then Debug.Print "send to client " & vbCrLf & vData
                End If
            End If
            BlockingServer(Index - 1) = False
        End If
    End If
End Sub

Private Sub CloseSocket(Index As Integer)
    On Error Resume Next
    
    InitializeSocket sckClient(Index)
    If Index <> 0 Then
        ServerConnection(Index).ClearBuffer
    End If
    
    InitializeSocket sckServer(Index)
    If Index <> 0 Then
        ClientConnection(Index).ClearBuffer
    End If
End Sub

Private Sub SendRequestHeader(Message As String)
    If Len(txtRequest.Text) > 16384 Then
        txtRequest.Text = ""
    End If
    txtRequest.Text = txtRequest.Text & Message & vbCrLf
End Sub

Private Sub SendResponseHeader(Message As String)
    If Len(txtResponse.Text) > 16384 Then
        txtResponse.Text = ""
    End If
    txtResponse.Text = txtResponse.Text & Message & vbCrLf
End Sub

Private Sub InitializeGrid()
    With flxStatistic
        .Clear
        .Rows = 1
        .Cols = 6
        
        .ColWidth(0) = 360
        .ColWidth(1) = 960
        .ColWidth(2) = 2010
        .ColWidth(3) = 915
        .ColWidth(4) = 960
        .ColWidth(5) = 960
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "IP Address"
        .TextMatrix(0, 2) = "Host Name"
        .TextMatrix(0, 3) = "Connection"
        .TextMatrix(0, 4) = "Received"
        .TextMatrix(0, 5) = "Sent"
    End With
End Sub

Public Sub AppendFile(FileName As String, Data As String, Optional FileLength As Long = 0) 'test
    Dim ff As Integer
    Dim i As Long
    Dim StartBytes As Long
    
    ff = FreeFile
    StartBytes = FileLength + 1
    Open FileName For Binary Access Write As #ff
    Put #ff, StartBytes, Data
    Close #ff
End Sub

Public Function FileExist(sFileName As String) As Boolean
    If Len(Trim(Dir(sFileName))) <> 0 Then
        If UCase(Trim(Dir(sFileName))) = UCase(Trim(Right(Dir(sFileName), Len(Trim(Dir(sFileName)))))) Then
            FileExist = True
        End If
    End If
End Function

Public Sub AppendLog(Index As Integer, Data As String)
    If chkLog.Value = vbChecked Then
        If FileExist("gw" & Index & ".log") Then
            AppendFile "gw" & Index & ".log", Data & vbCrLf, FileLen("gw" & Index & ".log")
        Else
            AppendFile "gw" & Index & ".log", Data & vbCrLf
        End If
    End If
End Sub

Public Function NameByAddr(strAddr As String) As String
    On Error Resume Next
    Dim nRet As Long
    Dim lIP As Long
    Dim strHost As String * 255: Dim strtemp As String
    Dim hst As hostent

    If IsIP(strAddr) Then
        'lIP = MakeIP(strAddr)
        lIP = vbInet_aToN(strAddr)
        nRet = gethostbyaddr(lIP, 4, 2)

        If nRet <> 0 Then
            RtlMoveMemory hst, nRet, Len(hst)
            RtlMoveMemory ByVal strHost, hst.h_name, 255
            strtemp = strHost
            If InStr(strtemp, Chr(0)) <> 0 Then strtemp = Left(strtemp, InStr(strtemp, Chr(0)) - 1)
            strtemp = Trim(strtemp)
            NameByAddr = strtemp
        Else
            NameByAddr = "Host name not found"
            'MsgBox "Host name not found", , "9003"
            Exit Function
        End If
    Else
        NameByAddr = "Invalid IP address"
        'MsgBox "Invalid IP address", , "9002"
        Exit Function
    End If

    If Err.Number > 0 Then
        'MsgBox Err.Description, , Err.Number
        Err.Clear
    End If
End Function

Public Function IsIP(ByVal strIP As String) As Boolean
    On Error Resume Next
    Dim t As String: Dim s As String: Dim i As Integer
    s = strIP
    While InStr(s, ".") <> 0
        t = Left(s, InStr(s, ".") - 1)
        If IsNumeric(t) And Val(t) >= 0 And Val(t) <= 255 Then s = Mid(s, InStr(s, ".") + 1) _
    Else Exit Function
        i = i + 1
    Wend
    t = s
    If IsNumeric(t) And InStr(t, ".") = 0 And Len(t) = Len(Trim(Str(Val(t)))) And _
    Val(t) >= 0 And Val(t) <= 255 And strIP <> "255.255.255.255" And i = 3 Then IsIP = True
    If Err.Number > 0 Then
        MsgBox Err.Description, , Err.Number
        Err.Clear
    End If
End Function

Public Function vbInet_aToN(address As String) As Long
    vbInet_aToN = inet_addr(address)
End Function


