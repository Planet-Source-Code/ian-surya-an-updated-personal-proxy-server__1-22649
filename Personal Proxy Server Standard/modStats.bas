Attribute VB_Name = "modStats"
Option Explicit

Public ConnectionRequest As New Collection

Public Sub AddConnectionStatistic(Socket As Winsock)
Dim CStat As Cconnection

    If IsInCollection(ConnectionRequest, Socket.RemoteHostIP) Then
        Set CStat = ConnectionRequest(Socket.RemoteHostIP)
        CStat.Stat_Connect_Count = CStat.Stat_Connect_Count + 1
    Else
        Set CStat = New Cconnection
        CStat.IPAddress = Socket.RemoteHostIP
        'CStat.HostName = Socket.RemoteHost
        'CStat.HostName = NameByAddr(Socket.RemoteHostIP)
        CStat.Key = CStat.IPAddress
        CStat.Stat_Connect_Count = CStat.Stat_Connect_Count + 1
        ConnectionRequest.Add CStat, CStat.Key
    End If
    
    Set CStat = Nothing
End Sub

Public Sub AddStatReceived(Socket As Winsock, BytesReceived As Long)
    If IsInCollection(ConnectionRequest, Socket.RemoteHostIP) Then
        ConnectionRequest(Socket.RemoteHostIP).Stat_Bytes_Received = ConnectionRequest(Socket.RemoteHostIP).Stat_Bytes_Received + BytesReceived
    End If
End Sub

Public Sub AddStatSent(Socket As Winsock, BytesSent As Long)
    If IsInCollection(ConnectionRequest, Socket.RemoteHostIP) Then
        ConnectionRequest(Socket.RemoteHostIP).Stat_Bytes_Sent = ConnectionRequest(Socket.RemoteHostIP).Stat_Bytes_Sent + BytesSent
    End If
End Sub


