VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArrayBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Const BUFFER_LENGTH = 2 ^ 16

Private Blocking As Boolean

Private m_Length As Long
Private m_Data() As Byte

Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
    'Data = Left$(m_Data, m_Length)
    Data = Left$(StrConv(m_Data, vbUnicode), m_Length)
End Property

Public Property Let Data(vData As String)
    Clear
    Append vData
End Property

Public Sub Append(vData As String)
Dim lenData As Long, i As Long
Dim tmpArray() As Byte

    lenData = Len(vData)
    If lenData <> 0 Then
        Do While Blocking
            DoEvents
        Loop
        Blocking = True
        tmpArray = StrConv(vData, vbFromUnicode)
        For i = 0 To UBound(tmpArray)
            ReDim Preserve m_Data(m_Length)
            m_Data(m_Length) = tmpArray(i)
            m_Length = m_Length + 1
        Next i
        Blocking = False
    End If
End Sub

Public Sub Clear()
    ReDim m_Data(0)
    m_Length = 0
End Sub

Public Function GetString() As String
    Do While Blocking
        DoEvents
    Loop
    Blocking = True
    GetString = Data
    Clear
    Blocking = False
End Function

