Attribute VB_Name = "modSystemICON"
Option Explicit

Private Declare Function Shell_NotifyIconA Lib "SHELL32.DLL" _
                (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Private mlngID As Long
Private mcolNID As Collection

'----------------------------------------------------------------------------------
Public Function AddToSystemTray(ByVal hWnd As Long, _
                                ByVal vlngCallbackMessage As Long, _
                                ByVal vipdIcon As IPictureDisp, _
                                ByVal vstrTip As String) As Long

    mlngID = mlngID + 1
   
    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = hWnd
        .uID = mlngID
        .uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
        .uCallbackMessage = vlngCallbackMessage
        .hIcon = CLng(vipdIcon)
        .szTip = vstrTip & vbNullChar
    End With

    If mcolNID Is Nothing Then Set mcolNID = New Collection

    mcolNID.Add hWnd, CStr(mlngID)

    Shell_NotifyIconA NIM_ADD, nidTemp
   
    AddToSystemTray = mlngID

End Function

Public Sub ModifySystemTrayMessage(ByVal vlngID As Long, _
                                   ByVal vlngCallbackMessage As Long)

    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = mcolNID(CStr(vlngID))
        .uID = vlngID
        .uFlags = NIF_MESSAGE
        .uCallbackMessage = vlngCallbackMessage
        .hIcon = 0
        .szTip = vbNullChar
    End With

    Shell_NotifyIconA NIM_MODIFY, nidTemp

End Sub

Public Sub ModifySystemTrayIcon(ByVal vlngID As Long, _
                                ByVal vipdIcon As IPictureDisp)

    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = mcolNID(CStr(vlngID))
        .uID = vlngID
        .uFlags = NIF_ICON
        .uCallbackMessage = 0
        .hIcon = CLng(vipdIcon)
        .szTip = vbNullChar
    End With

    Shell_NotifyIconA NIM_MODIFY, nidTemp

End Sub
Public Sub ModifySystemTrayTip(ByVal vlngID As Long, _
                               ByVal vstrTip As String)

    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = mcolNID(CStr(vlngID))
        .uID = vlngID
        .uFlags = NIF_TIP
        .uCallbackMessage = 0
        .hIcon = 0
        .szTip = vstrTip & vbNullChar
    End With

    Shell_NotifyIconA NIM_MODIFY, nidTemp

End Sub

Public Sub DeleteFromSystemTray(ByVal vlngID As Long)

    Dim nidTemp As NOTIFYICONDATA
   
    With nidTemp
        .cbSize = Len(nidTemp)
        .hWnd = mcolNID(CStr(vlngID))
        .uID = vlngID
        .uFlags = NIF_MESSAGE + NIF_ICON + NIF_TIP
    End With

    Shell_NotifyIconA NIM_DELETE, nidTemp

End Sub




