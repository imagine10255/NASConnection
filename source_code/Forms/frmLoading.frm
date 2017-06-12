VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YT NAS LINK"
   ClientHeight    =   3690
   ClientLeft      =   5310
   ClientTop       =   4365
   ClientWidth     =   8445
   Icon            =   "frmLoading.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLoading.frx":2A1AA
   ScaleHeight     =   3690
   ScaleWidth      =   8445
   Begin VB.Timer timEyeMinute 
      Interval        =   1000
      Left            =   7800
      Top             =   2880
   End
   Begin VB.CommandButton cmdOpenNAS 
      Caption         =   "開啟"
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2950
      Width           =   615
   End
   Begin VB.ComboBox cboOpenList 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.PictureBox picLogoutFalse 
      Height          =   255
      Left            =   1560
      Picture         =   "frmLoading.frx":37A1F
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLogoutTrue 
      Height          =   255
      Left            =   1200
      Picture         =   "frmLoading.frx":3CB1E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLoginDown 
      Height          =   255
      Left            =   840
      Picture         =   "frmLoading.frx":42802
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLoginTrue 
      Height          =   255
      Left            =   480
      Picture         =   "frmLoading.frx":47482
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLoginFalse 
      Height          =   255
      Left            =   120
      Picture         =   "frmLoading.frx":4C30C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtUserPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   0
      Top             =   1250
      Width           =   3135
   End
   Begin VB.Label txtWordTime 
      BackStyle       =   0  'Transparent
      Caption         =   "KK"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label txtWordTime 
      BackStyle       =   0  'Transparent
      Caption         =   "SS"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label txtWordTime 
      BackStyle       =   0  'Transparent
      Caption         =   "MM"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label txtWordTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "連續工作時間"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label txtStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label txtExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLoading.frx":50B8F
      Height          =   1575
      Left            =   5280
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Image cmdCancel 
      Height          =   480
      Left            =   5280
      Picture         =   "frmLoading.frx":50CD6
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Image imgBtnLogin 
      Height          =   480
      Left            =   3120
      Picture         =   "frmLoading.frx":55FF6
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Menu menuMyMenu 
      Caption         =   "功能表"
      Visible         =   0   'False
      Begin VB.Menu menuClose 
         Caption         =   "關閉選單"
      End
      Begin VB.Menu menuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath1"
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath2"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath3"
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath4"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath5"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath6"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath7"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath8"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath9"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu menuPath 
         Caption         =   "menuPath10"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu menuLine2 
         Caption         =   "menuLine2"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu menuSetup 
         Caption         =   "設置畫面"
      End
      Begin VB.Menu menuHelpSetup 
         Caption         =   "輔助設定"
      End
      Begin VB.Menu menuAboutSystem 
         Caption         =   "關於系統"
      End
      Begin VB.Menu menuExit 
         Caption         =   "結束程式"
      End
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Private Sub cmdOpenNAS_Click()
    Dim sNASPath As String
    sNASPath = Left(cboOpenList.Text, InStrRev(cboOpenList.Text, "\") - 1)
    Shell "Explorer.exe " & sNASPath, vbNormalFocus
End Sub


Private Sub Form_Load()
'初始化-SystemIcon
    
    'INI設定檔 路徑設定
    sPath = App.Path & "\NASConfig.ini"
    Call LoadEyeIni
       
    Call setTimeZero
       
    '禁止多開
    If App.PrevInstance Then Unload Me
    
    Call PrepareLoading

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'系統圖示 開啟功能表
    If Button = 2 Then PopupMenu menuMyMenu, 0 '滑鼠右鍵
    If Button = 1 Then
        frmLoading.Visible = True
        If mlngID <> 0 Then
            DeleteFromSystemTray mlngID
            mlngID = 0
        End If
    End If
  
  Debug.Print Button
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'滑鼠右鍵開啟功能表
    If Button = 2 Then
     PopupMenu menuMyMenu    '調用用戶自定義選單
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' Code Snippet
    Select Case UnloadMode
        Case Is = 0
            If mlngID = 0 Then
                mlngID = AddToSystemTray(Me.hWnd, WM_MOUSEMOVE, Me.Icon, "NASConnection")
            End If
            '''使用者從表單上的控制功能表中選取「關閉」指令。
            MsgBox "程式自動縮小於畫面右下時間列，需要時可再次呼叫", vbInformation
            frmLoading.Visible = False
            Cancel = -1 '取消X關閉
            
        Case Is = 1
            '''Unload 陳述式被程式碼呼叫。
        Case Is = 2
            '''目前 Microsoft Windows 作業環境任務結束。
        Case Is = 3
            '''Microsoft Windows 工作管理員正在關閉應用程式。
        Case Is = 4
            '''因為 MDI 表單正在關閉的緣故，MDI 子表單正在關閉。
        Case Is = 5
            '''表單因其擁有人關閉而關閉。
    End Select

    End Sub
Private Sub Form_Initialize()
'攔截程式多開
    If App.PrevInstance Then
        MsgBox "程式正在運行。請使用右下時間列呼叫", vbInformation
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'關閉主視窗時執行
    If mlngID <> 0 Then
        DeleteFromSystemTray mlngID
        mlngID = 0
    End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub menuSetup_Click()
    frmLoading.Visible = True
End Sub

Private Sub menuHelpSetup_Click()
    frmHelpSetup.Show
End Sub

Private Sub menuPath_Click(Index As Integer)
    Dim sNASPath As String
    sNASPath = Left(menuPath(Index).Caption, InStrRev(menuPath(Index).Caption, "\") - 1)
    Shell "Explorer.exe " & sNASPath, vbNormalFocus
End Sub

Private Sub f_menuExit_Click()
    End
End Sub

Private Sub imgBtnLogin_Click()

Dim sErrStr, sFunc As String

'-------------------------------------------------------------------------------
'映射網絡驅動器
'  Shell ("net use T: \\10.230.44.7\Public 密碼 /user:帳號)
'  Call AddConnection("\\10.230.44.7\Public", "Z:", 密碼, 帳號)
'---------------------------------------------
Dim aNASLink(10) As String
    
    If txtUserName.Text = "" Then
        MsgBox "UserName 不可空白"
        GoTo ExitHandler
    End If
    
    If txtUserPassword.Text = "" Then
        MsgBox "Password 不可空白"
        GoTo ExitHandler
    End If

    Call LoadNasIni(txtUserName.Text, txtUserPassword.Text)
    Call PrepareLoading
    
    
    
    
    GoTo ExitHandler
ErrorHandler:
    sErrStr = "Error " & Err.Number & ": " & Err.Description
    'Trace sFunc & ": " & sErrStr, 0
    MsgBox sErrStr, vbExclamation, sFunc
 
ExitHandler:
    On Error Resume Next
    
End Sub


Public Function PrepareLoading() As Boolean

Dim sErrStr, sFunc As String

    If cboOpenList.List(1) <> "" Then
        txtStatus.Caption = "連線狀態：已經登入"
        cboOpenList.Enabled = True
        cboOpenList.Text = cboOpenList.List(0)
        
        cmdOpenNAS.Enabled = True
   
        imgBtnLogin.Enabled = False
        imgBtnLogin.Picture = picLoginFalse.Picture
        cmdCancel.Enabled = True
        cmdCancel.Picture = picLogoutTrue.Picture
        txtUserName.Enabled = False
        txtUserPassword.Enabled = False
        
    Else
        txtStatus.Caption = "連線狀態：尚未登入"
        cboOpenList.Enabled = False
        cmdOpenNAS.Enabled = False
        imgBtnLogin.Enabled = True
        imgBtnLogin.Picture = picLoginTrue.Picture
        cmdCancel.Enabled = False
        cmdCancel.Picture = picLogoutFalse.Picture
        txtUserName.Enabled = True
        txtUserPassword.Enabled = True
    End If
    
    
    
    GoTo ExitHandler
ErrorHandler:
    sErrStr = "Error " & Err.Number & ": " & Err.Description
    'Trace sFunc & ": " & sErrStr, 0
    MsgBox sErrStr, vbExclamation, sFunc
 
ExitHandler:
    On Error Resume Next
    
End Function


Private Sub imgBtnLogin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgBtnLogin.Picture = picLoginDown.Picture
End Sub



Private Sub imgBtnLogin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgBtnLogin.Picture = picLoginTrue.Picture
End Sub


Private Sub menuExit_Click()
Dim Rtvl As Long
Rtvl = MsgBox("中斷程式將會中斷磁碟連線，您確定要中斷嗎？", vbYesNo, "確認視窗")
Select Case Rtvl
    Case vbYes
        cmdCancel_Click
            Unload frmLoading
        End
    Case vbNo
        
End Select

  

End Sub

Private Sub menuAboutSystem_Click()
    frmCopyRight.Show
End Sub




'讀取ini檔案，用來設定磁碟路徑與名稱
'EX: Call AddConnection("\\10.230.44.7\行政管理部-資訊", "X:", sUser_password, sUser_name)
Public Function LoadNasIni(sUser_name As String, sUser_password As String) As Boolean
    
Dim sErrStr, sFunc As String


    Dim svt As Long
    Dim loadd As String
    Dim sKey As String
    Dim sIniCont As String
    Dim aLinkList(10) As String
    Dim sTrueList As String
    Dim sFailList As String
    Dim iAddCount As Integer
    Dim i As Integer
    Dim aNASLink() As String
    
   
    '------------------------------
    '讀取 網路磁碟列表設定 [Config]
    '------------------------------
    iAddCount = 0
    For i = 0 To 9
        sIniCont = String(255, " ")
        sKey = "Folder" & Trim(Str(i + 1))
        svt = GetPrivateProfileString("Config", sKey, "", sIniCont, 256, sPath) '讀取INI檔案
        aLinkList(i) = Replace(Trim(sIniCont), Chr(0), "")
        If aLinkList(i) <> "" Then
            aNASLink = Split(aLinkList(i), ",")

            If AddConnection(aNASLink(1), aNASLink(0) & ":", sUser_password, sUser_name) Then
                sTrueList = sTrueList & "【" & aNASLink(0) & "】" & aNASLink(1) & Chr(13)
                cboOpenList.List(iAddCount) = aNASLink(0) & ":\" & Mid(aNASLink(1), InStrRev(aNASLink(1), "\") + 1)
                menuPath(iAddCount + 1).Enabled = True
                menuPath(iAddCount + 1).Visible = True
                menuPath(iAddCount + 1).Caption = cboOpenList.List(iAddCount)
                iAddCount = iAddCount + 1
            Else
                sFailList = sFailList & "【" & aNASLink(0) & "】" & aNASLink(1) & Chr(13)
            End If
        Else
            Exit For
        End If
    Next i
    
    If sTrueList <> "" Then
        sTrueList = "已成功連線" & Chr(13) & _
                    "--------------------------------------------------" & Chr(13) & _
                    sTrueList & Chr(13)
                    
        menuLine2.Enabled = True
        menuLine2.Visible = True
        menuLine2.Caption = "-"
    End If
    
    If sFailList <> "" Then
        sFailList = "連線失敗" & Chr(13) & _
                    "--------------------------------------------------" & Chr(13) & _
                    sFailList
    End If
    
    If sTrueList <> "" Or sFailList <> "" Then
        MsgBox sTrueList & sFailList, vbInformation
    End If
    

    GoTo ExitHandler
ErrorHandler:
    sErrStr = "Error " & Err.Number & ": " & Err.Description
    'Trace sFunc & ": " & sErrStr, 0
    MsgBox sErrStr, vbExclamation, sFunc
 
ExitHandler:
    On Error Resume Next

End Function

Public Function LoadEyeIni() As Boolean
    
Dim sErrStr, sFunc As String


    Dim svt As Long

    
    Dim sIniCont As String
    Dim aLinkList(10) As String
    Dim sTrueList As String
    Dim sFailList As String
    Dim iAddCount As Integer
    Dim i As Integer
    Dim aNASLink() As String
    
    
    
    
    '------------------------------------------------------------
    '讀取 眼睛保護提醒設定 [EyeProtection] (啟動狀態 與 倒數分鐘)
    '------------------------------------------------------------
    sIniCont = String(255, " ")
    svt = GetPrivateProfileString("EyeProtection", "Enabled", "", sIniCont, 256, sPath) '讀取INI檔案
    iEyeEnabled = Val(sIniCont)
    timEyeMinute = iEyeEnabled
    txtWordTitle.Visible = iEyeEnabled
    
    sIniCont = String(255, " ")
    svt = GetPrivateProfileString("EyeProtection", "Minute", "", sIniCont, 256, sPath) '讀取INI檔案
    iEyeMinute = Val(sIniCont)
    

    

    GoTo ExitHandler
ErrorHandler:
    sErrStr = "Error " & Err.Number & ": " & Err.Description
    'Trace sFunc & ": " & sErrStr, 0
    MsgBox sErrStr, vbExclamation, sFunc
 
ExitHandler:
    On Error Resume Next

End Function



Private Sub cmdCancel_Click()

Dim sErrStr, sFunc As String

'----------------------------------
'斷開網絡驅動器
'  Shell ("net use * /delete")
'  Call CancelConnection("Z:", True)
'----------------------------------
Dim iFailindex As Boolean
Dim i As Integer
Dim sNASPath As String
Dim sTrueList, sFailList As String

iFailindex = 0

    For i = 0 To cboOpenList.ListCount - 1
        sNASPath = Left(cboOpenList.List(0), InStrRev(cboOpenList.List(iFailindex), "\") - 1)
        If CancelConnection(sNASPath, True) Then
            sTrueList = sTrueList & "【" & cboOpenList.List(iFailindex) & "】" & Chr(13)
            cboOpenList.RemoveItem (iFailindex)
        Else
            sFailList = sFailList & "【" & cboOpenList.List(iFailindex) & "】" & Chr(13)
            iFailindex = iFailindex + 1
        End If
        
    Next i
    
    If sTrueList <> "" Then
        sTrueList = "已成功中斷" & Chr(13) & _
                    "--------------------------------------------------" & Chr(13) & _
                    sTrueList & Chr(13)
    End If
    
    If sFailList <> "" Then
        sFailList = "已成功中斷" & Chr(13) & _
                    "--------------------------------------------------" & Chr(13) & _
                    sFailList
    End If
    
    For i = 1 To 10
        menuPath(i).Enabled = False
        menuPath(i).Visible = False
        menuPath(i).Caption = ""
    Next i
        'menuLine2.Enabled = False
        menuLine2.Visible = False
        menuLine2.Caption = ""
        
    If sTrueList <> "" Or sFailList <> "" Then
        MsgBox sTrueList & sFailList, vbInformation
    End If
    Call PrepareLoading
    
    
    
    
    
    GoTo ExitHandler
ErrorHandler:
    sErrStr = "Error " & Err.Number & ": " & Err.Description
    'Trace sFunc & ": " & sErrStr, 0
    MsgBox sErrStr, vbExclamation, sFunc
 
ExitHandler:
    On Error Resume Next
    
End Sub

Private Sub timEyeMinute_Timer()
Dim NowSecond As Long

    If timEyeMinute = True Then    '若開關是開啟的，則第四欄位（1/100 秒）加一（interval=10ms）
        'txtWordTime(3).Caption = txtWordTime(3).Caption + 1
        'txtWordTime(2).Caption = txtWordTime(2).Caption + txtWordTime(3).Caption \ 100   '以下為進位過程
        txtWordTime(2).Caption = txtWordTime(2).Caption + 1  '以下為進位過程
        txtWordTime(3).Caption = txtWordTime(3).Caption Mod 100
        txtWordTime(1).Caption = txtWordTime(1).Caption + txtWordTime(2).Caption \ 60
        txtWordTime(2).Caption = txtWordTime(2).Caption Mod 60
        'txtWordTime(0).Caption = txtWordTime(0).Caption + txtWordTime(1).Caption \ 60
        'txtWordTime(1).Caption = txtWordTime(1).Caption Mod 60
        txtWordTitle = "連續工作時間：" & txtWordTime(1).Caption & " 分 " & txtWordTime(2).Caption & " 秒"
    
        If txtWordTime(1).Caption = Val(iEyeMinute) Then
            'Me.Show
            MsgBox "Hello！您所設定的 " & iEyeMinute & " 分鐘 時間已到", vbInformation
            Call setTimeZero
        End If
        
    End If
    'MsgBox "The end of the time"
End Sub

Public Function setTimeZero()
    txtWordTime(1) = 0
    txtWordTime(2) = 0
    txtWordTime(3) = 0
End Function

