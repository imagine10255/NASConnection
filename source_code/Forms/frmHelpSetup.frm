VERSION 5.00
Begin VB.Form frmHelpSetup 
   Caption         =   "輔助功能設定"
   ClientHeight    =   1920
   ClientLeft      =   8325
   ClientTop       =   5595
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   2595
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "眼睛保護功能"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox chkEye 
         Caption         =   "啟動"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtEyeMinute 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   670
         Width           =   615
      End
      Begin VB.Label lblEyeMinute 
         Caption         =   "倒數時間(分鐘)"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHelpSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEye_Click()
    If chkEye Then
        txtEyeMinute.Enabled = True
    Else
        txtEyeMinute.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    chkEye = iEyeEnabled
    txtEyeMinute.Text = iEyeMinute
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    iEyeEnabled = chkEye
    iEyeMinute = txtEyeMinute.Text
    frmLoading.txtWordTitle.Visible = iEyeEnabled
    frmLoading.timEyeMinute.Enabled = iEyeEnabled
    
    If chkEye = False Then
        txtEyeMinute.Enabled = False
        Call frmLoading.setTimeZero
    End If
    
    Me.Hide
End Sub


Private Sub Form_Load()
    chkEye = iEyeEnabled
    
    If chkEye Then
        txtEyeMinute.Enabled = True
    Else
        txtEyeMinute.Enabled = False
    End If
    txtEyeMinute.Text = iEyeMinute
End Sub

