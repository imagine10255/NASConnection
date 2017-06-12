VERSION 5.00
Begin VB.Form frmHelpSetep 
   Caption         =   "輔助功能設定"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
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
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   670
         Visible         =   0   'False
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
Attribute VB_Name = "frmHelpSetep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEye_Click()
    If chkEye Then
        txtEyeMinute.Visible = True
    Else
        txtEyeMinute.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    frmLoading.Show
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    frmLoading.Show
End Sub


