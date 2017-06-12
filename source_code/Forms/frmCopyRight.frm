VERSION 5.00
Begin VB.Form frmCopyRight 
   Caption         =   "關於系統"
   ClientHeight    =   2280
   ClientLeft      =   8685
   ClientTop       =   7560
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   6915
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCopyRight5 
      Caption         =   "客製化 於日月光集團  洋鼎科技股份有限公司  應用"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblCopyRight4 
      Caption         =   "本軟體由 Imagine 設計"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lblCopyRight3 
      Caption         =   "Copyight @ 2013 Fatansy. All right reserved."
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblCopyRight2 
      Caption         =   "版本 6.1 (組件 7601: Service Pack 1)"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblCopyRight1 
      Caption         =   "NASConnection"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Image imgIcons 
      Height          =   480
      Left            =   240
      Picture         =   "frmCopyRight.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      X1              =   240
      X2              =   6720
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmCopyRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
