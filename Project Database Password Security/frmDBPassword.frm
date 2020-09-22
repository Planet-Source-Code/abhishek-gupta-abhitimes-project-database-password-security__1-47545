VERSION 5.00
Begin VB.Form frmDBPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Password"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtP 
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Database password :"
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   1935
   End
End
Attribute VB_Name = "frmDBPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PW As String

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    PW = txtP
    Unload Me
    Call cn_open(PW)    'procedure for opening connection with password as argument
End Sub

Private Sub Form_Load()
    frmDBPassword.Height = 2000
    frmDBPassword.Width = 3000
    frmDBPassword.Top = (Screen.Height - frmDBPassword.Height) / 2.5
    frmDBPassword.Left = (Screen.Width - frmDBPassword.Width) / 2
End Sub
