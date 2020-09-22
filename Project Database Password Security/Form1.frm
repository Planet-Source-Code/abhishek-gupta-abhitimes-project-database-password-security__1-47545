VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DEMO Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd1 
      Caption         =   "Show"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   1860
      Width           =   1215
   End
   Begin VB.ListBox l3 
      Height          =   840
      Left            =   3120
      TabIndex        =   5
      Top             =   780
      Width           =   975
   End
   Begin VB.ListBox l2 
      Height          =   840
      Left            =   1800
      TabIndex        =   4
      Top             =   780
      Width           =   975
   End
   Begin VB.ListBox l1 
      Height          =   840
      Left            =   420
      TabIndex        =   3
      Top             =   780
      Width           =   975
   End
   Begin VB.Label lb1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   2520
      Width           =   2355
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Category"
      Height          =   195
      Left            =   3180
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Name"
      Height          =   195
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "iNo"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmd1_Click()
    If l1 = "" Or l2 = "" Or l3 = "" Then
        MsgBox "Select all choices.", vbInformation + vbOKOnly
    Else
        lb1 = l1.Text & ",  " & l2.Text & ",  " & l3.Text
    End If
End Sub

Private Sub Form_Load()
    rs.Open "select * from tab", cn, 3, 3
    rs.MoveFirst
    Do While Not rs.EOF
        l1.AddItem rs(0)
        l2.AddItem rs(1)
        l3.AddItem rs(2)
        rs.MoveNext
    Loop
End Sub
