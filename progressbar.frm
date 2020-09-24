VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Loading..."
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading: "
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendChat("blahblah")


End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption + "|"
Label3.Caption = Val(Label3.Caption) + 1
If Label3.Caption = "100" Then
Timer1.Enabled = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Unload Me
Form2.Show

End If
End Sub
