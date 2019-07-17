VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Go back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton loginbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtpass 
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtuser 
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Admin Login Page"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password :-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Email :-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Form2.Hide
Form3.Show

End Sub

Private Sub Form_load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\LoginDB.mdb;Persist Security Info=False"
rs.Open "Select * from Logintab", con, adOpenDynamic, adLockPessimistic
End Sub





Private Sub loginbtn_Click()

rs.Close
rs.Open "Select * from Logintab where Email='" + txtuser.Text + "'and Password='" + txtpass.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
Form2.Hide

Form1.Show


Else
MsgBox "Email or Password is incorrect", vbInformation
End If

End Sub

