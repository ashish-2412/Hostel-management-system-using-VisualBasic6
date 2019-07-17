VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "1"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17775
   FillColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   12375
   ScaleWidth      =   17775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton logout 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   14160
      TabIndex        =   40
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fees paid!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton findbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Previousbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   11280
      Width           =   2415
   End
   Begin VB.CommandButton Nextbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   11280
      Width           =   2415
   End
   Begin VB.CommandButton Lastbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   11280
      Width           =   2415
   End
   Begin VB.CommandButton Firstbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   11280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   32
      Top             =   2160
      Width           =   4815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Non- Veg"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   31
      Top             =   9240
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Veg"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   9240
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   29
      Top             =   6480
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14640
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "Form1.frx":9E3FE
      Left            =   5880
      List            =   "Form1.frx":9E400
      TabIndex        =   27
      Text            =   "Select Department"
      Top             =   7200
      Width           =   4935
   End
   Begin VB.CommandButton uploadbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton savebtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10200
      Width           =   1815
   End
   Begin VB.CommandButton updatebtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10200
      Width           =   1935
   End
   Begin VB.CommandButton deletebtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10200
      Width           =   1815
   End
   Begin VB.CommandButton addnewbtn 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10200
      Width           =   1815
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   1440
      Width           =   4815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Double room"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   17
      Top             =   8640
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Single room"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      ToolTipText     =   "Single"
      Top             =   8640
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   13680
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   7800
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   5760
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   5040
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   4320
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hostel Management System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   41
      Top             =   240
      Width           =   11895
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fees due ="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12360
      TabIndex        =   38
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   28
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profile Image "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13830
      TabIndex        =   21
      Top             =   3840
      Width           =   1740
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Meal Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   18
      Top             =   9120
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Occupancy Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian's Phone number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian's Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of birth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim confirm As Integer




Private Sub addnewbtn_Click()
rs.AddNew
clear
End Sub

Private Sub Command4_Click()

End Sub





Private Sub Command1_Click()
Form4.Show

End Sub

Private Sub Command2_Click()
rs.Fields("Fees_Due").Value = 0
rs.Update
display
End Sub

Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to delete the Student Profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "Message"
rs.Update
refreshdata
Else
MsgBox "Profile Not Deleted ..!!", vbInformation, "Message"
End If

End Sub
Sub refreshdata()
rs.Close
rs.Open "Select * from Boarders", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub


Private Sub findbtn_Click()
rs.Close
rs.Open "Select * from Boarders where ID='" + Text0.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found ..!!", vbInformation
End If

End Sub
Sub reload()
rs.Close
rs.Open "Select * from Boarders", con, adOpenDynamic, adLockPessimistic
End Sub


Private Sub Firstbtn_Click()
rs.MoveFirst
display
End Sub

Private Sub Form_load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB\Database.mdb;Persist Security Info=False"
rs.Open "Select * from Boarders", con, adOpenDynamic, adLockPessimistic
Combo1.AddItem "Information Teachnology"
Combo1.AddItem "Computer Science and Engineering"
Combo1.AddItem "Electrical and Communication Engineering"
Combo1.AddItem "Mechanical Engineering"
display
End Sub
Sub display()
Text10.Text = rs!Fees_Due
Text0.Text = rs!ID
Text1.Text = rs!Name
Text2.Text = rs!Phone
Text3.Text = rs!DOB
Text4.Text = rs!Email
Text5.Text = rs!Guardians_Name
Text6.Text = rs!Guardians_Phone
Text7.Text = rs!Students_Address
Text8.Text = rs!Dept
If rs!Occupancy_Type = "Single room" Then
Option1.Value = True
Else
Option2.Value = True
End If
If rs!Meal_type = "Veg" Then
Check1.Value = 1
ElseIf rs!Meal_type = "Non veg" Then
Check2.Value = 1
Else
Check1.Value = 1
Check2.Value = 1
End If
Picture1.Picture = LoadPicture(rs!Image)
End Sub
Sub clear()
Text0.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Combo1.Text = "Select Semester"
Picture1.Picture = LoadPicture("")
Option1.Value = False
Option2.Value = False
Check1.Value = 0
Check2.Value = 0



End Sub














Private Sub Lastbtn_Click()
rs.MoveLast
display

End Sub

Private Sub logout_Click()
Form1.Hide
Form3.Show

End Sub

Private Sub Nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If

End Sub



Private Sub Previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If

End Sub

Private Sub savebtn_Click()


rs.Fields("ID").Value = Text0.Text
rs.Fields("Name").Value = Text1.Text

rs.Fields("Phone").Value = Text2.Text
rs.Fields("Email").Value = Text4.Text
rs.Fields("Guardians_Name").Value = Text5.Text
rs.Fields("Guardians_Phone").Value = Text6.Text
rs.Fields("Students_Address").Value = Text7.Text
rs.Fields("Dept").Value = Combo1.Text
rs.Fields("Year").Value = Text8.Text
rs.Fields("DOB").Value = Text3.Text
If Option1.Value = True Then
rs.Fields("Occupancy_Type").Value = Option1.Caption
Else
rs.Fields("Occupancy_Type").Value = Option2.Caption
End If
If Check1.Value = 1 Then
rs.Fields("Meal_Type").Value = Check1.Caption
Else
rs.Fields("Meal_Type").Value = Check2.Caption
End If

rs.Fields("Image").Value = str

Dim fees As Integer
fees = 0
If rs.Fields("Occupancy_Type").Value = "Single room" Then
fees = fees + 3000
Else
fees = fees + 1700
End If
If rs.Fields("Meal_Type").Value = "Veg" Then
fees = fees + 2000
Else
fees = fees + 2500
End If
rs.Fields("Fees_Due").Value = fees










MsgBox "Data is saved successfully! ", vbInformation
rs.Update


End Sub







Private Sub updatebtn_Click()
rs.Fields("ID").Value = Text0.Text
rs.Fields("Name").Value = Text1.Text

rs.Fields("Phone").Value = Text2.Text
rs.Fields("Email").Value = Text4.Text
rs.Fields("Guardians_Name").Value = Text5.Text
rs.Fields("Guardians_Phone").Value = Text6.Text
rs.Fields("Students_Address").Value = Text7.Text
rs.Fields("Dept").Value = Combo1.Text
rs.Fields("Year").Value = Text8.Text
rs.Fields("DOB").Value = Text3.Text
If Option1.Value = True Then
rs.Fields("Occupancy_Type").Value = Option1.Caption
Else
rs.Fields("Occupancy_Type").Value = Option2.Caption
End If
If Check1.Value = 1 Then
rs.Fields("Meal_Type").Value = Check1.Caption
Else
rs.Fields("Meal_Type").Value = Check2.Caption
End If

rs.Fields("Fees_Due").Value = Text10.Text

MsgBox "Data is saved successfully! ", vbInformation
rs.Update


End Sub

Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)

End Sub
