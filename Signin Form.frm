VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "SIGNIN (New Customer)"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18735
   LinkTopic       =   "Form2"
   Picture         =   "Signin Form.frx":0000
   ScaleHeight     =   9765
   ScaleWidth      =   18735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BackColor       =   &H80000003&
      DataField       =   "add"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000003&
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   13
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000003&
      DataField       =   "pw"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   12
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "uname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   11
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   10
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton CmdSignin 
      Caption         =   "SIGNIN"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton CmdAddNew 
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SIGNIN"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from signin", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdSignin_Click()
rs.AddNew
rs("uname") = Text1.Text
rs("pw") = Text2.Text
rs("dob") = Text3.Text
rs("add") = Text4.Text
rs.Update
MsgBox "RECORD SAVED SUCCESSFULLY"
MsgBox "SIGNIN SUCCESSFULL"
Form1.Show
End Sub
Private Sub CmdExit_Click()
End
End Sub
Private Sub CmdAddNew_Click()
rs.AddNew
rs("uname") = Text1.Text
rs("pw") = Text2.Text
rs("dob") = Text3.Text
rs("add") = Text4.Text
rs.Update
MsgBox "RECORD CREATED SUCCESSFULLY"
End Sub
Private Sub CmdSave_Click()
rs("uname") = Text1.Text
rs("pw") = Text2.Text
rs("dob") = Text3.Text
rs("add") = Text4.Text
rs.Update
MsgBox "RECORD SAVED SUCCESSFULLY"
End Sub
Private Sub CmdClear_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text1.SetFocus
End Sub
Private Sub CmdDelete_Click()
rs("uname") = Text1.Text
rs("pw") = Text2.Text
rs("dob") = Text3.Text
rs("add") = Text4.Text
rs.Update
MsgBox "RECORD DELETED SUCCESSFULLY"
End Sub


