VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LOGIN"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18705
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Login Form.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   18705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H80000003&
      DataField       =   "pw"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6360
      TabIndex        =   7
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "uname"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "CANCEL"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   6240
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
      Left            =   5880
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "LOGIN"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "Password"
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      Caption         =   "Username"
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
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
      Height          =   1335
      Left            =   7560
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from login", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdLogin_Click()
rs.AddNew
rs("uname") = Text1.Text
rs("pw") = Text2.Text
rs.Update
MsgBox "LOGIN SUCCESSFUL"
Form4.Show
End Sub
Private Sub CmdSignin_Click()
Form1.Hide
Form2.Show
End Sub
Private Sub CmdCancel_Click()
End
End Sub


