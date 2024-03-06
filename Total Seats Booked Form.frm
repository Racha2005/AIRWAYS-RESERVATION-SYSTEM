VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "TOTAL SEATS BOOKED"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18690
   LinkTopic       =   "Form7"
   Picture         =   "Total Seats Booked Form.frx":0000
   ScaleHeight     =   10020
   ScaleWidth      =   18690
   StartUpPosition =   3  'Windows Default
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
      Left            =   10320
      TabIndex        =   20
      Top             =   8880
      Width           =   2295
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
      TabIndex        =   19
      Top             =   8880
      Width           =   1815
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
      Left            =   1320
      TabIndex        =   18
      Top             =   8880
      Width           =   2535
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
      Left            =   7560
      TabIndex        =   17
      Top             =   7920
      Width           =   2055
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
      Left            =   7560
      TabIndex        =   16
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton CmdCalculate 
      Caption         =   "CALCULATE"
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
      Left            =   4440
      TabIndex        =   15
      Top             =   7920
      Width           =   2535
   End
   Begin VB.CommandButton CmdFlightStatus 
      Caption         =   "FLIGHT STATUS"
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
      Left            =   960
      TabIndex        =   14
      Top             =   7920
      Width           =   2895
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000003&
      DataField       =   "tosb"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000003&
      DataField       =   "mica"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000003&
      DataField       =   "baecoca"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000003&
      DataField       =   "ecoca"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000003&
      DataField       =   "date"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000003&
      DataField       =   "flno"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "pname"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL SEATS BOOKED"
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
      Height          =   1095
      Left            =   3360
      TabIndex        =   21
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Total Seats Booked"
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
      Left            =   1320
      TabIndex        =   12
      Top             =   6720
      Width           =   3070
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mixed Cabin"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   5880
      Width           =   3075
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Basic Economy Cabin"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   5040
      Width           =   3070
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Economy Cabin"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   3840
      Width           =   3070
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Date"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   3070
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Flight No"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   3070
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Passenger Name"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3070
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub CmdFlightStatus_Click()
Form9.Show
End Sub
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from totseabo", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdCalculate_Click()
If (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "") Then
MsgBox "Fields should not be left Blank"
End If
Seats = Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text)
Text7.Text = Seats
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdAddNew_Click()
rs.AddNew
rs("pname") = Text1.Text
rs("flno") = Text2.Text
rs("date") = Text3.Text
rs("ecoca") = Text4.Text
rs("baecoca") = Text5.Text
rs("mica") = Text6.Text
rs("tosb") = Text7.Text
rs.Update
MsgBox "RECORD ADDED SUCCESSFULLY"
End Sub
Private Sub CmdSave_Click()
rs("pname") = Text1.Text
rs("flno") = Text2.Text
rs("date") = Text3.Text
rs("ecoca") = Text4.Text
rs("baecoca") = Text5.Text
rs("mica") = Text6.Text
rs("tosb") = Text7.Text
rs.Update
MsgBox "RECORD SAVED SUCCESSFULLY"
End Sub
Private Sub CmdClear_Click() '
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text1.SetFocus
End Sub
Private Sub CmdDelete_Click()
rs("pname") = Text1.Text
rs("flno") = Text2.Text
rs("date") = Text3.Text
rs("ecoca") = Text4.Text
rs("baecoca") = Text5.Text
rs("mica") = Text6.Text
rs("tosb") = Text7.Text
rs.Update
MsgBox "RECORD DELETED SUCCESSFULLY"
End Sub
