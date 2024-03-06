VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "TICKET DETAILS"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18690
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form6"
   Picture         =   "Ticket Details Form.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   18690
   StartUpPosition =   3  'Windows Default
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
      Left            =   9120
      TabIndex        =   23
      Top             =   7680
      Width           =   2175
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
      Left            =   12600
      TabIndex        =   22
      Top             =   7680
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
      Left            =   5880
      TabIndex        =   21
      Top             =   7680
      Width           =   1935
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
      Left            =   2160
      TabIndex        =   20
      Top             =   7680
      Width           =   2415
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
      Left            =   9120
      TabIndex        =   18
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton CmdTotalSeats 
      Caption         =   "TOTAL SEATS"
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
      Left            =   5280
      TabIndex        =   17
      Top             =   6240
      Width           =   3135
   End
   Begin VB.CommandButton CmdGetData 
      Caption         =   "GET DATA"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   6240
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000003&
      DataField       =   "seatsav"
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
      Left            =   9000
      TabIndex        =   15
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000003&
      DataField       =   "price"
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
      Left            =   9000
      TabIndex        =   14
      Top             =   3960
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H80000003&
      DataField       =   "cabclss"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9000
      TabIndex        =   13
      Top             =   3000
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000003&
      DataField       =   "destiny"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9000
      TabIndex        =   12
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text2 
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
      Left            =   3480
      TabIndex        =   11
      Top             =   4920
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000003&
      DataField       =   "flno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3480
      TabIndex        =   10
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "pno"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "arrdt"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   3480
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   122224641
      CurrentDate     =   45352
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TICKET DETAILS"
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
      Left            =   3240
      TabIndex        =   19
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      Caption         =   "Seats Available"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   "Price"
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
      Left            =   6240
      TabIndex        =   6
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Cabin Class"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Destination"
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
      Left            =   6240
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
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
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
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
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Passport No"
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
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Arrival Date"
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
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Combo1.AddItem "F101"
Combo1.AddItem "F102"
Combo1.AddItem "F103"
Combo1.AddItem "F104"
Combo2.AddItem "Himachal Pradesh"
Combo2.AddItem "Lucknow"
Combo2.AddItem "Antartica"
Combo2.AddItem "Dubai"
Combo3.AddItem "Economy"
Combo3.AddItem "Business"
Combo3.AddItem "Premium Economy"
Combo3.AddItem "Basic Economy"
Combo3.AddItem "Mixed"
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from tdetails", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdGetData_Click()
If Combo1 = "F101" Then
Text1.Text = "K2980"
Text2.Text = "Niharika"
Text3.Text = "1300"
Text4.Text = "50"
End If
If Combo1 = "F102" Then
Text1.Text = "K2990"
Text2.Text = "Akanksha"
Text3.Text = "5000"
Text4.Text = "80"
End If
If Combo1 = "F103" Then
Text1.Text = "K3000"
Text2.Text = "Krisha"
Text3.Text = "2000"
Text4.Text = "100"
End If
If Combo1 = "F104" Then
Text1.Text = "K3010"
Text2.Text = "Kriyan"
Text3.Text = "3000"
Text4.Text = "21"
End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdTotalSeats_Click()
Form8.Show
End Sub
Private Sub CmdAddNew_Click()
rs.AddNew
rs("arrdt") = DTPicker1.Value
rs("pno") = Text1.Text
rs("flno") = Combo1.Text
rs("pname") = Text2.Text
rs("destiny") = Combo2.Text
rs("cabclss") = Combo3.Text
rs("price") = Text3.Text
rs("seatsav") = Text4.Text
rs.Update
MsgBox "RECORD ADDED SUCCESSFULLY"
End Sub
Private Sub CmdSave_Click()
rs("arrdt") = DTPicker1.Value
rs("pno") = Text1.Text
rs("flno") = Combo1.Text
rs("pname") = Text2.Text
rs("destiny") = Combo2.Text
rs("cabclss") = Combo3.Text
rs("price") = Text3.Text
rs("seatsav") = Text4.Text
rs.Update
MsgBox "RECORD SAVED SUCCESSFULLY"
End Sub
Private Sub CmdClear_Click()
Text1.Text = " "
Combo1.Text = " "
Text2.Text = " "
Combo2.Text = " "
Combo3.Text = " "
Text3.Text = " "
Text4.Text = " "
Text1.SetFocus
End Sub
Private Sub CmdDelete_Click()
rs("arrdt") = DTPicker1.Value
rs("pno") = Text1.Text
rs("flno") = Combo1.Text
rs("pname") = Text2.Text
rs("destiny") = Combo2.Text
rs("cabclss") = Combo3.Text
rs("price") = Text3.Text
rs("seatsav") = Text4.Text
rs.Update
MsgBox "RECORD DELETED SUCCESSFULLY"
End Sub
