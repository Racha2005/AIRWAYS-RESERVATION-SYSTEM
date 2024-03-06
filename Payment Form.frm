VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   Caption         =   "PAYMENT DETAILS"
   ClientHeight    =   9900
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
   LinkTopic       =   "Form5"
   Picture         =   "Payment Form.frx":0000
   ScaleHeight     =   9900
   ScaleWidth      =   18705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdTicketDetails 
      Caption         =   "TICKET DETAILS"
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
      Left            =   13320
      TabIndex        =   14
      Top             =   7560
      Width           =   3255
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
      Left            =   10560
      TabIndex        =   13
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton CmdProceed 
      Caption         =   "PROCEED"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   7560
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "cdexdt"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11280
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      _Version        =   393216
      CalendarBackColor=   -2147483645
      CalendarTitleBackColor=   12648447
      CustomFormat    =   "mm-yyyy"
      Format          =   245366787
      CurrentDate     =   45352
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000003&
      DataField       =   "mobno"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11280
      TabIndex        =   9
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000003&
      DataField       =   "cvvno"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11280
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000003&
      DataField       =   "cdhona"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11280
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "cdno"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   11280
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000003&
      Height          =   615
      Left            =   11280
      TabIndex        =   15
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT DETAILS"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Total Price"
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Mobile No"
      Height          =   615
      Left            =   6600
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "CVV No"
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Card Holder Name"
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Card Expiry Date"
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Card No"
      Height          =   615
      Left            =   6600
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from payment", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdProceed_Click()
rs.AddNew
rs("cdno") = Text1.Text
rs("cdexdt") = DTPicker1.Value
rs("cdhona") = Text2.Text
rs("cvvno") = Text3.Text
rs("mobno") = Text4.Text
rs("tprice") = Label7.Caption
rs.Update
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Label7.Caption = ""
MsgBox "OTP SENT TO YOUR MOBILE! PLEASE ENTER IT."
MsgBox "PAYMENT SUCCESSFULL"
End Sub
Private Sub Label7_Click()
Label7.Caption = Val(Form4.txtTotal) + Val(Form5.txtTotal)
End Sub
Private Sub CmdExit_Click()
End
End Sub
Private Sub CmdTicketDetails_Click()
Form7.Show
End Sub
