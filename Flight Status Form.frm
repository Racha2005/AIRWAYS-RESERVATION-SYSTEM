VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9 
   Caption         =   "FLIGHT STATUS"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18645
   LinkTopic       =   "Form8"
   Picture         =   "Flight Status Form.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   18645
   StartUpPosition =   3  'Windows Default
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
      Left            =   8760
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "SHOW"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "date"
      DataSource      =   "Adodc2"
      Height          =   735
      Left            =   10440
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
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
      Format          =   123207681
      CurrentDate     =   45353
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "pnrno"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FLIGHT STATUS"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Enter PNR No"
      DataField       =   "pnr no"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from flstatus", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdShow_Click()
If Text1.Text = " " Or DTPicker1.Value Then
MsgBox "The Status is Confirmed!"
Form9.Show
Else
Dim sql1
sql1 = "select * from RESERVATION_DETAILS WHERE [PNR NUMBER] LIKE" & Text1.Text & "%"
Adodc1.RecordSource = sql1
Adodc1.Refresh
DataGrid1.Visible = True
End If
End Sub
Private Sub CmdCancel_Click()
Form9a.Show
End Sub
