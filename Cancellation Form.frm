VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9a 
   Caption         =   "CANCELLATION"
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18675
   LinkTopic       =   "Form9"
   Picture         =   "Cancellation Form.frx":0000
   ScaleHeight     =   9735
   ScaleWidth      =   18675
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   13680
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   122486785
      CurrentDate     =   45356
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6480
      TabIndex        =   3
      Top             =   3960
      Width           =   7815
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
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      DataField       =   "pnrno"
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
      Height          =   735
      Left            =   10080
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CANCELLATION"
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
      Left            =   7080
      TabIndex        =   2
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Enter PNR No"
      DataField       =   "pnrno"
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
      Height          =   735
      Left            =   7080
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form9a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from cancel", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub CmdCancel_Click()
If Text1.Text = " " Or DTPicker1.Value Then
MsgBox "Ticket is cancelled!"
Else
Dim sql1
sql1 = "select * from RESERVATION_DETAILS WHERE [PNR NUMBER] LIKE" & Text1.Text & "%"
Adodc1.RecordSource = sql1
Adodc1.Refresh
DataGrid1.Visible = True
End If
End
End Sub

