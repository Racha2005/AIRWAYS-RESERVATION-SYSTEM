VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form4 
   Caption         =   "PASSENGER'S RESERVATION"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18630
   LinkTopic       =   "Form3"
   Picture         =   "Passenger's Reservation Form.frx":0000
   ScaleHeight     =   10830
   ScaleWidth      =   18630
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "FLIGHT DETAILS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   19
      Top             =   6960
      Width           =   15495
      Begin MSACAL.Calendar Calendar3 
         Height          =   3015
         Left            =   8040
         TabIndex        =   36
         Top             =   240
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2024
         Month           =   2
         Day             =   28
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Height          =   495
         Left            =   4320
         TabIndex        =   33
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton CmdFood 
         Caption         =   "FOOD"
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
         Left            =   4320
         TabIndex        =   32
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton CmdFlightDateCalendar 
         Caption         =   "Show Calendar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   31
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00C0E0FF&
         DataField       =   "tprice"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   27
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtNoOfSeats 
         BackColor       =   &H00C0E0FF&
         DataField       =   "noseats"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   26
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtCost 
         BackColor       =   &H00C0E0FF&
         DataField       =   "flamt"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   25
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtFlightDate 
         BackColor       =   &H00C0E0FF&
         DataField       =   "fldt"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF00FF&
         Caption         =   "Total Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF00FF&
         Caption         =   "No Of Seats"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF00FF&
         Caption         =   "Flight Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF00FF&
         Caption         =   "Flight Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "BOOK NOW"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   15495
      Begin MSACAL.Calendar Calendar2 
         Height          =   3015
         Left            =   10680
         TabIndex        =   35
         Top             =   720
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2024
         Month           =   2
         Day             =   28
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3015
         Left            =   6240
         TabIndex        =   34
         Top             =   720
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2024
         Month           =   2
         Day             =   28
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton CmdBookNow 
         Caption         =   "BOOK NOW"
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
         Left            =   4320
         TabIndex        =   30
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CommandButton CmdDepartureDateCalendar 
         Caption         =   "Show Calendar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   29
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton CmdArrivalDateCalendar 
         Caption         =   "Show Calendar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   28
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtDepartureDate 
         BackColor       =   &H00C0C0FF&
         DataField       =   "depdt"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   18
         Top             =   3840
         Width           =   2415
      End
      Begin VB.OptionButton optFemale 
         BackColor       =   &H0080FFFF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H0080FFFF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtArrivalDate 
         BackColor       =   &H00C0C0FF&
         DataField       =   "arrdt"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   14
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00C0C0FF&
         DataField       =   "add"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00C0C0FF&
         DataField       =   "to"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtFrom 
         BackColor       =   &H00C0C0FF&
         DataField       =   "from"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtLname 
         BackColor       =   &H0000FF00&
         DataField       =   "lname"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtFname 
         BackColor       =   &H00C0C0FF&
         DataField       =   "fname"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000FF00&
         Caption         =   "Gender"
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   1375
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "Departure Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1375
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "Arrival Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Width           =   1375
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1375
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1375
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1375
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "(Last Name)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "(First Name)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   1250
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Passenger Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1375
      End
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSENGER'S RESERVATION"
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
      Left            =   1560
      TabIndex        =   37
      Top             =   360
      Width           =   12375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Paid As Integer
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from pres", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub Calendar1_Click()
txtArrivalDate.Text = Calendar1.Value
Calendar1.Visible = False
End Sub
Private Sub Calendar2_Click()
txtFlightDate.Text = Calendar2.Value
Calendar2.Visible = False
End Sub
Private Sub Calendar3_Click()
txtDepartureDate.Text = Calendar3.Value
Calendar3.Visible = False
End Sub
Private Sub CmdArrivalDateCalendar_Click()
Calendar1.Visible = True
Calendar3.Visible = False
End Sub
Private Sub CmdFlightDateCalendar_Click()
Calendar2.Visible = True
End Sub
Private Sub CmdDepartureDateCalendar_Click()
Calendar3.Visible = True
Calendar1.Visible = False
End Sub
Private Sub CmdFood_Click()
Form5.Show
End Sub
Private Sub CmdBookNow_click()
rs.AddNew
rs("fname") = txtFname.Text
rs("lname") = txtLname.Text
rs("from") = txtFrom.Text
rs("to") = txtTo.Text
rs("add") = txtAddress.Text
rs("arrdt") = txtArrivalDate.Text
rs("depdt") = txtDepartureDate.Text
rs("gender") = optMale.ToolTipText
rs("gender") = optFemale.ToolTipText
rs.Update
If optMale.Value = True Then
MsgBox "YOUR FLIGHT BOOKING IS SUCCESSFULL."
ElseIf optFemale.Value = True Then
MsgBox "YOUR FLIGHT BOOKING IS SUCCESSFULL."
End If
End Sub
Private Sub CmdCalculate_Click()
If txtFlightDate.Text = "" Or txtCost.Text = "" Or txtNoOfSeats = "" Or txtTotal = "" Then
MsgBox "TOTAL PRICE IS" & txtTotal.Text
End If
Total = txtCost.Text * txtNoOfSeats.Text
txtTotal.Text = Val(Total)
rs.AddNew
rs("fldt") = txtFlightDate.Text
rs("flamt") = txtCost.Text
rs("noseats") = txtNoOfSeats.Text
rs("tprice") = txtTotal.Text
rs.Update
End Sub
Private Sub SetState()
If optMale.Value = True Then
rs("gender") = True
Else
rs("gender") = False
End If
End Sub


