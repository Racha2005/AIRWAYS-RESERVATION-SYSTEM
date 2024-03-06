VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "FOOD & BEVERAGES"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18780
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "Food And Beverages Form.frx":0000
   ScaleHeight     =   9765
   ScaleWidth      =   18780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClear 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   8160
      TabIndex        =   20
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "DELETE"
      Height          =   615
      Left            =   11040
      TabIndex        =   18
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   5040
      TabIndex        =   17
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton CmdAddNew 
      Caption         =   "ADD NEW"
      Height          =   615
      Left            =   2280
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   17160
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdCalculate 
      Caption         =   "CALCULATE"
      Height          =   615
      Left            =   12600
      TabIndex        =   14
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton CmdPayment 
      Caption         =   "PAYMENT"
      Height          =   615
      Left            =   7920
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton CmdAssign 
      Caption         =   "ASSIGN"
      Height          =   615
      Left            =   3600
      TabIndex        =   12
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox txtTotal 
      BackColor       =   &H80000003&
      DataField       =   "tprice"
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
      Left            =   16320
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtQuantity 
      BackColor       =   &H80000003&
      DataField       =   "quantity"
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
      Left            =   13320
      TabIndex        =   10
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   10320
      TabIndex        =   9
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ComboBox CboMeal 
      BackColor       =   &H80000003&
      DataField       =   "meal"
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
      Height          =   465
      Left            =   6840
      TabIndex        =   8
      Text            =   "CboMeal"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.ComboBox CboSeat 
      BackColor       =   &H80000003&
      DataField       =   "seat"
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
      Height          =   465
      Left            =   3600
      TabIndex        =   7
      Text            =   "CboSeat"
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ListBox lstcities 
      BackColor       =   &H80000003&
      DataField       =   "destiny"
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
      Height          =   1440
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FOOD AND BEVERAGES"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Caption         =   "Total Price"
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
      Left            =   16320
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "Quantity"
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
      Left            =   13320
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
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
      Left            =   10320
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Meal Preference"
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
      Left            =   6840
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Seat Location"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Destination City"
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
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from foodbever", Con, adOpenDynamic, adLockOptimistic
lstcities.Clear
lstcities.AddItem "Japan"
lstcities.AddItem "Los Angeles"
lstcities.AddItem "Amsterdam"
lstcities.AddItem "Malaysia"
lstcities.AddItem "California"
lstcities.AddItem "London"
lstcities.AddItem "Australia"
lstcities.AddItem "Germany"
lstcities.AddItem "San Francisco"
lstcities.AddItem "Greece"
lstcities.AddItem "France"
lstcities.AddItem "Russia"
lstcities.AddItem "Texas"
lstcities.AddItem "Hong Kong"
lstcities.ListIndex = 0
CboSeat.AddItem "Aisle"
CboSeat.AddItem "Middle"
CboSeat.AddItem "Window"
CboSeat.ListIndex = 0
CboMeal.AddItem "Chicken"
CboMeal.AddItem "Mystery Meat"
CboMeal.AddItem "Kosher"
CboMeal.AddItem "Vegetarian"
CboMeal.AddItem "Fruit Plate"
CboMeal.Text = "No Preference"
End Sub
Private Sub CmdAssign_Click()
Dim Message As String
Message = "Destination: " + lstcities.Text + vbCr
Message = Message + "Seat Location: " + CboSeat.Text + vbCr
Message = Message + "Meal: " + CboMeal.Text + vbCr
MsgBox Message, vbOKOnly + vbInformation, "Your Preference is:"
End Sub
Private Sub CmdCalculate_Click()
If txtPrice.Text = "" Or txtQuantity.Text = "" Or txtTotal = "" Then
MsgBox "TOTAL PRICE IS" & txtTotal.Text
End If
Total = txtPrice.Text * txtQuantity.Text
txtTotal.Text = Val(Total)
End Sub
Private Sub CmdExit_Click()
End
End Sub
Private Sub CmdAddNew_Click()
rs.AddNew
rs("destiny") = lstcities.Text
rs("seat") = CboSeat.Text
rs("meal") = CboMeal.Text
rs("price") = txtPrice.Text
rs("quantity") = txtQuantity.Text
rs("tprice") = txtTotal.Text
rs.Update
MsgBox "RECORD CREATED SUCCESSFULLY"
End Sub
Private Sub CmdSave_Click()
rs("destiny") = lstcities.Text
rs("seat") = CboSeat.Text
rs("meal") = CboMeal.Text
rs("price") = txtPrice.Text
rs("quantity") = txtQuantity.Text
rs("tprice") = txtTotal.Text
rs.Update
MsgBox "RECORD SAVED SUCCESSFULLY"
End Sub
Private Sub CmdClear_Click()
lstcities.Text = " "
CboSeat.Text = " "
CboMeal.Text = " "
txtPrice.Text = " "
txtQuantity.Text = " "
txtTotal.Text = " "
lstcities.SetFocus
End Sub
Private Sub CmdDelete_Click()
rs("destiny") = lstcities.Text
rs("seat") = CboSeat.Text
rs("meal") = CboMeal.Text
rs("price") = txtPrice.Text
rs("quantity") = txtQuantity.Text
rs("tprice") = txtTotal.Text
rs.Update
MsgBox "ACCOUNT DELETED SUCCESSFULLY"
End Sub
Private Sub CmdPayment_Click()
Form6.Show
End Sub
