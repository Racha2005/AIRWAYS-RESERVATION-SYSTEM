VERSION 5.00
Begin VB.MDIForm Form3 
   BackColor       =   &H8000000C&
   Caption         =   "AIRLINES"
   ClientHeight    =   10050
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   18660
   LinkTopic       =   "MDIForm1"
   Picture         =   "Airlines MDI Form.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnumas 
      Caption         =   "&Master"
      Begin VB.Menu mnutickd 
         Caption         =   "Ticket Details"
      End
      Begin VB.Menu mnupare 
         Caption         =   "Passenger's Reservation"
      End
   End
   Begin VB.Menu mnuloc 
      Caption         =   "&Login Credentials"
      Begin VB.Menu mnulogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnusignin 
         Caption         =   "Signin"
      End
   End
   Begin VB.Menu mnutrans 
      Caption         =   "&Transactions"
      Begin VB.Menu mnupay 
         Caption         =   "Payment"
      End
      Begin VB.Menu mnucan 
         Caption         =   "Cancellation"
      End
   End
   Begin VB.Menu mnust 
      Caption         =   "&Status"
      Begin VB.Menu mnuflst 
         Caption         =   "Flight Status"
      End
      Begin VB.Menu mnufood 
         Caption         =   "Food & Beverages"
      End
      Begin VB.Menu mnutose 
         Caption         =   "Total Seats Booked"
      End
   End
   Begin VB.Menu mnure 
      Caption         =   "&Report"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub MDIForm_Load()
Con.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\racha\Documents\JETBLUEAIRWAYS.MDB;Persist Security Info=False")
rs.Open "select * from airlines", Con, adOpenDynamic, adLockOptimistic
End Sub
Private Sub mnucan_Click()
Form9a.Show
End Sub
Private Sub mnuflst_Click()
Form9.Show
End Sub
Private Sub mnufood_Click()
Form5.Show
End Sub
Private Sub mnulogin_Click()
Form1.Show
End Sub
Private Sub mnupare_Click()
Form4.Show
End Sub
Private Sub mnusignin_click()
Form2.Show
End Sub
Private Sub mnupay_Click()
Form6.Show
End Sub
Private Sub mnutickd_Click()
Form7.Show
End Sub
Private Sub mnutose_Click()
Form8.Show
End Sub
