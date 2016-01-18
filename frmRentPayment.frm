VERSION 5.00
Begin VB.Form frmRentPayment 
   Caption         =   "Rent Payment Records"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "News Gothic"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstHouses 
      Height          =   3840
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rent Payment"
      Height          =   5055
      Left            =   5640
      TabIndex        =   1
      Top             =   1440
      Width           =   5655
      Begin VB.Label lblFullDetails 
         Caption         =   "FullNames"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile Number"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label lblDateAdmitted 
         Caption         =   "Date Admitted"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lblAmountPaid 
         Caption         =   "Amount Paid"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label lblAmountDue 
         Caption         =   "Amount Due"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   4200
         Width           =   4335
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Occupied Houses Currently"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      Height          =   5055
      Left            =   360
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   360
      Top             =   240
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rent Payment Records"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   10815
   End
End
Attribute VB_Name = "frmRentPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub Load_AllHouses()
    lstHouses.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstHouses.AddItem "House No. " & Rs!houseno
        Rs.MoveNext
    Loop
    Rs.Close
End Sub

Private Sub lstHouses_Click()
Dim houzno As String
houzno = Right(lstHouses.Text, 2)
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants WHERE houseno = '" & houzno & "'", con, adOpenKeyset, adLockOptimistic
    lblFullDetails.Caption = Rs!fullnames
    lblMobile.Caption = "Mobile No: " & Rs!mobile
    lblDateAdmitted.Caption = "Admitted on: " & Rs!datein
    lblAmountPaid.Caption = "Amount Paid: " & Rs!amountpaid
    lblAmountDue.Caption = "Amount Due: " & Rs!amountdue
    Rs.Close
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\RealEstate.mdb;"
    con.Open
    Load_AllHouses
End Sub
