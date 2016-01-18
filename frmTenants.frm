VERSION 5.00
Begin VB.Form frmTenants 
   Caption         =   "Tenants Registration Records"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13080
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
   ScaleHeight     =   7965
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton lblTenant 
      Caption         =   "Edit Tenant"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Tenant"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quick Details:"
      Height          =   5055
      Left            =   7560
      TabIndex        =   5
      Top             =   2760
      Width           =   5055
      Begin VB.Label lblAmountDue 
         Caption         =   "Amount Due"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Label lblAmountPaid 
         Caption         =   "Amount Paid"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label lblDateAdmitted 
         Caption         =   "Date Admitted"
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label lblHouse 
         Caption         =   "House No."
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label lblMobile 
         Caption         =   "Mobile"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.ListBox lstTenants 
      Height          =   4470
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   6255
   End
   Begin VB.CommandButton cmdNewTenant 
      Caption         =   "Add New Tenant"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   5055
      Left            =   480
      Top             =   2760
      Width           =   6855
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   480
      Top             =   1680
      Width           =   12135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tenants Registration Records"
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
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   12015
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   480
      Top             =   240
      Width           =   12135
   End
End
Attribute VB_Name = "frmTenants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub cmdDelete_Click()
    If lstTenants.Text = "" Then
        MsgBox "Please select a tenant first to proceed!", vbCritical, App.Title
        Exit Sub
    ElseIf MsgBox("Are you sure you wish to delete this tenant?", vbCritical + vbYesNo, App.Title) = vbYes Then
        con.Execute "delete from tenants where fullnames = '" & lstTenants.Text & "'"
        MsgBox "Tenant deleted Succesfully", vbInformation, App.Title
    End If
    Exit Sub
    Load_AllTenants
End Sub

Private Sub cmdNewTenant_Click()
    frmNewTenant.Show
End Sub

Private Sub cmdRefresh_Click()
    Load_AllTenants
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\RealEstate.mdb;"
    con.Open
    Load_AllTenants
End Sub

Private Sub Load_AllTenants()
    lstTenants.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants", con, adOpenKeyset, adLockOptimistic
    Do Until Rs.EOF
        lstTenants.AddItem Rs!fullnames
        Rs.MoveNext
    Loop
    Rs.Close
End Sub

Private Sub lblTenant_Click()
    If lstTenants.Text = "" Then
        MsgBox "Please select a tenant first to proceed!", vbCritical, App.Title
        Exit Sub
    End If
    frmNewTenant.Show
End Sub

Private Sub lstTenants_Click()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants WHERE FullNames = '" & lstTenants.Text & "'", con, adOpenKeyset, adLockOptimistic
    lblMobile.Caption = "Mobile: " & Rs!mobile
    lblHouse.Caption = "House No: " & Rs!houseno
    lblDateAdmitted.Caption = "Admitted on: " & Rs!datein
    lblAmountPaid.Caption = "Amount Paid: " & Rs!amountpaid
    lblAmountDue.Caption = "Amount Due: " & Rs!amountdue
    Rs.Close
End Sub
