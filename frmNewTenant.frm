VERSION 5.00
Begin VB.Form frmNewTenant 
   Caption         =   "Tenants Registration Records"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13140
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
   ScaleHeight     =   7485
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Quick Options:"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   19
      Top             =   5880
      Width           =   12135
      Begin VB.CommandButton cmdSave 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         Height          =   615
         Left            =   9840
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   6840
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Register a New Tenant: "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   12135
      Begin VB.ComboBox cmbHouseNo 
         Height          =   435
         ItemData        =   "frmNewTenant.frx":0000
         Left            =   8400
         List            =   "frmNewTenant.frx":0040
         TabIndex        =   20
         Text            =   "Select a House No."
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtFullName 
         Height          =   435
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   9015
      End
      Begin VB.TextBox txtPhoneNo 
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtAmountDue 
         Height          =   435
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtTenantId 
         Height          =   435
         Left            =   8400
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtAmountPaid 
         Height          =   435
         Left            =   8400
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtDateAdmitted 
         Height          =   435
         Left            =   2640
         TabIndex        =   6
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "House No:"
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Full Names:"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Phone No. :"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Tenant ID:"
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Amount Due :"
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Amount Paid"
         Height          =   495
         Left            =   6000
         TabIndex        =   13
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Date Admitted:"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   2295
      End
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   480
      Top             =   240
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
End
Attribute VB_Name = "frmNewTenant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim TenantName As String


Private Sub clearData()
    txtTenantId.Text = ""
    txtFullName.Text = ""
    txtDateAdmitted.Text = ""
    txtPhoneNo.Text = ""
    txtAmountDue.Text = ""
    txtAmountPaid.Text = ""
    cmbHouseNo.Text = "Select a House No."
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If TenantName = "" Then
        MsgBox "Please select a tenant first to proceed!", vbCritical, App.Title
        Exit Sub
    ElseIf MsgBox("Are you sure you wish to delete this tenant?", vbCritical + vbYesNo, App.Title) = vbYes Then
        con.Execute "delete from tenants where fullnames = '" & TenantName & "'"
        ClearFromHouse
        MsgBox "Tenant deleted Succesfully", vbInformation, App.Title
    Exit Sub
    End If
    clearData
End Sub

Private Sub checkMatches()
    Dim house As String
    Dim foundhouse As String
    house = cmbHouseNo.Text
    
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants where houseno = '" & house & "'", con, adOpenKeyset, adLockOptimistic
    foundhouse = Rs!houseno
    
    If (Not IsEmpty(foundhouse) = True) Then
        MsgBox "Sorry that house is already occupied! Pick another please", vbCritical, App.Title
   Exit Sub
   End If
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub cmdNew_Click()
    If txtFullName.Text = "" Or txtPhoneNo.Text = "" Or txtTenantId.Text = "" Or txtDateAdmitted.Text = "" Or txtAmountPaid.Text = "" Or txtAmountDue.Text = "" Then
        MsgBox "You must enter all fields properly to proceed!", vbCritical, App.Title
    Exit Sub
    End If
    
    If cmbHouseNo.Text = "Select a House No." Then
        MsgBox "Select a house Number!", vbCritical, App.Title
    Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants", con, adOpenKeyset, adLockOptimistic
    Rs.AddNew
    Rs!tenantid = txtTenantId.Text
    Rs!fullnames = txtFullName.Text
    Rs!datein = txtDateAdmitted.Text
    Rs!mobile = txtPhoneNo.Text
    Rs!amountdue = txtAmountDue.Text
    Rs!amountpaid = txtAmountPaid.Text
    Rs!houseno = cmbHouseNo.Text
    Rs.Update
    
    UpdateHouse
    
    clearData
    MsgBox "New Tenant Saved succesfully", vbInformation, App.Title
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End Sub

Private Sub UpdateHouse()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from houses where houseno = '" & Trim(cmbHouseNo.Text) & "'", con, adOpenKeyset, adLockOptimistic
    Rs!tenant = txtFullName.Text
    Rs!dateadmitted = txtDateAdmitted.Text
    Rs.Update
End Sub

Private Sub ClearFromHouse()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from houses where tenant = '" & TenantName & "'", con, adOpenKeyset, adLockOptimistic
    Rs!tenant = ""
    Rs!dateadmitted = ""
    Rs.Update
    Rs.Close
End Sub

Private Sub cmdSave_Click()
    If txtFullName.Text = "" Or txtPhoneNo.Text = "" Or txtTenantId.Text = "" Or txtAmountDue.Text = "" Or txtAmountPaid.Text = "" Or txtAmountDue.Text = "" Then
        MsgBox "You must enter all fields properly to proceed!", vbCritical, App.Title
    Exit Sub
    End If
    
On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants where fullnames ='" & txtFullName.Text & "'", con, adOpenKeyset, adLockOptimistic
    Rs!tenantid = txtTenantId.Text
    Rs!fullnames = txtFullName.Text
    Rs!datein = txtDateAdmitted.Text
    Rs!mobile = txtPhoneNo.Text
    Rs!amountdue = txtAmountDue.Text
    Rs!amountpaid = txtAmountPaid.Text
    Rs!houseno = cmbHouseNo.Text
    Rs.Update
    
    UpdateHouse
    MsgBox "Tenant's information updated succesfully", vbInformation, App.Title
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
    
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source = " + App.Path + "\RealEstate.mdb;"
    con.Open
    Load_Tenant
End Sub

Private Sub Load_Tenant()
TenantName = frmTenants.lstTenants.Text
If TenantName = "" Then
    clearData
Else
    On Error GoTo ErrorHandler
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from tenants where fullnames = '" & TenantName & "'", con, adOpenKeyset, adLockOptimistic
    txtTenantId.Text = Rs!tenantid
    txtFullName.Text = Rs!fullnames
    txtDateAdmitted.Text = Rs!datein
    txtPhoneNo.Text = Rs!mobile
    txtAmountDue.Text = Rs!amountdue
    txtAmountPaid.Text = Rs!amountpaid
    cmbHouseNo.Text = Rs!houseno
    Exit Sub
ErrorHandler:
MsgBox Err.Description & " No. " & Err.Number
End If
End Sub


