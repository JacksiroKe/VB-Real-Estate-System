VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Real Estates Management Systems"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "News Gothic"
      Size            =   20.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUserAccout 
      Caption         =   "User Account"
      Height          =   975
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton cmdRentPayment 
      Caption         =   "Rent Payment Records"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Width           =   4935
   End
   Begin VB.CommandButton cmdTenants 
      Caption         =   "Tenants Registration Records"
      Height          =   1095
      Left            =   6240
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
   End
   Begin VB.CommandButton cmdHouse 
      Caption         =   "House Records"
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "Logged in as "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLogOut 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblLoginAs 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Logged as User "
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      Height          =   1455
      Left            =   6000
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Shape Shape5 
      Height          =   1455
      Left            =   360
      Top             =   4800
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Click on any of the buttons to continue:"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   10095
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   360
      Top             =   1800
      Width           =   10935
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   6000
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   360
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   360
      Top             =   240
      Width           =   10935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Estates Management System"
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
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHouse_Click()
    frmHouseRecords.Show
End Sub

Private Sub cmdRentPayment_Click()
    frmRentPayment.Show
End Sub

Private Sub cmdTenants_Click()
    frmTenants.Show
End Sub

Private Sub cmdUserAccout_Click()
    frmUserAccount.Show
End Sub

Private Sub Form_Load()
    lblLoginAs.Caption = frmLogin.txtUsername.Text
End Sub

Private Sub Form_Resize()
Dim str1 As Integer, srt2 As Integer, str3 As Integer, str4 As Integer
str1 = frmWelcome.Width
str2 = str1 / 2
str3 = str2 - 600
str4 = str3 - 550

    Shape1.Width = str1 - 720
    Label1.Width = str1 - 960
    Shape4.Width = str1 - 720
    Label2.Width = str1 - 1200
    
    Label3.Left = str1 / 4 - 300
    lblLoginAs.Left = str1 / 4 + 200
    lblLogOut.Left = str2 + 200
    Shape3.Left = str2 + 225
    cmdTenants.Left = str2 + 465
    Shape6.Left = str2 + 225
    cmdUserAccout.Left = str2 + 465
    
    Shape2.Width = str3
    Shape3.Width = str3
    Shape5.Width = str3
    Shape6.Width = str3
    
    cmdHouse.Width = str4
    cmdRentPayment.Width = str4
    cmdTenants.Width = str4
    cmdUserAccout.Width = str4
    
End Sub

Private Sub lblLogOut_Click()
    frmLogin.Show
    Unload Me
End Sub
