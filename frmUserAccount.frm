VERSION 5.00
Begin VB.Form frmUserAccount 
   Caption         =   "My User Account"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogout 
      Caption         =   "LogOut?"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   5655
   End
   Begin VB.Label lblUsername 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "You are Logged in as:"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      Height          =   4335
      Left            =   360
      Top             =   1560
      Width           =   8055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My User Account"
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
      Width           =   7575
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   360
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmUserAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogout_Click()
    Unload Me
    Unload frmWelcome
    frmLogin.Show
    
End Sub

Private Sub Form_Load()
    lblUsername.Caption = frmWelcome.lblLoginAs.Caption
End Sub
