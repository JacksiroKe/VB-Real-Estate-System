VERSION 5.00
Begin VB.Form frmUserAcc 
   Caption         =   "My User Account"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12120
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
   ScaleHeight     =   6390
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton txtClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton txtConfirm 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   2760
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create A New User Account"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Your Password:"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter your Username:"
      BeginProperty Font 
         Name            =   "News Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   600
      Top             =   1560
      Width           =   10935
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   600
      Top             =   240
      Width           =   10935
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   10815
   End
End
Attribute VB_Name = "frmUserAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
