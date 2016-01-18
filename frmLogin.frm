VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Real Estates Management Systems"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5730
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
   ScaleHeight     =   3555
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Text            =   "Enter your Password"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtUsername 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Text            =   "Enter your username"
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myusername As String
Dim mypassword As String
    
Private Sub cmdLogin_Click()
    myusername = txtUsername.Text
    mypassword = txtPassword.Text
    
    If myusername = "" Or myusername = "Enter your username" Then
        MsgBox "You must enter your username!", vbCritical
        Exit Sub
    End If
    
    If mypassword = "" Or mypassword = "Enter your password" Then
        MsgBox "You must enter your password!", vbCritical
        Exit Sub
    End If
        frmWelcome.Show
        Unload Me
    
End Sub

Private Sub txtPassword_Click()
    txtPassword.Text = ""
End Sub

Private Sub txtUsername_Click()
    txtUsername.Text = ""
End Sub
