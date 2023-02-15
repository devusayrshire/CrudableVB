VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim invalidpw As Integer

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtUserName = "" Then
        MsgBox "Please Enter your Username, try again!", , "Login"
        txtUserName.SetFocus
    ElseIf txtPassword = "" Then
        MsgBox "Please Enter your Password, try again!", , "Login"
        txtPassword.SetFocus
    Else
        RS.Open "SELECT * FROM tbl_user WHERE Trim(username) = '" & Trim(txtUserName.Text) & _
            "' AND Trim(password) = '" & Trim(txtPassword.Text) & "'", CN, adOpenKeyset, adLockReadOnly
        
        If RS.RecordCount = 0 Then
            invalidpw = invalidpw + 1
            If invalidpw = 1 Then
                MsgBox "Invalid Username and Password, try again!", , "Login 1st Attempt"
                RS.Close
                txtUserName.SetFocus
            ElseIf invalidpw = 2 Then
                MsgBox "Invalid Username and Password, try again!", , "Login 2nd Attempt"
                RS.Close
                txtUserName.SetFocus
            ElseIf invalidpw = 3 Then
                MsgBox "Invalid Username and Password, try again!", , "Login Last Attempt"
                RS.Close
                Unload Me
            Else
                MsgBox "Invalid Username and Password, try again!", , "Login"
                RS.Close
                txtUserName.SetFocus
            End If
        Else
            MsgBox "Access Granted.", vbInformation + vbOKOnly, "System message"
            Screen.MousePointer = vbNormal
            LoginSucceeded = True
            RS.Close
            Unload Me
            MDIForm1.Show
        End If
    End If
End Sub

Private Sub Form_Load()
    invalidpw = 0
End Sub
