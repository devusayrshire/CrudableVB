VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Record"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtText 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "h:mm:ss AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      _Version        =   393216
      MaxLength       =   11
      Mask            =   "##:##:## ??"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   405
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      _Version        =   393216
      Format          =   120717313
      CurrentDate     =   44968
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Text"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   315
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   555
   End
End
Attribute VB_Name = "frmAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If txtText.Text = "" Or _
        txtNumber.Text = "" Or _
        txtDate.Value = "" Or _
        txtTime.Text = "" Then
        MsgBox "Please fill-up the form properly", vbExclamation
    Else
        RS.Open "SELECT * FROM tbl_record WHERE Trim(xtext) = '" & Trim(txtText.Text) & "'", CN, adOpenKeyset, adLockReadOnly
        If RS.RecordCount = 0 Then
            RS.Close
            RS.Open "SELECT * FROM tbl_record", CN, adOpenStatic, adLockOptimistic
            With RS
                .AddNew
                .Fields("xtext") = (Me.txtText.Text)
                .Fields("xnumber") = (Me.txtNumber.Text)
                .Fields("xdate") = (Me.txtDate.Value)
                .Fields("xtime") = (Me.txtTime.Text)
                .Update
            End With
            MsgBox "Saved Successfully", vbExclamation
            Unload Me
        Else
            MsgBox "Record Already Exist", vbExclamation
            RS.Close
        End If
    End If
End Sub
