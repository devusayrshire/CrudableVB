VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtText 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2280
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
      Left            =   1200
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      _Version        =   393216
      Format          =   115736577
      CurrentDate     =   44968
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Text"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   315
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadInfo()
On Error GoTo Err_380_InvalidPropVal
    CloseRS
    RS.Open "SELECT * FROM tbl_record WHERE Trim(id) = '" & Trim(frmRecList.ListView1.SelectedItem.Text) & "'", CN, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then Exit Sub
    With RS
        Me.txtText.Text = .Fields("xtext")
        Me.txtNumber = .Fields("xnumber")
        Me.txtDate.Value = .Fields("xdate")
        Me.txtTime.Text = .Fields("xtime")
        .Close
    End With

Err_380_InvalidPropVal:
    Select Case Err.Number
        Case -380
            MsgBox "The value on your date/time is invalid. Contact the Administrator", vbCritical, "Oooops!"
            Exit Sub
    End Select
End Sub

Private Sub UpdateInfo()
    CloseRS
    RS.Open "SELECT * FROM tbl_record WHERE Trim(id) = '" & Trim(frmRecList.ListView1.SelectedItem.Text) & "'", CN, adOpenStatic, adLockOptimistic
    If RS.RecordCount = 0 Then Exit Sub
    With RS
        .Fields("xtext") = (Me.txtText.Text)
        .Fields("xnumber") = (Me.txtNumber.Text)
        .Fields("xdate") = (Me.txtDate.Value)
        .Fields("xtime") = (Me.txtTime.Text)
        .Update
    End With
    MsgBox "Update Successfully", vbExclamation
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdUpdate_Click()
    Call UpdateInfo
End Sub

Private Sub Form_Load()
    Call LoadInfo
End Sub
