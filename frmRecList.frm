VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRecList 
   Caption         =   "Rercord List"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Search for Text"
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   2775
      Begin VB.TextBox txtText 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search for ID"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3375
      Begin VB.CommandButton cmdSearchID 
         Caption         =   "Search"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtID 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1164
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "|< Top"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "< Previous"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next >"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Last >|"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   99
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Number"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmRecList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Item As ListItem
Dim PreviewPrint As Integer

Public Sub LoadRecord()
    Me.ListView1.ListItems.Clear
    CloseRS
    RS.Open "SELECT * FROM tbl_record", CN, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then Exit Sub
    While Not RS.EOF
        Set lst = Me.ListView1.ListItems.Add(, , RS("id"))
        lst.SubItems(1) = RS("xtext")
        lst.SubItems(2) = RS("xnumber")
        lst.SubItems(3) = RS("xdate")
        lst.SubItems(4) = RS("xtime")
        RS.MoveNext
    Wend
    RS.Close
    Me.ListView1.Refresh
End Sub

Private Sub EditRecord()
    frmEdit.Show 1
End Sub

Private Sub DeleteRecord()
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    
    If MsgBox("Delete Record No.:" & Me.ListView1.SelectedItem.Text & " - " & Me.ListView1.SelectedItem.ListSubItems(1).Text, vbQuestion + vbYesNo, "Record List") = vbYes Then
        CN.Execute "DELETE * FROM tbl_record WHERE Trim(id) = '" & Trim(Me.ListView1.SelectedItem.Text) & "'"
        MsgBox "Successfully Deleted", vbExclamation, "Record List"
        Call LoadRecord
    End If
End Sub

Private Sub RecordTop()
On Error Resume Next
    ListView1.SetFocus
    ListView1.ListItems(1).Selected = True
    ListView1.ListItems(1).EnsureVisible
    Toolbar1.Buttons(7).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
End Sub

Private Sub RecordPrevious()
On Error Resume Next
ListView1.SetFocus
    If ListView1.SelectedItem.Index <> 1 Then
        ListView1.ListItems(ListView1.SelectedItem.Index - 1).Selected = True
        ListView1.ListItems(ListView1.SelectedItem.Index - 1).EnsureVisible
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        Toolbar1.Buttons(10).Enabled = True
    Else
        ListView1.ListItems(1).Selected = True
        ListView1.ListItems(1).EnsureVisible
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = True
        Toolbar1.Buttons(10).Enabled = True
    End If
End Sub

Private Sub RecordNext()
On Error Resume Next
ListView1.SetFocus
    If ListView1.SelectedItem.Index <> ListView1.ListItems.Count Then
        ListView1.ListItems(ListView1.SelectedItem.Index + 1).Selected = True
        ListView1.ListItems(ListView1.SelectedItem.Index + 1).EnsureVisible
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        Toolbar1.Buttons(10).Enabled = True
    Else
        ListView1.ListItems(ListView1.ListItems.Count).Selected = True
        ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(9).Enabled = False
        Toolbar1.Buttons(10).Enabled = False
    End If
End Sub

Private Sub RecordBottom()
On Error Resume Next
    ListView1.SetFocus
    ListView1.ListItems(ListView1.ListItems.Count).Selected = True
    ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = False
    Toolbar1.Buttons(10).Enabled = False
End Sub

Private Sub SearchID()
    Me.ListView1.ListItems.Clear
    CloseRS
    RS.Open "SELECT * FROM tbl_record WHERE Trim(id) = '" & Trim(Me.txtID.Text) & "'", CN, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then Exit Sub
    While Not RS.EOF
        Set lst = Me.ListView1.ListItems.Add(, , RS("id"))
        lst.SubItems(1) = RS("xtext")
        lst.SubItems(2) = RS("xnumber")
        lst.SubItems(3) = RS("xdate")
        lst.SubItems(4) = RS("xtime")
        RS.MoveNext
    Wend
    RS.Close
    Me.ListView1.Refresh
End Sub

Private Sub FilterText()
    Me.ListView1.ListItems.Clear
    CloseRS
    RS.Open "SELECT * FROM tbl_record WHERE Trim(xtext) LIKE '%" & Trim(Me.txtText.Text) & "%'", CN, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then Exit Sub
    While Not RS.EOF
        Set lst = Me.ListView1.ListItems.Add(, , RS("id"))
        lst.SubItems(1) = RS("xtext")
        lst.SubItems(2) = RS("xnumber")
        lst.SubItems(3) = RS("xdate")
        lst.SubItems(4) = RS("xtime")
        RS.MoveNext
    Wend
    RS.Close
    Me.ListView1.Refresh
End Sub

Private Sub RecordReport()
    CloseRS
    RS.Open "SELECT * FROM tbl_record WHERE Trim(id) = '" & Trim(Me.ListView1.SelectedItem.Text) & "'", CN, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then Exit Sub
    Set rprtInfo.DataSource = RS
    
    If PreviewPrint = 1 Then
        rprtInfo.PrintReport
        PreviewPrint = 0
    ElseIf PreviewPrint = 2 Then
        rprtInfo.Show
        PreviewPrint = 0
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdSearchID_Click()
    Call SearchID
End Sub

Private Sub Form_Load()
    PreviewPrint = 0
    Call LoadRecord
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With ListView1
        .Width = Me.Width - 540
        .Height = Me.Height - 2300
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: frmAddRecord.Show 1
        Case 2: Call EditRecord
        Case 3: Call DeleteRecord
        
        Case 5: Call LoadRecord
    
        Case 7: Call RecordTop
        Case 8: Call RecordPrevious
        Case 9: Call RecordNext
        Case 10: Call RecordBottom
    
        Case 12: Call RecordReport
                PreviewPrint = 1
        Case 13: Call RecordReport
                PreviewPrint = 2
        Case 14: Unload Me
    End Select
End Sub

Private Sub txtText_Change()
    Call FilterText
End Sub
