Attribute VB_Name = "Cnn"

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset

Public Sub Main()
On Error GoTo Err:
    CN.Provider = "MICROSOFT.ACE.OLEDB.12.0;"
    CN.Open App.Path & "\Dbase\dbmain.mdb"
    frmLogin.Show
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical
    End
End Sub

Public Sub CloseRS()
    If RS.State = adStateOpen Then RS.Close
End Sub

Public Sub ResetErrIn(frm As Form)
    CN.Close
    CN.Open App.Path & "\Dbase\dbmain.mdb"
    MsgBox Err.Description, vbExclamation
    Err.Clear
End Sub
