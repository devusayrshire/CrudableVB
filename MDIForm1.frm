VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Crudable VB"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   10110
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Begin VB.Menu mRecordList 
         Caption         =   "Record List"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mRecordList_Click()
    frmRecList.Show
End Sub
