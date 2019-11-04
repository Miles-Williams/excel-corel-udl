VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UDL_Creator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnUDL_Click()
    useList = False
    ufSerials.Show
End Sub

Private Sub btnUDLList_Click()
    useList = True
    If Range("A2").Value = vbEmpty Then
        MsgBox "Please paste or create your data beginning in cell A2."
        Exit Sub
    End If
    Set top = Range("A2")
    Set sel = Range(top, top.End(xlDown))
    udlCount = sel.Rows.Count
    ufLayout.lblUDLCount.Caption = "You have a total of " & udlCount & " labels, how would you like to lay them out?"
    ufLayout.Show
End Sub
