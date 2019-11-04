VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufLayout 
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3900
   OleObjectBlob   =   "ufLayout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnLayout_Click()
    
    Dim continue As Variant
    
    If IsNumeric(tbNumCols.Value) = False Or IsNumeric(tbNumRows.Value) = False Then
        MsgBox "Please enter a valid number."
        Exit Sub
    Else
        numCols = tbNumCols.Value
        numRows = tbNumRows.Value
    End If
    
    If numCols < 1 Or numRows < 1 Then
        MsgBox "Please enter a valid number."
        Exit Sub
    End If
    
    If (numCols * udlWidth) + startX > maxWidth Then
        MsgBox "You can not fit " & numCols & " labels horizontally on a sheet." & _
                vbCrLf & vbCrLf & _
                "The maximum you can fit is " & ((maxWidth - startX - udlWidth) / udlWidth) & "." & _
                vbCrLf & vbCrLf & _
                "Please enter a valid number."
                Exit Sub
    End If
    
    If (numRows - 1) * numCols >= udlCount Then
        continue = MsgBox("You are asking for more rows than you require." & _
                            vbCrLf & vbCrLf & _
                            "This will result in a row of blank labels being created." & _
                            vbCrLf & vbCrLf & _
                            "Do you still wish to continue?", vbYesNo)
        If continue <> vbYes Then Exit Sub
    End If
    
    If numRows * numCols < udlCount Then
        MsgBox "You can not fit " & udlCount & " labels into a " & numCols & _
                " by " & numRows & " grid." & _
                vbCrLf & vbCrLf & _
                "Please enter approriate values."
        Exit Sub
    End If
    
    tbNumCols.Value = ""
    tbNumRows.Value = ""
    
    ufLayout.Hide
    
    Call InitCorel
End Sub
