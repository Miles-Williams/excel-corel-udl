VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSerials 
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3180
   OleObjectBlob   =   "ufSerials.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btnSerial_Click()
    
    startSerial = tbStartSerial.Value
    endSerial = tbEndSerial.Value
    
    If endSerial < startSerial Then
        MsgBox "The start value must be lower than the end value!", , "Error in values entered."
        Exit Sub
    End If
    
    tbStartSerial.Value = ""
    tbEndSerial.Value = ""
    
    ufSerials.Hide
    
    udlCount = endSerial - startSerial + 1
    ufLayout.lblUDLCount.Caption = "You have a total of " & udlCount & " labels, how would you like to lay them out?"
    
    ufLayout.Show
End Sub
