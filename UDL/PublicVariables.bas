Attribute VB_Name = "PublicVariables"
Option Explicit

Public Const udlWidth As Double = 20
Public Const udlHeight As Double = 5
Public Const maxWidth As Double = 600
Public Const maxHeight As Double = 50
Public Const maxTextWidth As Double = 18

Public Const startX As Double = 20
Public Const startY As Double = 450

Public Const redHex As String = "FF0000"
Public Const greenHex As String = "00FF00"

Public Const corelApp As String = "CorelDraw.Application.16"
Public Const corelDoc As String = "C:\Users\wau9917\Documents\IDENTIFICATION FOLDER\LASER TEMPLATES\__MAIN TEMPLATES\UDL_TEST.cdr"

Public sel As Range
Public top As Range

Public useList As Boolean

Public numCols As Long
Public numRows As Long

Public startSerial As Long
Public endSerial As Long

Public udlCount As Long

Public corel As CorelDRAW.Application
Public doc As CorelDRAW.Document


