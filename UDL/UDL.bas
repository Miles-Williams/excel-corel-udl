Attribute VB_Name = "UDL"
Option Explicit

Sub CreateUDL()

    Dim layer As CorelDRAW.layer
    Dim mainLayer As CorelDRAW.layer
    Dim textSh As CorelDRAW.Shape
    Dim rectSh As CorelDRAW.Shape
    Dim lineSh As CorelDRAW.Shape
    Dim shps As CorelDRAW.ShapeRange

    Dim totalRows As Long
    Dim i As Long
    Dim loopStart As Long
    Dim loopEnd As Long
    Dim colCount As Long
 
    Dim replaceVal As String

    colCount = 1
    
    Dim endX As Double
    Dim endY As Double

    endX = startX + (numCols * udlWidth)
    endY = startY - (numRows * udlHeight)
    
    Dim dx As Double
    Dim dy As Double
    
    dx = startX
    dy = startY
    
    If useList Then
        totalRows = sel.Rows.Count
        loopStart = 1
        loopEnd = totalRows
    Else
        loopStart = startSerial
        loopEnd = endSerial
    End If
    
    Set mainLayer = doc.ActivePage.Layers("Main")
    doc.ActivePage.CreateLayer "1"
    Set layer = doc.ActivePage.Layers("1")
    Set shps = mainLayer.Shapes.All
    
    
    For i = loopStart To loopEnd
        shps.PositionX = dx
        shps.PositionY = dy
        
        shps.CopyToLayer layer
        
        Set textSh = layer.Shapes.FindShape("Text")
        Set rectSh = layer.Shapes.FindShape("Rect")
        
        If useList Then
            replaceVal = sel.Cells(i)
        Else
            replaceVal = i
        End If
        
        textSh.Text.Replace "X", replaceVal, False
        
        If textSh.BoundingBox.Width > maxTextWidth Then
            textSh.SetSize maxTextWidth, textSh.BoundingBox.Height
        End If
        
        textSh.AlignToShape cdrAlignHCenter, rectSh
        textSh.AlignToShape cdrAlignVCenter, rectSh

        rectSh.Delete
        
        dx = dx + udlWidth
        
        If NeedNewLine(colCount) Then
            dx = startX
            dy = dy - 5
            colCount = 0
        End If
        
        colCount = colCount + 1
    Next i

    mainLayer.Visible = False
    
    layer.CreateRectangle startX, startY, endX, endY
    
    Set rectSh = layer.Shapes.First
    
    rectSh.Outline.Width = 0.01
    rectSh.Outline.Color.RGBAssign 255, 0, 0
    
    For i = (startX + udlWidth) To (endX - udlWidth) Step udlWidth
        layer.CreateLineSegment i, startY, i, endY
        Set lineSh = layer.Shapes.First
        lineSh.Outline.Width = 0.01
        lineSh.Outline.Color.HexValue = redHex
    Next i
    
    For i = startY - udlHeight To endY + udlHeight Step -udlHeight
        layer.CreateLineSegment startX, i, endX, i
        Set lineSh = layer.Shapes.First
        lineSh.Outline.Width = 0.01
        lineSh.Outline.Color.HexValue = redHex
    Next i
    
    layer.CreateLineSegment startX - 20, endY - 5, startX + 580, endY - 5
    Set lineSh = layer.Shapes.First
    lineSh.Outline.Width = 0.01
    lineSh.Outline.Color.HexValue = greenHex
    
End Sub


Sub InitCorel()

    Set corel = CreateObject("CorelDraw.Application.16")
    Set doc = corel.OpenDocument("C:\Users\wau9917\Documents\IDENTIFICATION FOLDER\LASER TEMPLATES\__MAIN TEMPLATES\UDL_TEST.cdr")
    doc.Activate
    doc.Unit = cdrMillimeter
    doc.ReferencePoint = cdrTopLeft
    
    corel.Optimization = True
    
    Call CreateUDL
    
    corel.Optimization = False
    corel.Visible = True
    corel.ActiveWindow.Refresh
    corel.Refresh
    
End Sub

Function NeedNewLine(l As Long)
    NeedNewLine = IIf(l Mod numCols = 0, True, False)
End Function


