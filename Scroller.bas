Attribute VB_Name = "Scroller"

'Scroller Sub. Has nothing to do with our 3D engine
Public Sub Scroll(ScMessage As String, YPos As Double, XSpeed As Double, YSpeed As Double)
    Dim RedCounter As Double
    Dim RedDown As Boolean
    Dim GreenCounter As Double
    Dim GreenDown As Boolean
    Dim BlueCounter As Double
    Dim BlueDown As Boolean
    
    For i = OUTBOX.ScaleWidth To -1700 Step -1
        StartFrame
        
            OUTBOX.Font = "Courier New"
            OUTBOX.FontSize = 12
            OUTBOX.FontBold = True
            OUTBOX.FontItalic = True
          
            If RedCounter >= 255 Then RedDown = True
            If RedCounter <= 100 Then RedDown = False
            If RedDown = True Then RedCounter = RedCounter - 3 Else RedCounter = RedCounter + 3
            If GreenCounter >= 255 Then GreenDown = True
            If GreenCounter <= 100 Then GreenDown = False
            If GreenDown = True Then GreenCounter = GreenCounter - 1 Else GreenCounter = GreenCounter + 1
            If BlueCounter >= 255 Then BlueDown = True
            If BlueCounter <= 100 Then BlueDown = False
            If BlueDown = True Then BlueCounter = BlueCounter - 2 Else BlueCounter = BlueCounter + 2
            
            OUTBOX.ForeColor = RGB(Int(RedCounter), Int(GreenCounter), Int(BlueCounter))

            For j = 1 To Len(ScMessage)
                OUTBOX.CurrentY = YPos + 20 * Cos(j + (i / 40))
                OUTBOX.CurrentX = i + j * 20
                OUTBOX.Print Mid(ScMessage, j, 1);
            Next j
            
        EndFrame
        DoEvents
    Next i
    
End Sub
