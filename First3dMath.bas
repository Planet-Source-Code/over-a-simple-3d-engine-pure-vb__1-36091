Attribute VB_Name = "First3dMath"
'Setting up some custome Types needed for our 3d calculations

Type TRANSFORMED3d
    x1 As Double
    x2 As Double
    x3 As Double
End Type

Type VECTOR3D
    x1 As Double
    x2 As Double
    x3 As Double
    ConvertedData As TRANSFORMED3d              'Normaly you shouldn't edit these vars
End Type

Type VECTOR2D
    x1 As Double
    x2 As Double
End Type

Type LINE2D
    x1 As Double
    x2 As Double
    y1 As Double
    y2 As Double
End Type

'Public Declarations
Public VP As VECTOR3D     'The Viewpoint
Public OUTBOX As PictureBox 'Picturebox used for graphical output
Public GRIDS As Integer     'Number of squares used for relative coorinates   Confusing description ? :) I have to apologize for my bad English
Public GRIDVISIBLE As Boolean   'Is the grid (framework) visible ?
Public OOOVisible As Boolean  'Is the orign visible ?
Public SHOWFPS As Boolean     'Should FPS be shown ?
Public FPSSTARTTIME As Double   'Needed for the calculation of the FPS
Public TOTALFRAMES As Long      'Frames already shown
Public ANIMATIONRUNNING As Boolean  ' Is an there currently an animation running ?

' Initializationprocess
'
' Example:
'
' Init3D aPictureBoxHere, InHowManyPartsShouldTheBoxBePartedToCreateAGrid, IsTheGridVisible, ShouldFPSbeShown
' Init3d picture1, 4, False, True

Public Sub Init3D(OutPutPictureBox As PictureBox, NumberOfSquaresInGrid As Integer, IsGridVisible As Boolean, ShowFramesPerSecond As Boolean)

    Set OUTBOX = OutPutPictureBox
    GRIDS = NumberOfSquaresInGrid
    GRIDVISIBLE = IsGridVisible
    SHOWFPS = ShowFramesPerSecond
    
End Sub

' Sets the Viewpoint's position, realtive to the center of the picturebox used for output
'
' Example:
'
' SetViewpoint XCoordinate, YCoodinate, ZCoordinate, True
' SetViewpoint 0,0,0,True

Public Sub SetViewpoint(x1 As Double, x2 As Double, x3 As Double, Visible As Boolean)

    Dim TwipsPerQuad As Long
    
    VP.x1 = x1
    VP.x2 = -x2
    VP.x3 = x3
    
    VecTrans3D VP
    
    If Visible = True Then
        OOOVisible = True
    End If
    
End Sub

' Commands executed when a new frame is created. For example clearing the Outputbox
'
' Example:
'
' StartFrame

Public Sub StartFrame()
    
    OUTBOX.Cls
    
End Sub

' Commands executed when a frame is finished. For example inserting a grid or marking
' the location of the Viewpoint

Public Sub EndFrame()

    TOTALFRAMES = TOTALFRAMES + 1
        
    If GRIDVISIBLE = True Then
    
        TwipsPerQuad = OUTBOX.ScaleWidth / GRIDS
    
        For i = 0 To GRIDS Step 1
            OUTBOX.Line ((OUTBOX.ScaleWidth / 2) + i * TwipsPerQuad, 0)-((OUTBOX.ScaleWidth / 2) + i * TwipsPerQuad, OUTBOX.ScaleHeight), RGB(0, 255, 0)
            OUTBOX.Line ((OUTBOX.ScaleWidth / 2) + -i * TwipsPerQuad, 0)-((OUTBOX.ScaleWidth / 2) + -i * TwipsPerQuad, OUTBOX.ScaleHeight), RGB(0, 255, 0)
            OUTBOX.Line (0, (OUTBOX.ScaleHeight / 2) + i * TwipsPerQuad)-(OUTBOX.ScaleWidth, (OUTBOX.ScaleHeight / 2) + i * TwipsPerQuad), RGB(0, 255, 0)
            OUTBOX.Line (0, (OUTBOX.ScaleHeight / 2) + -i * TwipsPerQuad)-(OUTBOX.ScaleWidth, (OUTBOX.ScaleHeight / 2) + -i * TwipsPerQuad), RGB(0, 255, 0)
        Next i
        
    End If
    
    If SHOWFPS = True And ANIMATIONRUNNING = True Then
        OUTBOX.CurrentX = 1
        OUTBOX.CurrentY = 1
        OUTBOX.ForeColor = RGB(255, 255, 255)
        OUTBOX.Print "Running for " & Int((Timer - FPSSTARTTIME) * 10) / 10 & " seconds"
        OUTBOX.Print "with " & Int(TOTALFRAMES / (Timer - FPSSTARTTIME)) & " FPS"
    End If

    If OOOVisible = True Then OUTBOX.Circle (VP.ConvertedData.x1, VP.ConvertedData.x2), 3, RGB(0, 255, 0)

End Sub

' Execute this Sub before you start a new animation.
' (For example when enabeling a timer which controls a aninimation)
' It is necessary for a correct calculation of the
' Frames shown per Second
'
' Example:
'
' StartCountingFrames

Public Sub StartCountingFrames()
    
    ANIMATIONRUNNING = True
    FPSSTARTTIME = Timer
    TOTALFRAMES = 0
    
End Sub

' Execute this one after your animation is finished
' (For example when disabling a timer which controls a aninimation)

Public Sub StopCountingFrames()
    
    ANIMATIONRUNNING = False
    
End Sub

' VB is normaly working whith twips or pixles. To make the creation of new 3d Objects easier
' we use relative Coordinates. The orign is not the upper left corner of the picturebox (vb standard)
' but the center of the box.
' The use of this adapted cordinates forces us evertime before painting into the picture box to
' convert our coordinates to twips, pixles etc. again.
'
' For understanding: P(0/0/0) is the CENTER of the used picture box
' Functions used to draw lines etc. will have to convert this coordinates.
' When your picturebox's width is 800 pixel and its hight 600 then the
' Transformed coordinates will be P(400,300,0)
'
' Example:
'
' VecTrans3D AVectorToBeConverted
'
'
' dim V1 as VECTOR3D
' V1.x1 = 1
' V1.x2 = 3
' V1.x3 = 0
' VecTrans3D V1

Public Sub VecTrans3D(Vector As VECTOR3D)
    Dim TwipsPerQuad As Long
  
    TwipsPerQuad = OUTBOX.ScaleWidth / GRIDS
    
    Vector.ConvertedData.x1 = (OUTBOX.ScaleWidth / 2) + Vector.x1 * TwipsPerQuad
    Vector.ConvertedData.x2 = (OUTBOX.ScaleHeight / 2) - Vector.x2 * TwipsPerQuad            'Achse umdrehen
    Vector.ConvertedData.x3 = Vector.x3 * TwipsPerQuad
          
End Sub

' Drawing a simple Line between two points. Here is the little 3D math done.
'
' Example:
'
' DrawLine V1,V2
'
'
' dim V1 as VECTOR3D
' dim V2 as VECTOR3D
'
' V1.x1 = 1
' V1.x2 = 3
' V1.x3 = 1
'
' V2.x1 = 2
' V2.x2 = 1
' V2.x3 = 1
'
' DrawLine V1,V2
'
' ... will draw a line from V1(1,3,1) to V2(2,1,1)

Public Sub DrawLine(p1 As VECTOR3D, p2 As VECTOR3D)

    Dim L As LINE2D
    Dim Skalar As Double
    
    VecTrans3D p1
    VecTrans3D p2
    
    TwipsPerQuad = OUTBOX.ScaleWidth / GRIDS

    Skalar = -TwipsPerQuad / (p1.ConvertedData.x3 - VP.ConvertedData.x3)
    L.x1 = (VP.ConvertedData.x1 + Skalar * (p1.ConvertedData.x1 - VP.ConvertedData.x1))
    L.x2 = (VP.ConvertedData.x2 + Skalar * (p1.ConvertedData.x2 - VP.ConvertedData.x2))
       
    Skalar = -TwipsPerQuad / (p2.ConvertedData.x3 - VP.ConvertedData.x3)
    L.y1 = (VP.ConvertedData.x1 + Skalar * (p2.ConvertedData.x1 - VP.ConvertedData.x1))
    L.y2 = (VP.ConvertedData.x2 + Skalar * (p2.ConvertedData.x2 - VP.ConvertedData.x2))
    
    OUTBOX.Line (L.x1, L.x2)-(L.y1, L.y2), RGB(255, 0, 0)

End Sub
