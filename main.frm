VERSION 5.00
Begin VB.Form main 
   Caption         =   "First3DApp"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdScroller 
      Caption         =   "Scroller"
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   10920
      Top             =   7800
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   8400
      TabIndex        =   6
      Top             =   8520
      Width           =   2655
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Move your mouse over the picturebox to adjust the viewpoint and use mousebuttons to zoom"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Timer MouseTimer 
      Interval        =   5000
      Left            =   11880
      Top             =   7800
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Show &Grid"
      Height          =   735
      Left            =   11160
      TabIndex        =   5
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "&Clear Screen"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   11400
      Top             =   7800
   End
   Begin VB.CommandButton cmdAni2 
      Caption         =   "Animation &2"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAni1 
      Caption         =   "Animation &1"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox OutPutBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8175
      Left            =   120
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   541
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   813
      TabIndex        =   1
      Top             =   120
      Width           =   12255
   End
   Begin VB.CommandButton cmdBox 
      Caption         =   "&Draw a Box"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   8640
      Width           =   1215
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'#################################################
'
' That's were everything begins ;)
'
' VERY BASIC 3D ENGINE, PURE VB only some math ;)
'
' Author: over (overkillpage@gmx.net)
'
' Comments and suggestions are of course welcome
'
'#################################################


'Preparing the Form, setting up a picturebox for 3d output, setting the viewpoint ...
'Please keep the order of sub calls
'For further information on the SUB functions have a look at the First3dMath.bas
Private Sub Form_Load()

    Init3D OutPutBox, 4, False, True
    
    StartFrame
    
    SetViewpoint 0, 0, 0, False
    
    EndFrame

End Sub

' This Sub is a simple example of drawing a 3D Object using First3dMath.bas
'
' The Initializationprocess was already done in Form_Load

Private Sub cmdBox_Click()


' #### Clean up ####
'
' The cleaning up process is only necessary if there were already objects drawn
' If you have just done the Initializationprocess you can leave this part out
        
        SetViewpoint 0, 0, 0, False
        Timer1.Enabled = False
        Timer2.Enabled = False
        
' #### Clean up ####



' Here we go ! ...

' Declearing two Vectors which will be used to store the coordinates of two points
' We will draw lines between these points.

Dim p1 As VECTOR3D
Dim p2 As VECTOR3D

' StartFrame is always used before drawing lines etc into a new frame. It clears the
' outputpicturebox etc.

StartFrame

' ######################################################
' ########################  BOX  #######################
' ########################  TOP  #######################
' ######################################################

' Now we are gonna draw some line to form a 3D box. The procedure is always the same:
'
' 1. Storing the Coordinates of the two points in p1 and p2
' Example:
'
' p1.x1 = -1
' p1.x2 = 1
' p1.x3 = -1
'
' p2.x1 = 1
' p2.x2 = 1
' p2.x3 = -1
'
' 2. Drawing the line
' Example:
'
' DrawLine p1,p2
'
' ... do the former steps for every line ...

p1.x1 = -1
p1.x2 = 1
p1.x3 = -1

p2.x1 = 1
p2.x2 = 1
p2.x3 = -1

DrawLine p1, p2

p1.x1 = 1
p1.x2 = 1
p1.x3 = -1

p2.x1 = 1
p2.x2 = 1
p2.x3 = -3

DrawLine p1, p2

p1.x1 = 1
p1.x2 = 1
p1.x3 = -3

p2.x1 = -1
p2.x2 = 1
p2.x3 = -3

DrawLine p1, p2

p1.x1 = -1
p1.x2 = 1
p1.x3 = -3

p2.x1 = -1
p2.x2 = 1
p2.x3 = -1

DrawLine p1, p2

' ######################## bottom #######################

p1.x1 = -1
p1.x2 = -1
p1.x3 = -1

p2.x1 = 1
p2.x2 = -1
p2.x3 = -1

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -1

p2.x1 = 1
p2.x2 = -1
p2.x3 = -3

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -3

p2.x1 = -1
p2.x2 = -1
p2.x3 = -3

DrawLine p1, p2

p1.x1 = -1
p1.x2 = -1
p1.x3 = -3

p2.x1 = -1
p2.x2 = -1
p2.x3 = -1

DrawLine p1, p2

' ######################## sides #######################

p1.x1 = -1
p1.x2 = -1
p1.x3 = -1

p2.x1 = -1
p2.x2 = 1
p2.x3 = -1

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -1

p2.x1 = 1
p2.x2 = 1
p2.x3 = -1

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -3

p2.x1 = 1
p2.x2 = 1
p2.x3 = -3

DrawLine p1, p2

p1.x1 = -1
p1.x2 = -1
p1.x3 = -3

p2.x1 = -1
p2.x2 = 1
p2.x3 = -3

DrawLine p1, p2

'EndFrame is always used when a new frame is complete
'It prints out the Frames per Second, draws our Viewpoint if requested ...
EndFrame

End Sub

'Similar to cmdBox_Click, but in a timer with some changing vars to create a animation

Private Sub Timer1_Timer()

Dim p1 As VECTOR3D
Dim p2 As VECTOR3D

StartFrame

AnimationCounter1 = AnimationCounter1 - 0.1
AniMovement = AniMovement + (Sin(AnimationCounter1)) / 10

' #######################################################
' ######################## QUADER #######################
' ######################## DECKEL #######################
' #######################################################

p1.x1 = -1 + AniMovement
p1.x2 = 1
p1.x3 = -1 + AniMovement

p2.x1 = 1 - AniMovement
p2.x2 = 1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

p1.x1 = 1 - AniMovement
p1.x2 = 1
p1.x3 = -1 + AniMovement

p2.x1 = 1 - AniMovement
p2.x2 = 1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

p1.x1 = 1 - AniMovement
p1.x2 = 1
p1.x3 = -3 + AniMovement

p2.x1 = -1 + AniMovement
p2.x2 = 1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

p1.x1 = -1 + AniMovement
p1.x2 = 1
p1.x3 = -3 + AniMovement

p2.x1 = -1 + AniMovement
p2.x2 = 1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

' ######################## Boden #######################

p1.x1 = -1
p1.x2 = -1
p1.x3 = -1 + AniMovement

p2.x1 = 1
p2.x2 = -1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -1 + AniMovement

p2.x1 = 1
p2.x2 = -1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -3 + AniMovement

p2.x1 = -1
p2.x2 = -1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

p1.x1 = -1
p1.x2 = -1
p1.x3 = -3 + AniMovement

p2.x1 = -1
p2.x2 = -1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

' ######################## Seiten #######################

p1.x1 = -1
p1.x2 = -1
p1.x3 = -1 + AniMovement

p2.x1 = -1 + AniMovement
p2.x2 = 1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -1 + AniMovement

p2.x1 = 1 - AniMovement
p2.x2 = 1
p2.x3 = -1 + AniMovement

DrawLine p1, p2

p1.x1 = 1
p1.x2 = -1
p1.x3 = -3 + AniMovement

p2.x1 = 1 - AniMovement
p2.x2 = 1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

p1.x1 = -1
p1.x2 = -1
p1.x3 = -3 + AniMovement

p2.x1 = -1 + AniMovement
p2.x2 = 1
p2.x3 = -3 + AniMovement

DrawLine p1, p2

EndFrame

End Sub

' Necessary for the Zoom function
Private Sub MouseTimer_Timer()
    VP.x1 = 0
    VP.x2 = 0
    VecTrans3D VP
    MouseTimer.Enabled = False
End Sub

' Clears out the PictureBox by drawing an empty frame
Private Sub cmdCls_Click()

StartFrame
EndFrame

End Sub

'Enable/Disable the timer used for Animation1
Private Sub cmdAni1_Click()
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        StopCountingFrames
    Else
        Timer2.Enabled = False
        StartCountingFrames
        SetViewpoint 0, 0, 0, False
        Timer1.Enabled = True
    End If
End Sub

'Enable/Disable the timer used for Animation2. Not implented yet.
Private Sub cmdAni2_Click()
    If Timer2.Enabled = True Then
        Timer2.Enabled = False
        StopCountingFrames
    Else
        Timer1.Enabled = False
        StartCountingFrames
        Timer2.Enabled = True
    End If
End Sub

'Draws/Hides the used Grid
Private Sub cmdGrid_Click()
    
    StartFrame
    If GRIDVISIBLE = True Then GRIDVISIBLE = False Else GRIDVISIBLE = True
    EndFrame
    
End Sub

' Necessary to make the viewpoint follow the mouspointer
Private Sub OutPutBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MouseTimer.Enabled = False
    MouseTimer.Enabled = True
    
    VP.ConvertedData.x1 = X
    VP.ConvertedData.x2 = Y
        
End Sub

' Necessary for the zooming function
Private Sub OutPutBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseDown = True
        
    TwipsPerQuad = OUTBOX.ScaleWidth / GRIDS

    Do
        If Button = 1 Then VP.ConvertedData.x3 = VP.ConvertedData.x3 + 0.01
        If Button = 2 And VP.ConvertedData.x3 > -TwipsPerQuad + 20 Then VP.ConvertedData.x3 = VP.ConvertedData.x3 - 0.01
        DoEvents
    Loop Until MouseDown = False
    
End Sub

' Necessary for the zooming function
Private Sub OutPutBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub

' Displays a scrooler. Has nothing to do with our 3D animation.
Private Sub cmdScroller_Click()

    'Clean up
    Timer1.Enabled = False
    Timer2.Enabled = False
    StopCountingFrames
    'Clean up
    
    StartCountingFrames
    Scroll "THANK YOUR FOR TRYING MY LITTLE 3D ENGINE, Contact me: overkillpage@gmx.net", OUTBOX.ScaleHeight / 2, 1, 0.0001
    StopCountingFrames
    
End Sub



