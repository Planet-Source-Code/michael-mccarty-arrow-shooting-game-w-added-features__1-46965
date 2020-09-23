VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmArrow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF00&
   Caption         =   "Arrow"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   120.684
   ScaleMode       =   0  'User
   ScaleWidth      =   154.173
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGround 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   15375
      TabIndex        =   2
      Top             =   9960
      Width           =   15375
   End
   Begin ComctlLib.ProgressBar pbarArrow 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrMove 
      Interval        =   10
      Left            =   4320
      Top             =   1920
   End
   Begin VB.Label lblWind 
      BackColor       =   &H00FFFF00&
      Caption         =   "Wind: 0 mph West"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12120
      TabIndex        =   3
      Top             =   0
      Width           =   2445
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H00FFFF00&
      Caption         =   "Speed: 0 mph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9840
      TabIndex        =   1
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "frmArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Velocity As Single
Dim x1, x2, y1, y2, oldx, oldy, xdif, ydif, oldx1, oldx2, oldy1, oldy2 As Single
Dim bowx1, bowx2, bowy1, bowy2 As Single
Dim Angle As Single
Dim Direction As Single
Dim T As Single
Dim Starter, Stopper, CurTime As Long
Dim InProgress As Boolean
Dim TargetLeft As Single
Dim TargetTop As Single
Dim Shots As Integer
Dim Hits As Integer
Dim Times As Integer
Dim Windspeed As Integer
Dim XDist As Single
Dim YDist As Single
Dim Dist As Single

'constants
Const Pi As Single = 3.14159                'Pi
Const TargetRadius As Single = 2.5          'radius of target in feet
Const Length As Single = 2.5                'length of arrow in feet
Const H As Single = 50                      'initial height in feet
Const MaxVelocity As Single = 100           'max arrow speed in mph
Const MinVelocity As Single = 5             'min arrow speed in mph
Const MaxWind As Integer = 20               'max wind speed in mph
Const HorizontalDistance As Single = 100    'horizontal distance of screen in feet

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'make sure the arrow isn't already in flight
    If T > 0 Then Exit Sub
    If KeyCode = vbKeyUp Then
        If Direction > -Pi / 3 Then
            Direction = Direction - 0.1
            UpdateDirection
        End If
    ElseIf KeyCode = vbKeyDown Then
        If Direction < Pi / 4 Then
            Direction = Direction + 0.1
            UpdateDirection
        End If
    Else
        'get current time when repeating press starts
        If Times = 1 Then
            Starter = GetTickCount
        End If
        
        InProgress = True
        Times = Times + 1
    
        'get current time difference
        If Starter > 0 Then
            CurTime = GetTickCount
            CurTime = CurTime - Starter
            CurTime = CurTime / 20
        Else
            CurTime = 0
        End If
        'make sure it doesn't go past 100, and display
        If CurTime < 100 Then
            pbarArrow.Value = CurTime
        Else
            pbarArrow.Value = 100
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'make sure the arrow isn't already in flight
    If InProgress = False Then Exit Sub
    InProgress = False
    
    'calculate time key was held down
    Stopper = GetTickCount()
    
    If Starter > 0 Then
        If Stopper - Starter > 1 Then
            Velocity = (Stopper - Starter) / 20
        Else
            Velocity = (Stopper - Starter) * 20
        End If
    Else
        Velocity = MinVelocity
    End If
        
    'don't let it go too fast or slow
    If Velocity > MaxVelocity Then Velocity = MaxVelocity
    If Velocity < MinVelocity Then Velocity = MinVelocity
    
    'show speed
    lblSpeed.Caption = "Speed: " & Velocity & " mph"
    
    'reset timer
    Starter = 0
    pbarArrow.Value = 0
    Times = 0
    
    'shoot arrow
    Play "twang"
    Arrow
End Sub

Private Sub Form_Load()

    'set values
    Randomize
    Me.ScaleWidth = HorizontalDistance
    Me.ScaleHeight = (Me.Height * ScaleWidth) / Me.Width
    SetTarget
    SetWind
    Line (0, ScaleHeight - H - ScaleHeight / 37)-(Length, ScaleHeight - H - ScaleHeight / 37)
    
    'draw lines for progress bar
    Line (pbarArrow.Left + pbarArrow.Width / 4, pbarArrow.Top + pbarArrow.Height) _
    -(pbarArrow.Left + pbarArrow.Width / 4, pbarArrow.Top + pbarArrow.Height + ScaleHeight / 50)
    Line (pbarArrow.Left + pbarArrow.Width / 2, pbarArrow.Top + pbarArrow.Height) _
    -(pbarArrow.Left + pbarArrow.Width / 2, pbarArrow.Top + pbarArrow.Height + ScaleHeight / 50)
    Line (pbarArrow.Left + (pbarArrow.Width / 4) * 3, pbarArrow.Top + pbarArrow.Height) _
    -(pbarArrow.Left + (pbarArrow.Width / 4) * 3, pbarArrow.Top + pbarArrow.Height + ScaleHeight / 50)
End Sub

Private Sub Arrow()

    'don't let the arrow keep going off the screen
    Shots = Shots + 1
    Do Until x1 > ScaleWidth + ScaleWidth / 200 Or x1 < 0 Or y2 > picGround.Top
        'time interval
        Wait 10
        T = T + 0.01
        
        'keep track of old place
        oldx = x1
        oldy = y1
        
        'find new place
        x1 = ((Velocity * Cos(Direction)) * T - Windspeed * T)
        y1 = (ScaleHeight - H) + ((Velocity * Sin(Direction)) * T + 17 * T ^ 2)
        
        'find difference to see how far it went
        xdif = x1 - oldx
        ydif = y1 - oldy
            
        'calculate the angle the arrow made as it traveled
        Angle = (Tan(ydif / xdif) ^ -1)
        
        'determine where the second points will be by proportionatly drawing
        'the arrow to the appropriate length
        x2 = ((Length * xdif) / (Sqr(xdif ^ 2 + ydif ^ 2))) + x1
        y2 = ((Length * ydif) / (Sqr(xdif ^ 2 + ydif ^ 2))) + y1
        
        'erase old arrow, and draw new one
        Line (oldx1, oldy1)-(oldx2, oldy2), Me.BackColor
        UpdateDirection
        Line (x1, y1)-(x2, y2), RGB(153, 102, 51)
        oldx1 = x1
        oldx2 = x2
        oldy1 = y1
        oldy2 = y2
        
        'collision detection
        XDist = Abs(TargetLeft - x2)
        YDist = Abs(TargetTop - y2)
        Dist = Sqr(XDist ^ 2 + YDist ^ 2)
        If Dist < TargetRadius + ScaleWidth / 200 Then
            
            'erase old target
            Me.DrawWidth = 8
            Me.FillColor = Me.BackColor
            Circle (TargetLeft, TargetTop), TargetRadius, Me.BackColor
            Me.FillColor = vbBlack
            Me.DrawWidth = 3
            Play "break"
            
            'place, and draw new target
            SetTarget
            Hits = Hits + 1
            
            'Show 10 target score, and reset
            If Hits = 10 Then
                MsgBox "You hit 10 targets in " & Shots & " shots!", , "Score"
                Shots = 0
                Hits = 0
            End If
        End If
    Loop
    
    'leave old arrow there if in the screen
    If x1 < ScaleWidth And x1 > 0 Then
        Line (x1, y1)-(x2, y2), RGB(153, 102, 51)
        Play "land"
    End If
    
    'put arrow back at start
    x1 = 0
    x2 = x1
    y1 = ScaleHeight - H
    y2 = y1
    oldx1 = x1
    oldx2 = x2
    oldy1 = y1
    oldy2 = y2
    T = 0
    
    'get new wind speed
    SetWind
End Sub

Private Sub SetTarget()
    TargetLeft = Rnd * ((ScaleWidth / 10) * 6) + ((ScaleWidth / 10) * 3)
    TargetTop = Rnd * ((ScaleHeight / 10) * 4) + ((ScaleHeight / 10) * 4)
    
    'draw target
    Circle (TargetLeft, TargetTop), TargetRadius, vbRed
    Me.DrawWidth = 8
    PSet (TargetLeft, TargetTop), vbYellow
    Me.DrawWidth = 3
End Sub

Private Sub SetWind()

    'set new wind speed
    Windspeed = Rnd * MaxWind
    lblWind.Caption = "Wind: " & Windspeed & " mph West"
End Sub

Private Sub UpdateDirection()
    
    'erase old bow
    Line (bowx1, bowy1)-(bowx2, bowy2), Me.BackColor
    
    'find new arrow
    bowx1 = 0
    bowy1 = ScaleHeight - H
    bowx2 = (Length) * (Sin((Pi / 2) - Direction))
    bowy2 = Sin(Direction) * (Length) + ScaleHeight - H
    
    'draw new arrow
    Line (bowx1, bowy1)-(bowx2, bowy2)
End Sub


'old equations
'((Velocity * 5280) / 3600) * T
'(ScaleHeight - H) + 17 * T ^ 2
