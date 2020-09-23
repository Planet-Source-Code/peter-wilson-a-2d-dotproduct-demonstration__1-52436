VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   405
      Left            =   4935
      TabIndex        =   0
      ToolTipText     =   "Note: The chessy animation has nothing to do with dot-products, it's just something I did for fun."
      Top             =   2175
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://dev.midar.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   60
      MouseIcon       =   "frmSplash.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Visit my web site today."
      Top             =   2295
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 Peter Wilson"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Note: The chessy animation has nothing to do with dot-products, it's just something I did for fun."
      Top             =   390
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A 2D Dot Product Demonstration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Note: The chessy animation has nothing to do with dot-products, it's just something I did for fun."
      Top             =   60
      Width           =   4950
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mdrAnt
    PosX As Single
    PosY As Single
    dx As Single
    dy As Single
    Speed As Single
    DeadCount As Single
End Type

Private myAnts() As mdrAnt

Private m_ZoomActual As Single
Private m_ZoomRequired As Single

Private m_intClosingIn As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub AdjustZoom()

    Dim sngDelta As Single
    
    sngDelta = (m_ZoomActual - m_ZoomRequired)
    
    m_ZoomActual = m_ZoomActual - (sngDelta / 16)
    Call Form_Resize


End Sub

Private Sub Check4Collision()

    Dim intN As Integer
    Dim intK As Integer
    
    Dim deltaX As Single
    Dim deltaY As Single
    
    ' Check Ants for Collisions (a very slow process)
    For intN = LBound(myAnts) To UBound(myAnts)
        For intK = (intN + 1) To UBound(myAnts)
        
            deltaX = myAnts(intN).PosX - myAnts(intK).PosX
            deltaY = myAnts(intN).PosY - myAnts(intK).PosY
            
            If Abs(deltaX) < 0.01 Then
                If Abs(deltaY) < 0.01 Then
                    ' Ants have collided, thus reset them both!
                    myAnts(intN).dx = myAnts(intN).dx + (Rnd - 0.5) / 50
                    myAnts(intN).dy = myAnts(intN).dy + (Rnd - 0.5) / 50

                    myAnts(intK).dx = myAnts(intK).dx + (Rnd - 0.5) / 50
                    myAnts(intK).dy = myAnts(intK).dy + (Rnd - 0.5) / 50
                    
                    myAnts(intN).DeadCount = 0
                    myAnts(intK).DeadCount = 0
                    
                End If
            End If
            
        Next intK
    Next intN
    
End Sub

Private Sub DrawAnts()

    ' Draw Ants.
    ' ==========
    Dim intN As Integer
    Dim intActive As Integer
    Dim sngAvg As Single
    
    Me.Cls
    Me.DrawWidth = 6
    
    sngAvg = 0
    intActive = 0
    For intN = LBound(myAnts) To UBound(myAnts)
        
        ' Calculate speed (optional I suppose)
        myAnts(intN).Speed = Sqr(myAnts(intN).dx ^ 2 + myAnts(intN).dy ^ 2)
        
        sngAvg = sngAvg + myAnts(intN).Speed
        
        ' 0.01          Fastest!
        ' 0.005             1/2 speed of maximum (ie. 0.01)
        ' 0.0025            1/4
        ' 0.00125           1/8
        ' 0.000625          1/16
        ' 0.0003125         1/32
        ' 0.00015625        1/64
        ' 0.000078125       1/128
        ' 0.0000390625      1/256
        ' 0.00001953125     1/512
        ' 0.000009765625    1/1024
        
        If myAnts(intN).DeadCount > 0 Then
            If myAnts(intN).DeadCount > 150 Then
                Me.ForeColor = RGB(0, 0, 127)
            ElseIf myAnts(intN).DeadCount > 100 Then
                Me.ForeColor = RGB(0, 0, 255)
            ElseIf myAnts(intN).DeadCount > 50 Then
                Me.ForeColor = RGB(0, 127, 255)
            End If
        Else
            If myAnts(intN).Speed > 0.005 Then
                Me.ForeColor = RGB(0, 255, 0)
                intActive = intActive + 1
            ElseIf myAnts(intN).Speed > 0.0025 Then
                Me.ForeColor = RGB(255, 255, 0)
            ElseIf myAnts(intN).Speed > 0.00125 Then
                Me.ForeColor = RGB(255, 127, 0)
            ElseIf myAnts(intN).Speed > 0.000625 Then
                Me.ForeColor = RGB(255, 0, 0)
            ElseIf myAnts(intN).Speed > 0.0003125 Then
                Me.ForeColor = RGB(127, 0, 0)
            End If
            
        End If
        Me.PSet (myAnts(intN).PosX, myAnts(intN).PosY)
    
    Next intN
    
    sngAvg = sngAvg / intN
    
End Sub

Private Sub Reset()
    
    ReDim myAnts(190)
    Dim intN As Integer
    
    For intN = LBound(myAnts) To UBound(myAnts)
        myAnts(intN).PosX = Rnd - 0.5
        myAnts(intN).PosY = Rnd - 0.5
        
        myAnts(intN).dx = (Rnd - 0.5) / 50
        myAnts(intN).dy = (Rnd - 0.5) / 50
        myAnts(intN).Speed = Sqr(myAnts(intN).dx ^ 2 + myAnts(intN).dy ^ 2)
    Next intN

    m_ZoomRequired = 1
    m_ZoomActual = 25
    
End Sub

Private Sub Form_Click()
    Call Reset
End Sub

Private Sub Form_Load()
        
    Me.AutoRedraw = True
    m_ZoomActual = 0
    
    Call Reset
    
End Sub

Private Sub Form_Resize()

    If m_ZoomActual <> 0 Then
        Me.ScaleLeft = -(m_ZoomActual / 2)
        Me.ScaleWidth = m_ZoomActual
        Me.ScaleTop = -(m_ZoomActual / 2)
        Me.ScaleHeight = m_ZoomActual
    End If

End Sub

Private Sub Label1_Click(Index As Integer)
    
    If Index = 2 Then Call LaunchAnyFile(Me, "http://dev.midar.com/default.asp")
    
End Sub

Private Sub Timer1_Timer()

    Dim intN As Integer
    Dim intK As Integer
    
    Call MoveAnts
    
    Call DrawAnts

    Call Check4Collision
    
    Call AdjustZoom
    
End Sub


Private Sub MoveAnts()

    Dim intN As Integer
    
    ' Move Ants.
    ' ==========
    For intN = LBound(myAnts) To UBound(myAnts)
        myAnts(intN).PosX = myAnts(intN).PosX + myAnts(intN).dx
        myAnts(intN).PosY = myAnts(intN).PosY + myAnts(intN).dy
        
        ' Enforce Windows Bounds by making ants bounce off walls.
        If Abs(myAnts(intN).PosX) > 0.5 Then myAnts(intN).dx = -myAnts(intN).dx
        If Abs(myAnts(intN).PosY) > 0.5 Then myAnts(intN).dy = -myAnts(intN).dy
        
        ' Slow down ants.
        If myAnts(intN).dx <> 0 Then myAnts(intN).dx = myAnts(intN).dx * 0.95
        If Abs(myAnts(intN).dx) < 0.00001953125 Then myAnts(intN).dx = 0
        If myAnts(intN).dy <> 0 Then myAnts(intN).dy = myAnts(intN).dy * 0.95
        If Abs(myAnts(intN).dy) < 0.00001953125 Then myAnts(intN).dy = 0
        
        
        If (myAnts(intN).dx = 0) And (myAnts(intN).dy = 0) Then
            ' Ant is inactive.
            myAnts(intN).DeadCount = myAnts(intN).DeadCount + 1
        End If
                
    Next intN
    
End Sub





