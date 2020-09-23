VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "DotProduct Example"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   6075
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Quit Application"
         Index           =   99
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionItem 
         Caption         =   "Normalize Vectors"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpItem 
         Caption         =   "Contents"
         Index           =   0
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelpItem 
         Caption         =   "About this Application"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_DblClick()

    Me.Height = Me.Width
    Me.Show
    
End Sub

Private Sub Form_Load()

    Me.Show
    Me.Height = Me.Width
    Me.BackColor = vbBlack
    
    Me.Show
    Call Form_Resize
    
    Load frmSplash
    frmSplash.btnClose.Visible = True
    frmSplash.Show vbModeless, Me
    
End Sub


Private Sub DrawCrossHairs(Canvas As Form)

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Canvas.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    Canvas.Line (Canvas.ScaleLeft, 0)-(Canvas.ScaleWidth, 0), RGB(0, 128, 128)
    
    ' Draw Vertical line
    Canvas.Line (0, Canvas.ScaleTop)-(0, Canvas.ScaleHeight), RGB(0, 160, 160)
    
    Dim sngX As Single
    Dim sngY As Single

    Canvas.ForeColor = RGB(0, 160, 160)
    Canvas.FontName = "Arial"
    Canvas.FontSize = 7
    For sngX = 0 To Abs(Canvas.ScaleLeft) Step (Abs(Canvas.ScaleLeft) / 8)
        For sngY = 0 To Abs(Canvas.ScaleTop) Step (Abs(Canvas.ScaleTop) / 8)
            
            ' Draw first quadrant...
            Canvas.PSet (sngX, sngY)
            
            ' ...the flip the signs to draw other quadrants.
            Canvas.PSet (sngX, -sngY)
            Canvas.PSet (-sngX, sngY)
            Canvas.PSet (-sngX, -sngY)
        
        Next sngY
    Next sngX
    
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim V1 As mdrVector
    Dim V2 As mdrVector
    Dim V1a As mdrVector
    Dim V2b As mdrVector
    Dim V3 As mdrVector
    Dim V3b As mdrVector
    Dim sngDP As Single
    Static s_LeftX As Single
    Static s_LeftY As Single
    Static s_RightX As Single
    Static s_RightY As Single
    
    If Button = vbLeftButton Then
        s_LeftX = X
        s_LeftY = Y
    ElseIf Button = vbRightButton Then
        s_RightX = X
        s_RightY = Y
    End If
    
    ' ========================
    ' ***** DRAW ROUTINE *****
    ' ========================
    Me.Cls
    Call DrawCrossHairs(Me)
    
    ' Set Vector 1
    V1.X = s_LeftX
    V1.Y = s_LeftY
    
    ' Set Vector 2
    V2.X = s_RightX
    V2.Y = s_RightY
        
    ' Normalize Vectors (optional)
    If Me.mnuOptionItem(0).Checked = True Then
        V1a = Vec3Normalize(V1)
        V2b = Vec3Normalize(V2)
    Else
        V1a = V1
        V2b = V2
    End If
    
    ' Calculate DotProduct
    sngDP = DotProduct(V1a, V2b)
    Me.CurrentX = Me.ScaleLeft
    Me.CurrentY = Me.ScaleTop
    Me.ForeColor = RGB(255, 255, 255)
    Me.FontSize = 8
    If sngDP > 0 Then
        Me.Print "The two vectors are within + or - 90 degrees of each other."
    ElseIf sngDP < 0 Then
        Me.Print "The two vectors are NOT within + or - 90 degrees of each other."
    Else
        Me.Print "Move the mouse whilst pressing either the Left or Right mouse buttons."
    End If
    
    ' Calculate new vector.
    V3.X = V1a.X * sngDP
    V3.Y = V1a.Y * sngDP
    
    ' Normalize Vector (optional)
    If Me.mnuOptionItem(0).Checked = True Then
        V3b = Vec3Normalize(V3)
    Else
        V3b = V3
    End If
    
    Call DrawVector(V1a, 2, RGB(64, 64, 255))
    Call DrawVector(V2b, 2, RGB(255, 64, 64))
    Call DrawVector(V3b, 2, RGB(64, 255, 64))
    
End Sub

Private Sub Form_Resize()

    Me.ScaleLeft = -1.25
    Me.ScaleWidth = 2.5
    Me.ScaleTop = -1.25
    Me.ScaleHeight = 2.5
    
    Call Form_MouseMove(0, 0, 0, 0)
    
End Sub

Private Sub DrawVector(V As mdrVector, DrawWidth As Integer, ForeColor As OLE_COLOR)

    Me.DrawWidth = DrawWidth
    Me.ForeColor = ForeColor
    Me.Line (0, 0)-(V.X, V.Y)
    
    Me.DrawWidth = DrawWidth * 3
    Me.PSet (V.X, V.Y)
    
    Me.Print "  " & Format(V.X, "0.0") & "," & Format(V.Y, "0.0")
    
End Sub

Private Sub mnuFileItem_Click(Index As Integer)

    Unload Me
    
End Sub


Private Sub mnuHelpItem_Click(Index As Integer)

    Select Case Index
        Case 0
            ' Launch readme file.
            Call LaunchAnyFile(Me, App.Path & "/readme.txt")
            
        Case 2  '   About this Application
            frmSplash.Show vbModeless, Me
            
    End Select
    
End Sub

Private Sub mnuOptionItem_Click(Index As Integer)
    Me.mnuOptionItem(0).Checked = Not Me.mnuOptionItem(0).Checked
End Sub


