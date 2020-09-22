VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   30
      TabIndex        =   1
      Top             =   6930
      Width           =   8265
      Begin VB.ComboBox cmbQuan 
         Height          =   315
         Left            =   3540
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   30
         Width           =   825
      End
      Begin VB.CheckBox chkDebug 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show Debug Info"
         Height          =   255
         Left            =   6630
         TabIndex        =   4
         Top             =   60
         Width           =   1575
      End
      Begin VB.ComboBox cmbRock 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pause Simulation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4680
         TabIndex        =   7
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Number of Rocks:"
         Height          =   195
         Left            =   2190
         TabIndex        =   6
         Top             =   90
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Follow:"
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   30
      Top             =   30
   End
   Begin VB.PictureBox pScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   6945
      Left            =   0
      ScaleHeight     =   6885
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrRocks As Integer    'The Current Number of Rocks Displayed
Dim FollowRock As Integer   'The Rock the user is 'following'

Private Sub cmbQuan_Click()
    
    'If Listindex=0, then display 1 rock, etc..
    Call AssignRockProperties(cmbQuan.ListIndex + 1)

End Sub

Private Sub cmbRock_Click()
    
    'Set 'Follow' Rock
    FollowRock = cmbRock.ListIndex

End Sub

Private Sub Form_Load()
    
    'Populate "Number of Rocks" Combo Box
    Dim X As Integer
    For X = 1 To 70
        cmbQuan.AddItem X
    Next
    
    'Start off with 15 rocks to display..
    'you can change this to whatever you want.
    Call AssignRockProperties(15)
    
End Sub

Private Sub AssignRockProperties(RockCount As Integer)

    Dim X As Integer
    
    CurrRocks = RockCount   'Set CurrRocks for use with other subs
    FollowRock = -1         'Reset Follow Rock Value so none are followed
    
    
    '(Re-)Populate 'Follow' Combobox
    cmbRock.Clear
    For X = 1 To RockCount
        cmbRock.AddItem "Rock #" & X
    Next
    cmbRock.AddItem "Follow None"
    
    
    'Assign Random Sizes, Speeds, Radii, slopes, etc..
    For X = 0 To RockCount - 1
        
        Rocks(X).Radius = RandomNum(100, 1000)
        Rocks(X).Speed = RandomNum(1, 50)
        Rocks(X).YSpot = RandomNum(1000, pScreen.ScaleHeight - 1000)
        Rocks(X).XSpot = RandomNum(1000, pScreen.ScaleWidth - 1000)
        Rocks(X).XSlope = RandomNum(1, 5)
        Rocks(X).YSlope = RandomNum(1, 5)
        If RandomNum(1, 2) = 1 Then Rocks(X).XSlope = (-1 * Rocks(X).XSlope)
        If RandomNum(1, 2) = 1 Then Rocks(X).YSlope = (-1 * Rocks(X).YSlope)
        Rocks(X).XStart = Rocks(X).XSpot
        Rocks(X).YStart = Rocks(X).YSpot
    
    Next
    
    'Draw the rocks
    Call DrawRocks
    
    'Start the time which will continue to move/draw rocks
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Resize()
    
    'Resize Form's Controls when resized
    
    On Error GoTo Err
    
    pScreen.Move 0, 0, ScaleWidth, ScaleHeight - Frame1.Height
    Frame1.Move 0, ScaleHeight - Frame1.Height, ScaleWidth
    chkDebug.Move ScaleWidth - (120 + chkDebug.Width)
    
Err:
    
End Sub

Private Sub Label3_Click()
    
    If Label3.Caption = "Pause Simulation" Then
        Label3.Caption = "Continue Simulation"
        Timer1.Enabled = False
    
    Else
        Label3.Caption = "Pause Simulation"
        Timer1.Enabled = True
    
    End If

End Sub

Private Sub Timer1_Timer()
    
    Dim X As Integer
    For X = 0 To CurrRocks - 1
        
        'Increment the X/Y Spots by  slope*speed
        Rocks(X).XSpot = Rocks(X).XSpot + (Rocks(X).XSlope * Rocks(X).Speed)
        Rocks(X).YSpot = Rocks(X).YSpot + (Rocks(X).YSlope * Rocks(X).Speed)
    
    Next
    
    'Draw the rocks once their posistions have been modified
    Call DrawRocks
    
End Sub

Sub DrawRocks()
    
    'Clear the Rock Field
    pScreen.Cls

    Dim X As Integer
    For X = 0 To CurrRocks - 1
        
        'Draw the Rocks(x)
        pScreen.Circle (Rocks(X).XSpot, Rocks(X).YSpot), Rocks(X).Radius, IIf(X = FollowRock, vbRed, vbWhite)
        
        
        'if Debug is checked, display info.
        If chkDebug.Value = 1 Then
            
            pScreen.ForeColor = IIf(X = FollowRock, &HFF&, &HC0FFFF)
            pScreen.FontBold = (X = FollowRock)
            pScreen.FontUnderline = (X = FollowRock)
            pScreen.Print X + 1 & ": (S:" & Rocks(X).Speed & ") (" & Rocks(X).XSpot & ", " & Rocks(X).YSpot & ")"
        
        End If
        
        
        'Check for rock going off the screen, and wrap it around, like pacman
        If (Rocks(X).XSpot > pScreen.ScaleWidth + Rocks(X).Radius) Then
            'Off Right Side of Screen
            Rocks(X).XSpot = Rocks(X).Radius * -1 'Set to left
        
        ElseIf (Rocks(X).YSpot > pScreen.ScaleHeight + Rocks(X).Radius) Then
            'Off Bottom of Screen
            Rocks(X).YSpot = Rocks(X).Radius * -1 'Set to top
        
        ElseIf (Rocks(X).XSpot < (-1 * Rocks(X).Radius)) Then
            'Off Left Side of Screen
            Rocks(X).XSpot = pScreen.ScaleWidth + Rocks(X).Radius 'Set to Right
        
        ElseIf (Rocks(X).YSpot < (-1 * Rocks(X).Radius)) Then
            'Off Top of Screen
            Rocks(X).YSpot = pScreen.ScaleHeight + Rocks(X).Radius 'Set to Bottom
        
        End If
    Next

End Sub
