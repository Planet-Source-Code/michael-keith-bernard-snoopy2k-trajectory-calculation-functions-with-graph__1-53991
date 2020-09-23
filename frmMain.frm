VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trajectory Calculator"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framGraph 
      Caption         =   "Graph"
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   5895
      Begin VB.CommandButton cmdClearPlot 
         Caption         =   "Clear Plot"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop Plotting"
         Height          =   255
         Left            =   3960
         TabIndex        =   25
         Top             =   3720
         Width           =   1815
      End
      Begin VB.PictureBox picGraph 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   5655
         TabIndex        =   24
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame framOutput 
      Caption         =   "Output"
      Height          =   2775
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   6
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   5
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   4
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Y at Time:"
         Height          =   195
         Index           =   8
         Left            =   2040
         TabIndex        =   19
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Range"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Apex"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "X at Time:"
         Height          =   195
         Index           =   7
         Left            =   2040
         TabIndex        =   16
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Air Time:"
         Height          =   195
         Index           =   6
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Y Velocity:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   750
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "X Velocity:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   240
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtVelocity 
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Velocity:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Angle:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LineCOORD
    X       As Double
    Y       As Double
    OldX    As Double
    OldY    As Double
End Type

Private bAllowPlot As Boolean

Private Sub cmdCalc_Click()

CalculateValues
bAllowPlot = True
Plot

End Sub

Private Sub CalculateValues()

Dim dblVelocity As Double, dblAngle As Double, dblTime As Double

dblVelocity = CDbl(Val(txtVelocity.Text))
dblAngle = CDbl(Val(txtAngle.Text))
dblTime = CDbl(Val(txtTime.Text))

txtOutput(0).Text = VeloX(dblVelocity, dblAngle)
txtOutput(1).Text = VeloY(dblVelocity, dblAngle)
txtOutput(2).Text = Apex(dblVelocity, dblAngle, EARTHS_GRAVITY)
txtOutput(3).Text = Range(dblVelocity, dblAngle, EARTHS_GRAVITY)
txtOutput(4).Text = AirTime(dblVelocity, dblAngle, EARTHS_GRAVITY)
txtOutput(5).Text = XPosAtTime(dblVelocity, dblAngle, dblTime)
txtOutput(6).Text = YPosAtTime(dblVelocity, dblAngle, EARTHS_GRAVITY, dblTime)

End Sub

Private Sub Plot()

Dim dblVelocity As Double, dblAngle As Double, dblTime As Double, PlotCOORD As LineCOORD

dblVelocity = CDbl(Val(txtVelocity.Text))
dblAngle = CDbl(Val(txtAngle.Text))

dblTime = 0
picGraph.AutoRedraw = True

With PlotCOORD
    .X = 0
    .Y = 0
    .OldX = 0
    .OldY = 0
End With

If bAllowPlot = False Then
    Exit Sub
End If

framGraph.Caption = "Graph - PLOTTING"

Do While PlotCOORD.OldY > -1

    DoEvents
    If bAllowPlot = False Then
        framGraph.Caption = "Graph"
        Exit Sub
    End If
    
    With PlotCOORD
        .X = (XPosAtTime(dblVelocity, dblAngle, dblTime))
        .Y = (YPosAtTime(dblVelocity, dblAngle, EARTHS_GRAVITY, dblTime))
    
        picGraph.Line (.OldX, picGraph.Height - .OldY)-(.X, picGraph.Height - .Y), vbBlack
        
        .OldX = .X
        .OldY = .Y
    End With
    
    dblTime = dblTime + 0.1
    
Loop

framGraph.Caption = "Graph"

End Sub

Private Sub ClearValues()

Dim i As Long

For i = 0 To 6
    txtOutput(i).Text = "0"
Next i
bAllowPlot = False

End Sub

Private Sub cmdClear_Click()

ClearValues

End Sub

Private Sub cmdClearPlot_Click()

picGraph.Cls

End Sub

Private Sub cmdStop_Click()

bAllowPlot = False

End Sub

Private Sub Form_Load()

ClearValues

End Sub

Private Sub Form_Terminate()

bAllowPlot = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

bAllowPlot = False

End Sub
