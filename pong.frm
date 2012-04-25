VERSION 5.00
Begin VB.Form Pong 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pong - HywC Productions 01"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   720
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   6600
      Picture         =   "pong.frx":0000
      ScaleHeight     =   1380
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   360
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   360
      Picture         =   "pong.frx":1A22
      ScaleHeight     =   1380
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System X3"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "System X3"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00808080&
      Height          =   255
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "K = Up      M = Down"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "A = Up      Z = Down"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pong - HywC  8.12.01"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7200
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7200
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Pong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const maxheight = 3480
Private Const minheight = 720
Private Const movement = 300
Private Const ballspeed = 1
Private Const ballmaxside = 300
Private Const ballmaxup = 300
Private Const speed = 120
Private a As Variant
Private m0ve As Integer
Private ballupmovement As Integer
Private ballsidemovement As Integer
Private leftscore As Integer
Private rightscore As Integer
Private rememberside As Integer
Private rememberup As Integer
Private continue As Boolean
Private wholeball As Integer
Dim PauseTime, Start
Private Sub Form_Initialize()
rand_max = 1
Randomize
a = 0.5 + (0.5 * Rnd())
ballsidemovement = Int(speed * a)
    If ballsidemovement = 0 Then ballsidemovement = 1
ballupmovement = Int((((speed ^ 2) * (1 - (a ^ 2))) ^ 0.5))
    If ballupmovement = 0 Then ballupmovement = 1
End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 65 Or KeyCode = 97) Then
    Picture1.Top = Picture1.Top - movement
    If Picture1.Top < minheight Then
    Picture1.Top = minheight
    End If
    If Picture1.Top > maxheight Then
    Picture1.Top = maxheight
    End If
End If
If (KeyCode = 90 Or KeyCode = 122) Then
    Picture1.Top = Picture1.Top + movement
    If Picture1.Top < minheight Then
    Picture1.Top = minheight
    End If
    If Picture1.Top > maxheight Then
    Picture1.Top = maxheight
    End If
End If
If (KeyCode = 75 Or KeyCode = 107) Then
    Picture2.Top = Picture2.Top - movement
    If Picture2.Top < minheight Then
    Picture2.Top = minheight
    End If
    If Picture2.Top > maxheight Then
    Picture2.Top = maxheight
    End If
End If
If (KeyCode = 77 Or KeyCode = 109) Then
    Picture2.Top = Picture2.Top + movement
    If Picture2.Top < minheight Then
    Picture2.Top = minheight
    End If
    If Picture2.Top > maxheight Then
    Picture2.Top = maxheight
    End If
End If
End Sub
Private Sub Timer1_Timer()
Timer1.Interval = 1

If (ballsidemovement > 0) Then
If ((Shape1.Top + Shape1.Height > Picture1.Top) And (Shape1.Top < (Picture1.Top + (Picture1.Height / 3))) And (Shape1.Left > (Picture1.Left + (Picture1.Width * 0.5))) And (Shape1.Left < (Picture1.Left + Picture1.Width + 5))) Then
ballsidemovement = 0 - ballsidemovement
ballupmovement = ballupmovement - 20
ElseIf ((Shape1.Top > (Picture1.Top + (Picture1.Height * (2 / 3)))) And (Shape1.Top < (Picture1.Top + (Picture1.Height + 5))) And (Shape1.Left > (Picture1.Left + (Picture1.Width * 0.5))) And (Shape1.Left < (Picture1.Left + Picture1.Width + 5))) Then
ballsidemovement = 0 - ballsidemovement
ballupmovement = ballupmovement + 20
ElseIf Picture1.Top + Picture1.Height + 5 > Shape1.Top And Picture1.Top - Shape1.Height <= Shape1.Top And (Shape1.Left > (Picture1.Left + (Picture1.Width * 0.5))) And (Shape1.Left < (Picture1.Left + Picture1.Width + 15)) Then
ballsidemovement = 0 - ballsidemovement
ballsidemovement = ballsidemovement * 1.25
ballupmovement = ballupmovement * 1.25
End If
End If

If (ballsidemovement < 0) Then
If ((Shape1.Top + Shape1.Height > Picture2.Top) And (Shape1.Top < (Picture2.Top + (Picture2.Height / 3)))) And ((Shape1.Left + Shape1.Width) < (Picture2.Left + (Picture2.Width * 0.5))) And (Shape1.Left + Shape1.Width > (Picture2.Left - 5)) Then
ballsidemovement = 0 - ballsidemovement
ballupmovement = ballupmovement + 20
ElseIf ((Shape1.Top >= (Picture2.Top + (Picture2.Height * (2 / 3)))) And (Shape1.Top < (Picture2.Top + (Picture2.Height + 5))) And ((Shape1.Left + Shape1.Width) < (Picture2.Left + (Picture2.Width * 0.5))) And (Shape1.Left + Shape1.Width > (Picture2.Left - 5))) Then
ballsidemovement = 0 - ballsidemovement
ballupmovement = ballupmovement - 20
ElseIf ((Shape1.Top + Shape1.Height >= (Picture2.Top))) And (Shape1.Top < (Picture2.Top + Picture2.Height + 5)) And ((Shape1.Left + Shape1.Width) < (Picture2.Left + (Picture2.Width * 0.5))) And (Shape1.Left + Shape1.Width > (Picture2.Left - 5)) Then
ballsidemovement = 0 - ballsidemovement
ballsidemovement = ballsidemovement * 1.25
ballupmovement = ballupmovement * 1.25
End If
End If

If ballupmovement > 100 Then
ballupmovement = 100
End If
If ballupmovement < -100 Then
ballupmovement = -100
End If
If ballsidemovement < -300 Then
ballsidemovement = -300
End If
If ballsidemovement > 300 Then
ballsidemovement = 300
End If


If Shape1.Left + Shape1.Width > Picture2.Left + Picture2.Width Then
leftscore = leftscore + 1
Label2.Caption = leftscore
ballsidemovement = 0
ballupmovement = 0
Shape1.Top = 2520
Shape1.Left = 3480
Picture1.Top = 2040
Picture2.Top = 2040
rand_max = 1
Randomize
a = 0.5 + (0.5 * Rnd())
ballsidemovement = Int(speed * a)
ballupmovement = Int((((speed ^ 2) * (1 - (a ^ 2))) ^ 0.5))
End If

If Shape1.Left < Picture1.Left Then
rightscore = rightscore + 1
Label6.Caption = rightscore
ballsidemovement = 0
ballupmovement = 0
Shape1.Top = 2520
Shape1.Left = 3480
Picture1.Top = 2040
Picture2.Top = 2040
rand_max = 1
Randomize
a = 0.5 + (0.5 * Rnd())
ballsidemovement = Int(speed * a)
ballupmovement = Int((((speed ^ 2) * (1 - (a ^ 2))) ^ 0.5))
End If

If (ballupmovement < 0) Then
ElseIf ballupmovement > 0 Then
    If (Shape1.Top > 600) Then
    ElseIf (Shape1.Top < 600) Then
    ballupmovement = 0 - ballupmovement
    End If
End If


If (ballupmovement > 0) Then
ElseIf (ballupmovement < 0) Then
    If ((Shape1.Height + Shape1.Top) < 4920) Then
    ElseIf ((Shape1.Height + Shape1.Top) > 4920) Then
    ballupmovement = 0 - ballupmovement
    End If
End If

Shape1.Top = Shape1.Top - ballupmovement
Shape1.Left = Shape1.Left - ballsidemovement

Picture1.Refresh
Picture2.Refresh
End Sub

