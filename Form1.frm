VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "modSelectAngle Demo"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      ClipControls    =   0   'False
      Height          =   1545
      Left            =   1890
      ScaleHeight     =   1485
      ScaleWidth      =   1665
      TabIndex        =   2
      Top             =   90
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      ClipControls    =   0   'False
      Height          =   1545
      Left            =   90
      MousePointer    =   2  'Cross
      ScaleHeight     =   1485
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   90
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1800
      Width           =   3525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DrawSelecterAngle Picture1, X, Y
Label1.Caption = theAngle
If Button = 1 Then theColor = vbRed Else theColor = vbBlack
Picture2.PSet (InterpretedMouse.X, InterpretedMouse.Y), theColor
End Sub

