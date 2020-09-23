VERSION 5.00
Begin VB.Form frmExampleTwo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rohit's Four Colour Gradient Test 2"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmExampleTwo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   6600
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   4800
      Width           =   255
   End
   Begin VB.PictureBox picFour 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   3480
      ScaleHeight     =   4035
      ScaleWidth      =   3015
      TabIndex        =   6
      Top             =   4800
      Width           =   3075
   End
   Begin VB.PictureBox picThree 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   3480
      ScaleHeight     =   4035
      ScaleWidth      =   3015
      TabIndex        =   4
      Top             =   360
      Width           =   3075
   End
   Begin VB.PictureBox picTwo 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Left            =   60
      ScaleHeight     =   4035
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   360
      Width           =   3075
   End
   Begin VB.PictureBox picOne 
      AutoRedraw      =   -1  'True
      Height          =   3375
      Left            =   60
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label lblSat 
      Caption         =   "Hue / Saturation (VB, Paintbrush Style)"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   4500
      Width           =   3075
   End
   Begin VB.Label lblBlack 
      Caption         =   "Hue / Blackness"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label lblHue 
      Caption         =   "Hue / Whiteness"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label lblStyle1 
      Caption         =   "Another Style"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   4560
      Width           =   1695
   End
End
Attribute VB_Name = "frmExampleTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Form_Load()
Dim iTick As Long
    
    iTick = GetTickCount
    '' Hue Whiteness Model
    DrawGradient picTwo.hdc, 10, 10, 30, 250, vbRed, vbYellow, vbWhite, vbWhite, True
    DrawGradient picTwo.hdc, 40, 10, 30, 250, vbYellow, vbGreen, vbWhite, vbWhite, True
    DrawGradient picTwo.hdc, 70, 10, 30, 250, vbGreen, vbCyan, vbWhite, vbWhite, True
    DrawGradient picTwo.hdc, 100, 10, 30, 250, vbCyan, vbBlue, vbWhite, vbWhite, True
    DrawGradient picTwo.hdc, 130, 10, 30, 250, vbBlue, vbMagenta, vbWhite, vbWhite, True
    DrawGradient picTwo.hdc, 160, 10, 30, 250, vbMagenta, vbRed, vbWhite, vbWhite, True
    
    
    '' Hue Whiteness Model
    DrawGradient picThree.hdc, 10, 10, 30, 250, vbRed, vbYellow, vbBlack, vbBlack, True
    DrawGradient picThree.hdc, 40, 10, 30, 250, vbYellow, vbGreen, vbBlack, vbBlack, True
    DrawGradient picThree.hdc, 70, 10, 30, 250, vbGreen, vbCyan, vbBlack, vbBlack, True
    DrawGradient picThree.hdc, 100, 10, 30, 250, vbCyan, vbBlue, vbBlack, vbBlack, True
    DrawGradient picThree.hdc, 130, 10, 30, 250, vbBlue, vbMagenta, vbBlack, vbBlack, True
    DrawGradient picThree.hdc, 160, 10, 30, 250, vbMagenta, vbRed, vbBlack, vbBlack, True
    
    '' Hue Saturation Model
    DrawGradient picFour.hdc, 10, 10, 30, 250, vbRed, vbYellow, RGB(127, 127, 127), RGB(127, 127, 127), True
    DrawGradient picFour.hdc, 40, 10, 30, 250, vbYellow, vbGreen, RGB(127, 127, 127), RGB(127, 127, 127), True
    DrawGradient picFour.hdc, 70, 10, 30, 250, vbGreen, vbCyan, RGB(127, 127, 127), RGB(127, 127, 127), True
    DrawGradient picFour.hdc, 100, 10, 30, 250, vbCyan, vbBlue, RGB(127, 127, 127), RGB(127, 127, 127), True
    DrawGradient picFour.hdc, 130, 10, 30, 250, vbBlue, vbMagenta, RGB(127, 127, 127), RGB(127, 127, 127), True
    DrawGradient picFour.hdc, 160, 10, 30, 250, vbMagenta, vbRed, RGB(127, 127, 127), RGB(127, 127, 127), True
        
    '' Draw Three Colour Linear Gradient
    DrawGradient picFive.hdc, 0, 0, picFive.ScaleWidth, picFive.ScaleHeight / 2, vbWhite, vbWhite, vbRed, vbRed, True
    DrawGradient picFive.hdc, 0, picFive.ScaleHeight / 2, picFive.ScaleWidth, picFive.ScaleHeight / 2, vbRed, vbRed, vbBlack, vbBlack, True
    
    ' Topleft
    DrawGradient picOne.hdc, 10, 10, 100, 100, vbRed, vbYellow, vbWhite, vbWhite, True
    'BottomRight
    DrawGradient picOne.hdc, 110, 110, 100, 100, vbWhite, vbCyan, vbMagenta, vbBlue, True
    'Top Right
    DrawGradient picOne.hdc, 110, 10, 100, 100, vbYellow, vbGreen, vbWhite, vbCyan, True
    'Bottom Left
    DrawGradient picOne.hdc, 10, 110, 100, 100, vbWhite, vbWhite, vbRed, vbMagenta, True
    
    Caption = Caption & " (Rendering all screens took - " & (GetTickCount - iTick) / 1000 & " seconds)"
    
    
End Sub

Private Sub Label1_Click()

End Sub
