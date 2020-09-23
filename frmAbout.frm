VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Rohit's Four Colour Gradients"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox picCaption 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rohito's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   1380
         TabIndex        =   1
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "© Rohit Kulshreshtha"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1140
      Width           =   3195
   End
   Begin VB.Label lblOne 
      Caption         =   "Rohit's Four Colour Gradient"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   900
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    DrawGradient picCaption.hdc, 0, 0, 349, 48, vbRed, vbGreen, vbYellow, vbBlue, True
    DrawGradient picCaption.hdc, 0, 48, 174, 2, vbYellow, vbWhite, vbYellow, vbWhite, True
    DrawGradient picCaption.hdc, 174, 48, 175, 2, vbWhite, vbBlue, vbWhite, vbBlue, True
End Sub
