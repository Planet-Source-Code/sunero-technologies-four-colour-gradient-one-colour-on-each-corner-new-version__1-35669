VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmExample 
   Caption         =   "Rohit's Four Colour Gradient Test"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   707
      TabIndex        =   5
      Top             =   7725
      Width           =   10605
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   4140
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.CheckBox chkSmooth 
         Caption         =   "Smoothing (Slower)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmdExTwo 
         Caption         =   "Second Example"
         Height          =   375
         Left            =   2580
         TabIndex        =   7
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblRender 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1305
      End
   End
   Begin MSComDlg.CommonDialog cDLG 
      Left            =   6660
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHandleD 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   4
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox picHandleC 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10140
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   3
      Top             =   6900
      Width           =   375
   End
   Begin VB.PictureBox picHandleB 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10140
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox picHandleA 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   60
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   707
      TabIndex        =   0
      Top             =   0
      Width           =   10605
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub chkSmooth_Click()
    DrawG
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdExTwo_Click()
    frmExampleTwo.Show
End Sub

Private Sub Form_Load()
    DrawG
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    picBox.Height = ScaleHeight - picBar.ScaleHeight
    picBox.Width = Me.ScaleWidth
    picHandleA.Move 0, 0
    picHandleB.Move picBox.ScaleWidth - picHandleB.ScaleWidth - 2, 0
    picHandleC.Move picBox.ScaleWidth - picHandleC.ScaleWidth - 2, picBox.ScaleHeight - picHandleC.ScaleWidth - 2
    picHandleD.Move 0, picBox.ScaleHeight - picHandleD.ScaleWidth - 2
    DrawG
End Sub

Private Sub picHandleA_Click()
    cDLG.Color = picHandleA.BackColor
    cDLG.ShowColor
    picHandleA.BackColor = cDLG.Color
    DrawG
End Sub

Private Function DrawG()
    Dim iTick As Long
    iTick = GetTickCount
    picBox.Cls
    If chkSmooth.Value = 1 Then
        DrawGradient picBox.hdc, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight, picHandleA.BackColor, picHandleB.BackColor, picHandleD.BackColor, picHandleC.BackColor, True
    Else
        DrawGradient picBox.hdc, 0, 0, picBox.ScaleWidth, picBox.ScaleHeight, picHandleA.BackColor, picHandleB.BackColor, picHandleD.BackColor, picHandleC.BackColor
    End If
    picBox.Refresh
    lblRender = "Rendering took: " & ((GetTickCount - iTick) / 1000) & " seconds."
End Function

Private Sub picHandleB_Click()
    cDLG.Color = picHandleB.BackColor
    cDLG.ShowColor
    picHandleB.BackColor = cDLG.Color
    DrawG
End Sub

Private Sub picHandleC_Click()
    cDLG.Color = picHandleC.BackColor
    cDLG.ShowColor
    picHandleC.BackColor = cDLG.Color
    DrawG
End Sub

Private Sub picHandleD_Click()
    cDLG.Color = picHandleD.BackColor
    cDLG.ShowColor
    picHandleD.BackColor = cDLG.Color
    DrawG
End Sub
