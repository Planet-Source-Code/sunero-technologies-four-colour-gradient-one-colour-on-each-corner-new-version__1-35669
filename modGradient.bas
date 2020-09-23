Attribute VB_Name = "modGradient"

    '*******************************************************************************
    ' MODULE:       modGradient
    ' AUTHOR:       Rohit Kulshreshtha
    ' CREATED:      04-09-2002
    ' COPYRIGHT:    Copyright 2002 Rohit Kulshreshtha. All Rights Reserved.
    '
    ' DESCRIPTION:
    '
    ' This module is totally independent. Just drop it in and start using.
    '*******************************************************************************
Option Explicit

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type cRGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0

    '*******************************************************************************
    ' DrawGradient (FUNCTION)
    '
    ' DESCRIPTION:
    ' This function is used to draw gradients with four colours
    '
    ' Arguments:
    ' hDC - The device to draw on
    ' Top - Distance in pixels, from top
    ' Left - Distance in pixels, from left
    ' Width - In pixels
    ' Height - In pixels
    ' colourTopLeft - The colour of the top-left corner
    ' colourTopRight - The colour of the top-right corner
    ' colourBottomLeft - The colour of the bottom-left corner
    ' colourBottomRight - The colour of the bottom-right corner
    '*******************************************************************************
Public Function DrawGradient(hdc As Long, Left As Long, Top As Long, Width As Long, Height As Long, colourTopLeft As Long, colourTopRight As Long, colourBottomLeft As Long, colourBottomRight As Long, Optional bSmooth As Boolean = False)
    Dim bi24BitInfo     As BITMAPINFO
    Dim bBytes()        As Byte
    Dim LeftGrads()     As cRGB
    Dim RightGrads()    As cRGB
    Dim MiddleGrads()   As cRGB
    Dim TopLeft         As cRGB
    Dim TopRight        As cRGB
    Dim BottomLeft      As cRGB
    Dim BottomRight     As cRGB
    Dim iLoop           As Long
    Dim bytesWidth      As Long
    Dim iShift          As Long
    Dim iMode           As Long
    Dim iRaise          As Integer
    
    With TopLeft
        .Red = Red(colourTopLeft)
        .Green = Green(colourTopLeft)
        .Blue = Blue(colourTopLeft)
    End With
    
    With TopRight
        .Red = Red(colourTopRight)
        .Green = Green(colourTopRight)
        .Blue = Blue(colourTopRight)
    End With
    
    With BottomLeft
        .Red = Red(colourBottomLeft)
        .Green = Green(colourBottomLeft)
        .Blue = Blue(colourBottomLeft)
    End With
    
    With BottomRight
        .Red = Red(colourBottomRight)
        .Green = Green(colourBottomRight)
        .Blue = Blue(colourBottomRight)
    End With
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        
        If Width >= 255 Then
            .biWidth = 255
        Else
            iRaise = 1
            Do Until (2 ^ iRaise) - Width > 0
                iRaise = iRaise + 1
            Loop
            .biWidth = 2 ^ iRaise - 1
        End If
        
        If Height >= 255 Then
            .biHeight = 255
        Else
            iRaise = 1
            Do Until (2 ^ iRaise) - Height > 0
                iRaise = iRaise + 1
            Loop
            .biHeight = 2 ^ iRaise - 1
        End If
        
    End With
    GradateColoursRGB LeftGrads, bi24BitInfo.bmiHeader.biHeight, BottomLeft, TopLeft
    GradateColoursRGB RightGrads, bi24BitInfo.bmiHeader.biHeight, BottomRight, TopRight
    
    ReDim bBytes(0 To (bi24BitInfo.bmiHeader.biHeight * bi24BitInfo.bmiHeader.biWidth * 3) + (bi24BitInfo.bmiHeader.biHeight * 3)) As Byte
   
    For iLoop = 0 To bi24BitInfo.bmiHeader.biHeight - 1
        GradateColoursRGB MiddleGrads, bi24BitInfo.bmiHeader.biWidth, LeftGrads(iLoop), RightGrads(iLoop)
        CopyMemory bBytes(iLoop * bi24BitInfo.bmiHeader.biWidth * 3 + (iLoop * 3)), MiddleGrads(0), bi24BitInfo.bmiHeader.biWidth * 3
    Next iLoop
    
    If bSmooth Then iMode = SetStretchBltMode(hdc, 4)
    StretchDIBits hdc, Left, Top, Width, Height, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, bBytes(0), bi24BitInfo, BI_RGB, vbSrcCopy
    If bSmooth Then SetStretchBltMode hdc, iMode
        
    
End Function
    '*******************************************************************************
    ' GradateColours (FUNCTION)
    '
    ' DESCRIPTION:
    ' This function is to blend colour1 to colour2
    '*******************************************************************************
Private Function GradateColoursRGB(cResults() As cRGB, Length As Long, Colour1 As cRGB, Colour2 As cRGB)
    Dim fromR   As Integer
    Dim toR     As Integer
    Dim fromG   As Integer
    Dim toG     As Integer
    Dim fromB   As Integer
    Dim toB     As Integer
    Dim stepR   As Single
    Dim stepG   As Single
    Dim stepB   As Single
    Dim iLoop   As Long
    
    ReDim cResults(0 To Length)
    
    fromR = Colour1.Red
    fromG = Colour1.Green
    fromB = Colour1.Blue
    
    toR = Colour2.Red
    toG = Colour2.Green
    toB = Colour2.Blue
    
    stepR = Divide(toR - fromR, Length)
    stepG = Divide(toG - fromG, Length)
    stepB = Divide(toB - fromB, Length)
    
    For iLoop = 0 To Length
        cResults(iLoop).Red = fromR + (stepR * iLoop)
        cResults(iLoop).Green = fromG + (stepG * iLoop)
        cResults(iLoop).Blue = fromB + (stepB * iLoop)
    Next iLoop
End Function

    '*******************************************************************************
    ' Blue (FUNCTION)
    '
    ' DESCRIPTION:
    ' Retrieve Blue from Long
    '*******************************************************************************
Private Function Blue(Colour As Long) As Long
    Blue = (Colour And &HFF0000) / &H10000
End Function

    '*******************************************************************************
    ' Green (FUNCTION)
    '
    ' DESCRIPTION:
    ' Retrieve Green as long
    '*******************************************************************************
Private Function Green(Colour As Long) As Long
    Green = (Colour And &HFF00&) / &H100
End Function

    '*******************************************************************************
    ' Red (FUNCTION)
    '
    ' DESCRIPTION:
    ' Retrieve Red from Long
    '*******************************************************************************
Private Function Red(Colour As Long) As Long
    Red = (Colour And &HFF&)
End Function

    '*******************************************************************************
    ' Divide (FUNCTION)
    '
    ' DESCRIPTION:
    ' Division function to avoid division by 0 error
    '*******************************************************************************
Private Function Divide(Numerator, Denominator) As Single
    If Numerator = 0 Or Denominator = 0 Then
        Divide = 0
    Else
        Divide = Numerator / Denominator
    End If
End Function


