VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6120
      Top             =   3300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type udtStar
    r As Long
    x As Long
    y As Long
    a As Single
    spd As Long
    q As Long
    w As Long
    offScreen As Long
    qi As Long
    wi As Long
End Type


Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal dX&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal fShow As Integer) As Integer
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Enum eBrushStyle
    gdiBSDibPattern = 5
    gdiBSDibPatternPt = 6
    gdiBSHatched = 2
    gdiBSNull = 1
    gdiBSPattern = 3
    gdiBSSolid = 0
End Enum
Private Type uSquare
    nSize As Long
    xPos As Long
    yPos As Long
    xSpd As Long
    ySpd As Long
    Angle As Single
    StarSpd As Long
End Type

Private mRows As Long
Private mCols As Long
Private mFixedGrid As Boolean


Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type
Private Const SRCCOPY = &HCC0020

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   right As Long
   bottom As Long
End Type
Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Const Pi As Single = 3.14159265358978
Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Const Message = "Hello !"
Const OPAQUE = 2
Const TRANSPARENT = 1
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2

Private bStop As Boolean
Private scrL As Long
Private scrT As Long
Private scrW As Long
Private scrH As Long
Private scrR As Long
Private scrB As Long
Private midX As Long
Private midY As Long
Private mBlankDIB As cDIBSection
Private mBufferDIB As cDIBSection

Private bBUILDMODE As Boolean
Private CMM As cMonitors
Private uStars() As udtStar


Private Sub ApplyBlend()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    Blend.SourceConstantAlpha = 60 '255 no blend - 0 major blurry!!
    
    CopyMemory BlendPtr, Blend, 4
    
    AlphaBlend mBufferDIB.hDC, scrL, scrT, scrW, scrH, mBlankDIB.hDC, 0, 0, scrW, scrH, BlendPtr
End Sub

Private Function CreateMyFont(nSize&, sFontFace$, bBold As Boolean, bItalic As Boolean) As Long
Static r&, d&

    DeleteDC r: r = GetDC(0)
    d = GetDeviceCaps(r, LOGPIXELSY)
    CreateMyFont = CreateFont(-MulDiv(nSize, d, 72), 0, 0, 0, _
                              IIf(bBold, FW_BOLD, FW_NORMAL), bItalic, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                              CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, sFontFace) 'gdi 2
End Function
Private Sub SplitRGB(ByVal clr&, r&, g&, b&)
    r = clr And &HFF: g = (clr \ &H100&) And &HFF: b = (clr \ &H10000) And &HFF
End Sub
Private Sub SetFont(DC&, sFace$, nSize&)
Static c&
    ReleaseDC DC, c: DeleteDC c
    c = CreateMyFont(nSize, sFace, False, False)
    DeleteObject SelectObject(DC, c)
End Sub

Private Sub Form_Click()
bStop = True
End Sub

Private Sub Form_DblClick()
bStop = True
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
bStop = True
End Sub

Private Sub Form_Load()

If Mid$(Command(), 2, 1) = "p" Then End
If Mid$(Command(), 2, 1) = "c" Then
    MsgBox "There are no settings for this screen saver." & vbCrLf & vbCrLf & "Michael Toye" & vbCrLf & "michael_toye@yahoo.co.uk", vbInformation, "Dotty Starfield"
    End
End If
Randomize Timer

bBUILDMODE = True

If Not bBUILDMODE Then
    Set CMM = New cMonitors
    scrL = CMM.VirtualScreenLeft
    scrT = CMM.VirtualScreenTop
    scrW = CMM.VirtualScreenWidth
    scrH = CMM.VirtualScreenHeight
Else
    scrL = 0
    scrT = 0
    scrW = Me.Width \ Screen.TwipsPerPixelX
    scrH = Me.Height \ Screen.TwipsPerPixelY
End If
scrR = scrW - scrL
scrB = scrH - scrT
midX = scrL + (scrW / 2)
midY = scrT + (scrH / 2)

CreateDIB mBlankDIB, scrW, scrH
CreateDIB mBufferDIB, scrW, scrH

SetupStars

Timer1.Enabled = True
End Sub

Sub SetTopmostWindow(ByVal hwnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hwnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Function GimmeX(ByVal aIn As Single, lIn As Long) As Long
'(pi/180)
    GimmeX = sIN(aIn * 0.01745329251994) * lIn

End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Long
'(pi/180)
    GimmeY = Cos(aIn * 0.01745329251994) * lIn
End Function
Function Sine(Degrees_Arg)
'Atn(1) / 45
Sine = sIN(Degrees_Arg * 0.01745329251994)
End Function

Function Cosine(Degrees_Arg)
'Atn(1) / 45
Cosine = Cos(Degrees_Arg * 0.01745329251994)
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
    bStop = True
End Sub

Private Sub COUT(sIN$, x&, y&)

    SetTextColor mBufferDIB.hDC, RGB(50, 50, 50)
    TextOut mBufferDIB.hDC, x, y, sIN, Len(sIN)

End Sub
Private Sub CreateDIB(ByRef tDIB As cDIBSection, scrW&, scrH&)
Set tDIB = New cDIBSection
With tDIB
    .Create scrW, scrH
    SetBkMode .hDC, TRANSPARENT
    SetFont .hDC, "Tahoma", 8
End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static X0 As Integer, Y0 As Integer
    If Not bBUILDMODE Then
        If ((X0 = 0) And (Y0 = 0)) Or _
           ((Abs(X0 - x) < 8) And (Abs(Y0 - y) < 8)) Then ' small mouse movement...
            X0 = x                          ' Save current x coordinate
            Y0 = y                          ' Save current y coordinate
            Exit Sub                        ' Exit
        End If
        bStop = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mBlankDIB = Nothing
    Set mBufferDIB = Nothing

    Set CMM = Nothing
    ShowCursor -1
    Screen.MousePointer = vbDefault
End Sub
Sub SetupStars()
Dim n&, st&
    ReDim uStars(20)
    For n = 0 To UBound(uStars)
        With uStars(n)
            .a = Int(Rnd * 360)
            st = Int(Rnd * scrW * 0.2)
            .x = midX + GimmeX(.a, st)
            .y = midY + GimmeY(.a, st)
            .r = 10
            .q = Int(Rnd * 360)
            .w = Int(Rnd * 360)
            .spd = 2 + Int(Rnd * 40)
            .qi = (1 + Int(Rnd * 5)) * IIf((Rnd * 1000) > 600, -1, 1)
            .wi = (1 + Int(Rnd * 5)) * IIf((Rnd * 1000) > 600, -1, 1)
            .offScreen = 0
        End With
    Next
End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    DoEvents
    
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    If Not bBUILDMODE Then
        Me.Move scrL * Screen.TwipsPerPixelX, scrT * Screen.TwipsPerPixelY, scrW * Screen.TwipsPerPixelX, scrH * Screen.TwipsPerPixelY
        SetTopmostWindow Me.hwnd
        ShowCursor 0
    Else
        Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - (Me.Height * 1.2))
    End If

    

    DisplaySS
End Sub

Private Sub DisplaySS()

Dim n&, s&, st&, t&, f&, fp&, ag!
't = GetTickCount: f = 0 'reinstate for FPS
    Do
        
        BitBlt mBufferDIB.hDC, scrL, scrT, scrW, scrH, mBlankDIB.hDC, 0, 0, vbSrcCopy
        'ApplyBlend
        
        
        'reinstate for FPS  -->
'        f = f + 1
'        If GetTickCount - t >= 1000 Then
'            t = GetTickCount: fp = f: f = 0
'        End If
'        COUT CStr(fp), 3, 3
        '<-- reinstate for FPS
        
        For n = 0 To UBound(uStars)
            With uStars(n)
                PlotBall mBufferDIB.hDC, .x, .y, .r, .q, .w, .offScreen
            End With
        Next
        BitBlt Me.hDC, scrL, scrT, scrW, scrH, mBufferDIB.hDC, 0, 0, vbSrcCopy
        
        For n = 0 To UBound(uStars)
            With uStars(n)
                .q = .q + .qi
                If .q > 360 Then .q = .q - 360
                If .q < 0 Then .q = .q + 360
                .w = .w + .wi
                If .w > 360 Then .w = .w - 360
                If .w < 0 Then .w = .w + 360
                ag = .a * 0.01745329251994
                .x = .x + (sIN(ag) * .spd)
                .y = .y + (Cos(ag) * .spd)
                 .r = .r + 1
                If .offScreen = 1 Then
                    .a = Int(Rnd * 360)
                    ag = .a * 0.01745329251994
                    st = Int(Rnd * scrW * 0.2)
                    .x = midX + (sIN(ag) * st)
                    .y = midY + (Cos(ag) * st)
                    .r = 10
                    .q = Int(Rnd * 360)
                    .w = Int(Rnd * 360)
                    .spd = 2 + Int(Rnd * 10)
                    .offScreen = 0
                End If
            End With
        Next

        DoEvents
        Sleep 30
    Loop Until bStop

    ShowCursor -1
    Screen.MousePointer = vbDefault
    End
End Sub
Sub PlotBall(ByRef DC&, xOff&, yOff&, r&, q&, w&, ByRef OSc&)
Dim y&, x&, x1&, z&, Ta2!, Ta1&, c&, Ta1a!, Ta2a!

    c = r

    If c > 255 Then c = 255
    OSc = 1
    For Ta1 = 15 To 165 Step 15
        y = Cos(Ta1 * 0.01745329251994) * r
        x = sIN(Ta1 * 0.01745329251994) * r
        For Ta2 = 0 To 340 Step 20
            x1 = sIN(Ta2 * 0.01745329251994) * x
            z = Cos(Ta2 * 0.01745329251994) * x
            PlotPt DC, xOff, yOff, x1, y, z, q, w, 2000, 5000, RGB(c, c, c), OSc
        Next
    Next

End Sub

Sub PlotPt(ByRef DC&, xOff&, yOff&, x1 As Long, y1 As Long, z1 As Long, Theta As Long, Alt As Long, Size As Long, Perspective As Long, lColor As Long, ByRef OSc&)
Dim cX As Single, cY As Single
Dim vX As Single, vY As Single, vZ As Single
Dim pX1 As Single, pY1 As Single
Dim pX2 As Single, pY2 As Single
Dim Phi As Single
Dim Sin_Theta As Single, Cos_Theta As Single, Sin_Phi   As Single, Cos_Phi   As Single
    
    Phi = 180 - Alt: Sin_Theta = sIN(Theta * 0.01745329251994): Cos_Theta = Cos(Theta * 0.01745329251994):
    Sin_Phi = sIN(Phi * 0.01745329251994): Cos_Phi = Cos(Phi * 0.01745329251994)
    vX = -x1 * Sin_Theta + y1 * Cos_Theta
    vY = -x1 * Cos_Theta * Cos_Phi - y1 * Sin_Theta * Cos_Phi + z1 * Sin_Phi
    vZ = -x1 * Cos_Theta * Sin_Phi - y1 * Sin_Theta * Sin_Phi - z1 * Cos_Phi + Perspective
    pX1 = xOff + Size * vX / vZ: pY1 = yOff - Size * vY / vZ

    If (pX1 < scrL Or pX1 > scrR) Or (pY1 < scrT Or pY1 > scrB) Then Exit Sub
    
    SetPixel DC, CLng(pX1), CLng(pY1), lColor
    OSc = 0
End Sub

