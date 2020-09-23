VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Screensaver"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "vbRipple.frx":0000
   ScaleHeight     =   4140
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   75
      Left            =   960
      Picture         =   "vbRipple.frx":004E
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox say 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   1920
      Picture         =   "vbRipple.frx":009C
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos& Lib "user32.dll" (lpPoint As POINTAPI)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type pt
    X As Long
    Y As Long
End Type
Private Type POINTAPI
    X As Integer
    z As Integer
    Y As Integer
End Type

Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6

Public DD As DirectDraw
Private Lookup3(256) As Long
Private Lookup4(300, 125) As Long
Private Lookup5(300, 125) As Long
Private Lookup7(250) As pt
Private LSin(1000)
Private LCos(1000)

Private LookupBulge(800, 600) As pt
Private LookupSpirals(800, 600) As pt
Private LookupFlyout(800, 600) As pt
Private LookupMove(800, 600) As pt
Private LookupRollOut(800, 600) As pt
Private LookupRollIn(800, 600) As pt

Public qq, rr, var, jj       '}____  explained in
Public pp, fps, w, h, wind   '}      form_load
Private pt As POINTAPI

Sub DoEffect()
On Error Resume Next

Dim pict() As Byte
Dim pict2() As Byte
Dim pict3() As Byte

Dim sa As SAFEARRAY2D, bmp As BITMAP
Dim sa2 As SAFEARRAY2D, bmp2 As BITMAP
Dim sa3 As SAFEARRAY2D, bmp3 As BITMAP

Dim r As Integer, c As Integer, nc As Integer, pos As Integer

'info on bitmaps in each buffer
GetObjectAPI Form1.Picture, Len(bmp), bmp
GetObjectAPI p1.Picture, Len(bmp2), bmp2
GetObjectAPI say.Picture, Len(bmp3), bmp3

'must be 8bpp
If bmp.bmPlanes <> 1 Or bmp.bmBitsPixel <> 8 Then
    If wind = 0 Then
        ShowCursor 1
    End If
    End
End If

'point to pixels of each buffer
With sa
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp.bmWidthBytes
    .pvData = bmp.bmBits
End With
CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4

With sa2
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp2.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp2.bmWidthBytes
    .pvData = bmp2.bmBits
End With
CopyMemory ByVal VarPtrArray(pict2), VarPtr(sa2), 4

With sa3
    .cbElements = 1
    .cDims = 2
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = bmp3.bmHeight
    .Bounds(1).lLbound = 0
    .Bounds(1).cElements = bmp3.bmWidthBytes
    .pvData = bmp3.bmBits
End With
CopyMemory ByVal VarPtrArray(pict3), VarPtr(sa3), 4

'adjust variables
rr = rr + qq
If rr = 250 Then
    qq = -qq
    var = Int(Rnd * 6)
    hh = Int(Rnd * 7) + 1
End If
If rr = 0 Then
    qq = -qq
    var = Int(Rnd * 6)
    hh = Int(Rnd * 7) + 1
End If

If pp <> -1 Then
    var = pp
End If
pp = -1

'do transformations
If var = 0 Then
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupBulge(c%, r%).X, LookupBulge(c%, r%).Y))
    Next
Next
ElseIf var = 1 Then
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupSpirals(c%, r%).X, LookupSpirals(c%, r%).Y))
    Next
Next
ElseIf var = 2 Then
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupFlyout(c%, r%).X, LookupFlyout(c%, r%).Y))
    Next
Next
ElseIf var = 3 Then
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupMove(c%, r%).X, LookupMove(c%, r%).Y))
    Next
Next
ElseIf var = 4 Then
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupRollOut(c%, r%).X, LookupRollOut(c%, r%).Y))
    Next
Next
Else
For c% = 0 To UBound(pict, 1)
    For r% = 0 To UBound(pict, 2)
        pict(c%, r%) = (pict2(LookupRollIn(c%, r%).X, LookupRollIn(c%, r%).Y))
    Next
Next
End If

'draw circle
For i% = 1 To 300
    pict(Int(Rnd * 20) + Lookup4(i%, Int(rr / 2)), Int(Rnd * 20) + Lookup5(i%, Int(rr / 2))) = 255
Next i%
'draw Line
For i% = 1 To UBound(pict, 2)
    pict(Int(Rnd * 20) + rr, i%) = 255
    pict(Int(Rnd * 20) + rr, i%) = 255
Next i%

'draw flying dot
For i% = -5 To 5
For X% = -5 To 5
    gg% = Int(Rnd * 2)
    If gg% = 0 Then
        pict(Lookup7(rr).X + i%, 200 - Lookup7(rr).Y + X%) = 255
    End If
Next X%
Next i%

'see whether to flash lightning (1/50 chance per frame)
jj = Int(Rnd * 50)

If jj = 1 Then
For i% = 0 To UBound(pict, 1)
For X% = 0 To UBound(pict, 2)
    If pict3(i%, X%) <> 0 Then
        pict(i%, X%) = pict3(i%, X%)
    End If
Next X%
Next i%
End If

jj = 0

'blur and fade
For c% = 3 To UBound(pict2, 1) - 3 Step 1
    For r% = 1 To UBound(pict2, 2) - 1 Step 1
        pict2(c%, r%) = Lookup3((CInt(pict(c%, r%)) + _
        CInt(pict(c% - 1, r%)) + _
        CInt(pict(c% + 1, r%)) + _
        CInt(pict(c%, r% - 1)) + _
        CInt(pict(c%, r% + 1))) \ 5)
    Next
Next

CopyMemory ByVal VarPtrArray(pict), 0&, 4
CopyMemory ByVal VarPtrArray(pict2), 0&, 4
CopyMemory ByVal VarPtrArray(pict3), 0&, 4

Form1.Refresh

'changes the palette with bitmaps
'must replace with dx palette stuff later
If hh = 2 Then
    Form1.Picture = LoadPicture(App.Path & "/grey.bmp")
    p1.Picture = LoadPicture(App.Path & "/grey.bmp")
ElseIf hh = 3 Then
    Form1.Picture = LoadPicture(App.Path & "/rose.bmp")
    p1.Picture = LoadPicture(App.Path & "/rose.bmp")
ElseIf hh = 4 Then
    Form1.Picture = LoadPicture(App.Path & "/bluegreen.bmp")
    p1.Picture = LoadPicture(App.Path & "/bluegreen.bmp")
ElseIf hh = 5 Then
    Form1.Picture = LoadPicture(App.Path & "/fire.bmp")
    p1.Picture = LoadPicture(App.Path & "/fire.bmp")
ElseIf hh = 6 Then
    Form1.Picture = LoadPicture(App.Path & "/yellow.bmp")
    p1.Picture = LoadPicture(App.Path & "/yellow.bmp")
End If
hh = 1
End Sub


Private Sub Form_DblClick()
'end program
If wind = 0 Then
    ShowCursor 1
End If
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
pp = Int(Chr$(KeyAscii)) - 1
End Sub

Private Sub Form_Load()
hh = Int(Rnd * 5) + 1 'loads random palette at start
If hh = 1 Then
    Form1.Picture = LoadPicture(App.Path & "/grey.bmp")
    p1.Picture = LoadPicture(App.Path & "/grey.bmp")
ElseIf hh = 2 Then
    Form1.Picture = LoadPicture(App.Path & "/rose.bmp")
    p1.Picture = LoadPicture(App.Path & "/rose.bmp")
ElseIf hh = 3 Then
    Form1.Picture = LoadPicture(App.Path & "/bluegreen.bmp")
    p1.Picture = LoadPicture(App.Path & "/bluegreen.bmp")
ElseIf hh = 4 Then
    Form1.Picture = LoadPicture(App.Path & "/fire.bmp")
    p1.Picture = LoadPicture(App.Path & "/fire.bmp")
ElseIf hh = 5 Then
    Form1.Picture = LoadPicture(App.Path & "/yellow.bmp")
    p1.Picture = LoadPicture(App.Path & "/yellow.bmp")
End If

wind = 0 'determines whether program will run in window or fullscreen


Randomize

w = 320  '}___ width and height
h = 200  '}    of screen

If wind = 0 Then
    ShowCursor 0
End If

Form1.Width = (w * Screen.TwipsPerPixelX)
Form1.Height = (h * Screen.TwipsPerPixelY)
p1.Width = (w * Screen.TwipsPerPixelX)
p1.Height = (h * Screen.TwipsPerPixelY)

If wind = 0 Then
    'direct draw stuff to change res
    DirectDrawCreate ByVal 0&, DD, Nothing
    DD.SetCooperativeLevel Me.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
    DD.SetDisplayMode 320, 200, 16
End If

Form1.Visible = True

If wind = 1 Then
    StayOnTop Me
End If

Print "     Loading..."

SetLooks 'create lookup tables

hh = 1 'determines which palette
rr = 1 'increases/decreases by qq each frame  }__  used to create line going back and forth
qq = 1 'amount to change rr                   }    and circle going large and small
pp = -1 'value for new transformation when key is pressed
var = Int(Rnd * 6) 'determines which transformation

'main loop
Do
DoEvents
DoEffect
Loop
End Sub

Sub SetLooks()
'transformation lookup tables
For i = 0 To w
For X = 0 To h
    LookupBulge(i, X).X = 1 + i + (Sin(i / 50) * -30 * Cos(X / 50)) / 4
    If LookupBulge(i, X).X >= w Then
        LookupBulge(i, X).X = w - 1
    End If
    If LookupBulge(i, X).X < 0 Then
        LookupBulge(i, X).X = 0
    End If
    
    LookupBulge(i, X).Y = 0 + X + (Cos(i / 50) * -30 * Sin(X / 50)) / 4
    If LookupBulge(i, X).Y >= h Then
        LookupBulge(i, X).Y = h - 1
    End If
    If LookupBulge(i, X).Y < 0 Then
        LookupBulge(i, X).Y = 0
    End If
Next X
Next i

For i = 0 To w
For X = 0 To h
    LookupSpirals(i, X).X = 1 + i + (Sin(i / 50) * -100 * Cos(X / 50)) / 8
    If LookupSpirals(i, X).X >= w Then
        LookupSpirals(i, X).X = w - 1
    End If
    If LookupSpirals(i, X).X < 0 Then
        LookupSpirals(i, X).X = 0
    End If
    
    LookupSpirals(i, X).Y = 2 + X + (Cos(i / 50) * 100 * Sin(X / 50)) / 16
    If LookupSpirals(i, X).Y >= h Then
        LookupSpirals(i, X).Y = h - 1
    End If
    If LookupSpirals(i, X).Y < 0 Then
        LookupSpirals(i, X).Y = 0
    End If
Next X
Next i

For i = 0 To w
For X = 0 To h
    LookupFlyout(i, X).X = ((i / 1.1) + 15)
    If LookupFlyout(i, X).X >= w Then
        LookupFlyout(i, X).X = w - 1
    End If
    If LookupFlyout(i, X).X < 0 Then
        LookupFlyout(i, X).X = 0
    End If
    
    LookupFlyout(i, X).Y = ((X / 1.1) + 9)
    If LookupFlyout(i, X).Y >= h Then
        LookupFlyout(i, X).Y = h - 1
    End If
    If LookupFlyout(i, X).Y < 0 Then
        LookupFlyout(i, X).Y = 0
    End If
Next X
Next i

For i = 0 To w
For X = 0 To h
    LookupMove(i, X).X = i + (Sin(i) * 5)
    If LookupMove(i, X).X >= w Then
        LookupMove(i, X).X = w - 1
    End If
    If LookupMove(i, X).X < 0 Then
        LookupMove(i, X).X = 0
    End If
    
    LookupMove(i, X).Y = X + (Sin(X) * 5)
    If LookupMove(i, X).Y >= h Then
        LookupMove(i, X).Y = h - 1
    End If
    If LookupMove(i, X).Y < 0 Then
        LookupMove(i, X).Y = 0
    End If
Next X
Next i

asin = Sin(1.5) * 1.05
acos = Cos(1.5) * 1.05
For i = 0 To w
For X = 0 To h
    LookupRollOut(i, X).X = (w / 2) + (i - (w / 2)) * asin - (X - (h / 2)) * acos
    If LookupRollOut(i, X).X >= w Then
        LookupRollOut(i, X).X = w - 1
    End If
    If LookupRollOut(i, X).X < 0 Then
        LookupRollOut(i, X).X = 0
    End If
    
    LookupRollOut(i, X).Y = (h / 2) + (X - (h / 2)) * asin + (i - (w / 2)) * acos
    If LookupRollOut(i, X).Y >= h Then
        LookupRollOut(i, X).Y = h - 1
    End If
    If LookupRollOut(i, X).Y < 0 Then
        LookupRollOut(i, X).Y = 0
    End If
Next X
Next i

asin = Sin(1.5) * 0.9
acos = Cos(1.5) * 0.9
For i = 0 To 320
For X = 0 To 200
    LookupRollIn(i, X).X = (w / 2) + (i - (w / 2)) * asin - (X - (h / 2)) * acos
    If LookupRollIn(i, X).X >= 320 Then
        LookupRollIn(i, X).X = 319
    End If
    If LookupRollIn(i, X).X < 0 Then
        LookupRollIn(i, X).X = 0
    End If
    
    LookupRollIn(i, X).Y = (h / 2) + (X - (h / 2)) * asin + (i - (w / 2)) * acos
    If LookupRollIn(i, X).Y >= 200 Then
        LookupRollIn(i, X).Y = 199
    End If
    If LookupRollIn(i, X).Y < 0 Then
        LookupRollIn(i, X).Y = 0
    End If
Next X
Next i

'fade lookup table
For i = 0 To 256
    Lookup3(i) = i - 3
    If Lookup3(i) < 0 Then
        Lookup3(i) = 0
    End If
Next

'lookup table for position of dot
For i = 0 To 250
    Lookup7(i).X = Sin(i / 10) * 100 + (w / 2)
    If Lookup7(i).X < 0 Then
        Lookup7(i).X = 0
    End If
    If Lookup7(i).X > 320 Then
        Lookup7(i).X = 319
    End If
    
    Lookup7(i).Y = Cos(i / 50) * 50 + (h / 2)
    If Lookup7(i).Y < 0 Then
        Lookup7(i).Y = 0
    End If
    If Lookup7(i).Y > 320 Then
        Lookup7(i).Y = 319
    End If
Next

'lookup table for circle X
For i = 0 To 300
For X = 0 To 125
    Lookup4(i, X) = (Sin(i) * X / 2) + (w / 2)
Next X
Next i

'lookup table for circle Y
For X = 0 To 125
For i = 0 To 300
    Lookup5(i, X) = (Cos(i) * X / 2) + (h / 2)
Next i
Next X
End Sub

Sub StayOnTop(TheForm As Form) 'make form stay on top, if window
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
