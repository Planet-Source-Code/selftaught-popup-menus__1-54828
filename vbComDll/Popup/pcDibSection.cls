VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcDibSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
'Private Const BI_RLE4 = 2&
'Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8

' Handle to the current DIBSection:
Private m_hDib As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

Public Function CopyToClipboard() As Boolean
Dim lhDCDesktop As Long
Dim lHDC As Long
Dim lhBmpOld As Long
Dim hObj As Long
Dim lFmt As Long
Dim b() As Byte
Dim tBI As BITMAPINFO
Dim lPtr As Long
Dim hDibCopy As Long

    lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    If (lhDCDesktop <> 0) Then
        lHDC = CreateCompatibleDC(lhDCDesktop)
        If (lHDC <> 0) Then
            'If (bAsDIB) Then
               'MsgBox "I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!"
                ' Create a duplicate DIBSection and copy
                ' to the clipboard:
                'LSet tBI = m_tBI
                'hDibCopy = CreateDIBSection( _
                '        lhDC, _
                '        m_tBI, _
                '        DIB_RGB_COLORS, _
                '        lPtr, _
                '        0, 0)
                'If (hDibCopy <> 0) Then
                '    lhBmpOld = SelectObject(lhDC, hObj)
                '    BitBlt lhDC, 0, 0, Width, Height, m_hDC, 0, 0, vbSrcCopy
                '    SelectObject lhDC, lhBmpOld
                '    lFmt = CF_DIB
                '
                '     '....
                                    
                'Else
                '    hObj = 0
                'End If
            'Else
                ' Create a compatible bitmap and copy to
                ' the clipboard:
                hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
                If (hObj <> 0) Then
                    lhBmpOld = SelectObject(lHDC, hObj)
                    PaintPicture lHDC
                    SelectObject lHDC, lhBmpOld
                    lFmt = CF_BITMAP
                    ' Now set the clipboard to the bitmap:
                    If (OpenClipboard(0) <> 0) Then
                        EmptyClipboard
                        If (SetClipboardData(lFmt, hObj) <> 0) Then
                            CopyToClipboard = True
                        End If
                        CloseClipboard
                    End If
                    DeleteObject hObj
                End If
            'End If
            DeleteDC lHDC
        End If
        DeleteDC lhDCDesktop
    End If
End Function

Public Function CreateDIB( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lHDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lHDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If (lhDCDesktop <> 0) Then
            lHDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lHDC <> 0) Then
                lhBmpOld = SelectObject(lHDC, picThis.handle)
                LoadPictureBlt lHDC
                SelectObject lHDC, lhBmpOld
                DeleteDC lHDC
            End If
        End If
        DeleteDC lhDCDesktop
    End If
End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDib)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDib)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
End Sub


Public Sub PaintPicture( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hDc() As Long
    hDc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDib
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim x As Long, y As Long
Dim lC As Long
Dim tSA As SAFEARRAY2D
Dim xEnd As Long
    
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    Debug.Assert pIsNull(VarPtrArray(bDib()))
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:
    Randomize Timer
    
    xEnd = (Width - 1) * 3
    If (bGray) Then
        For y = 0 To m_tBI.bmiHeader.biHeight - 1
            For x = 0 To xEnd Step 3
                lC = Rnd * 255
                bDib(x, y) = lC
                bDib(x + 1, y) = lC
                bDib(x + 2, y) = lC
            Next x
        Next y
    Else
        For x = 0 To xEnd Step 3
            For y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(x, y) = 0
                bDib(x + 1, y) = Rnd * 255
                bDib(x + 2, y) = Rnd * 255
            Next y
        Next x
    End If
    
    ' Clear the temporary array descriptor
   ' This is necessary under NT4.
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDib <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDib
        End If
        DeleteDC m_hDC
    End If
    m_hDC = 0: m_hDib = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Private Function pIsNull(ByVal Addr As Long) As Boolean
    Dim liTemp As Long
    CopyMemory liTemp, ByVal Addr, 4&
    pIsNull = (liTemp = 0&)
End Function

Private Function ResampleDib(ByRef cDibTo As pcDibSection) As Boolean
Dim bDibFrom() As Byte
Dim bDibTo() As Byte

Dim tSAFrom As SAFEARRAY2D
Dim tSATo As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    Debug.Assert pIsNull(VarPtrArray(bDibFrom()))
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr
    End With
    

    Debug.Assert pIsNull(VarPtrArray(bDibTo()))
    
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

Dim xScale As Single
Dim yScale As Single

Dim x As Long, y As Long, xEnd As Long, xOut As Long

Dim fX As Single, fY As Single
Dim ifY As Long, ifX As Long
Dim dx As Single, dy As Single
Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
Dim g As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single
Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
Dim ir1 As Long, ig1 As Long, ib1 As Long
Dim ir2 As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For y = 0 To cDibTo.Height - 1
        
        fY = y * yScale
        ifY = Int(fY)
        dy = fY - ifY
        
        For x = 0 To xEnd
            fX = x * xScale
            ifX = Int(fX)
            dx = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): r1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = r1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = b1 * (1 - dy) + b3 * dy
            ir2 = r2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = b2 * (1 - dy) + b4 * dy
            ' Interpolate in y:
            r = ir1 * (1 - dx) + ir2 * dx: g = ig1 * (1 - dx) + ig2 * dx: b = ib1 * (1 - dx) + ib2 * dx
            
            ' Set output:
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            If (b < 0) Then b = 0
            If (b > 255) Then
                b = 255
            End If
            xOut = x * 3
            bDibTo(xOut, y) = b
            bDibTo(xOut + 1, y) = g
            bDibTo(xOut + 2, y) = r
            
        Next x
        
    Next y

    ' Clear the temporary array descriptor
    ' This is necessary under NT4.
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function

Public Sub Fade( _
      ByVal lAmount As Long _
   )
Dim bDib() As Byte
Dim x As Long, y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim lA As Long, lA2 As Long
Dim tSA As SAFEARRAY2D
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine
        .pvData = DIBSectionBitsPtr
    End With
    Debug.Assert pIsNull(VarPtrArray(bDib))
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = Height - 1
    xMax = Width - 1
    
    For x = 0 To (xMax * 3) Step 3
        For y = 0 To yMax
            lB = lAmount * bDib(x, y) \ 255
            lG = lAmount * bDib(x + 1, y) \ 255
            lR = lAmount * bDib(x + 2, y) \ 255
            bDib(x, y) = lB
            bDib(x + 1, y) = lG
            bDib(x + 2, y) = lR
        Next y
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub Lighten( _
      ByVal lAmount As Long _
   )
Dim bDib() As Byte
Dim x As Long, y As Long
Dim xMax As Long, yMax As Long
Dim bContinue As Boolean
Dim lB As Long, lG As Long, lR As Long
Dim h As Single, s As Single, l As Single
Dim tSA As SAFEARRAY2D
    
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine
        .pvData = DIBSectionBitsPtr
    End With
    Debug.Assert pIsNull(VarPtrArray(bDib))
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
    yMax = Height - 1
    xMax = Width - 1
    
    For x = 0 To (xMax * 3) Step 3
        For y = 0 To yMax
            Draw_HLSforRGB bDib(x + 2, y), bDib(x + 1, y), bDib(x, y), h, s, l
            l = l * (1 + (lAmount / 100))
            If (l > 1) Then l = 1
            Draw_RGBforHLS h, s, l, lR, lG, lB
            bDib(x, y) = lB
            bDib(x + 1, y) = lG
            bDib(x + 2, y) = lR
        Next y
    Next x
    
    ' clear the temporary array descriptor
    ' without destroying the local temporary array
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
        
End Sub


Private Sub Class_Terminate()
    ClearUp
End Sub

