VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Binding bitmap with DC
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Create DC that compatible to current device
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Create Bitmap that compatible to specified DC
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, _
    ByVal handle As Long, ByVal dw As Long) As Long
'Delete DC
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'Delete Bitmap
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Paint
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Copy memory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

'Bitmap info header
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

'RGB color table
Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

'Bitmap info
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Const DIB_RGB_COLORS = 0                        'Color table
Private Const BITBLT_TRANSPARENT_WINDOWS = &H40000000   'Enables to capture transparent windows

Private bi          As BITMAPINFO                       'Bitmap info
Private hhDC        As Long                             'Memory DC handle
Private hhBmp       As Long                             'Memory Bitmap handle
Private lpData      As Long                             'Pointer to the bitmap data
Private bSize       As Long                             'Memory size of the Bitmap data (in bytes)

'Bitmap width
Public Property Get iWidth() As Long
    iWidth = bi.bmiHeader.biWidth
End Property

'Bitmap height
Public Property Get iHeight() As Long
    iHeight = bi.bmiHeader.biHeight
End Property

'Bitmap color bit count
Public Property Get iBitCount() As Integer
    iBitCount = bi.bmiHeader.biBitCount
End Property

'Bitmap data size (in bytes)
Public Property Get iImageSize() As Long
    iImageSize = bi.bmiHeader.biSizeImage
End Property

'Handle to the created DC
Public Property Get hDC() As Long
    hDC = hhDC
End Property

'Handle to the created Bitmap
Public Property Get hBmp() As Long
    hBmp = hhBmp
End Property

'Pointer to the Bitmap data
Public Property Get lpBitData() As Long
    lpBitData = lpData
End Property

'To create memory DC
'Args:   Width, Height: Width and height of the memory DC£¨In pixels£©, respectively
'        BitCount: Color bit count£¬can be 0, 1, 4, 8, 16, 24, 32. For jpg or png format, color bit should be 0
'        hDCfrom: Create DC that compatibles to the specified DC, default = 0
'
'Return: True if succeed, False if failed
Public Function CreateMemDC(ByVal iWidth As Long, ByVal iHeight As Long, _
    Optional ByVal iBitCount As Integer = 16, Optional ByVal FromHdc As Long = 0) As Boolean
    
    'Delete previous memory DC and Bitmap
    If hhDC <> 0 Or hhBmp <> 0 Then
        Call DeleteMemDC
    End If
    
    'Set info of the Bitmap
    With bi.bmiHeader
        .biBitCount = iBitCount
        .biWidth = iWidth
        .biHeight = iHeight
        .biSize = Len(bi)
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * .biBitCount / 8
        bSize = .biSizeImage
    End With
    
    'Create memory DC
    hhDC = CreateCompatibleDC(FromHdc)
    
    'Create memory Bitmap
    hhBmp = CreateDIBSection(hhDC, bi, DIB_RGB_COLORS, ByVal VarPtr(lpData), 0, 0)
    
    'Bind Bitmap and DC
    SelectObject hhDC, hhBmp
    
    CreateMemDC = (hhBmp <> 0)
End Function

'To delete created memory DC and Bitmap
Public Sub DeleteMemDC()
    If hhDC <> 0 Then
        DeleteDC hhDC
    End If
    If hhBmp <> 0 Then
        DeleteObject hhBmp
    End If
End Sub

'To paint from the specified DC to the created DC (Others -> Me)
'Args:   FromHdc: Specific a DC handle
'        FromX, FromY: X, Y position of the source, respectively
'        ToX, ToY: X, Y position of the created DC, respectively
'        iWidth, iHeight: Width and height of the bitmap, respectively
'        DrawMode: Painting mode, default = vbSrcCopy
'Return: True if succeed, False if failed
Public Function BitBltFrom(FromHdc As Long, FromX As Long, FromY As Long, _
    ToX As Long, ToY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltFrom = BitBlt(hhDC, ToX, ToY, iWidth, iHeight, FromHdc, FromX, FromY, DrawMode)
    Else
        BitBltFrom = False
    End If
End Function

'To paint to the specified DC (Me -> Others)
'Args:   ToHdc: Specific a DC handle
'        ToX, ToY: X, Y position of the target, respectively
'        FromX, FromY: X, Y position of the created DC, respectively
'        iWidth, iHeight: Width and height of the bitmap, respectively
'        DrawMode: Painting mode, default = vbSrcCopy
'Return: True if succeed, False if failed
Public Function BitBltTo(ToHdc As Long, ToX As Long, ToY As Long, _
    FromX As Long, FromY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltTo = BitBlt(ToHdc, ToX, ToY, iWidth, iHeight, hhDC, FromX, FromY, DrawMode)
    Else
        BitBltTo = False
    End If
End Function

'To copy the specified data to the data of created bitmap (Other data -> Me)
'Args:   FromArray: Specified data
Public Sub CopyDataFrom(FromArray() As Byte)
    'Safety precaution: Check if Bitmap size is smaller than the array size
    If UBound(FromArray) + 1 < bi.bmiHeader.biSize Then
        CopyMemory ByVal lpData, FromArray(0), ByVal UBound(FromArray) + 1
    Else
        CopyMemory ByVal lpData, FromArray(0), ByVal bi.bmiHeader.biSizeImage
    End If
End Sub

'To copy the data of created bitmap to the specified data region (Me -> Other data)
'Args:   ToArray: Specified data region. Note: The array must be large enough to contain the bitmap data
'Return: True if succeed, False if failed
Public Function CopyDataTo(ToArray() As Byte) As Boolean
    'Safety precaution: Check if the array is large enough to contain the bitmap data
    If UBound(ToArray) + 1 < bi.bmiHeader.biSizeImage Then
        CopyDataTo = False
        Exit Function
    End If
    
    CopyMemory ToArray(0), ByVal lpData, ByVal bi.bmiHeader.biSizeImage
    CopyDataTo = True
End Function

'Release Bitmap and DC when the class is terminating
Private Sub Class_Terminate()
    Call DeleteMemDC
End Sub
