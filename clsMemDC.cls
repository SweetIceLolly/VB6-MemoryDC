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

'绑定位图到DC
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'创建与设备兼容的DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'创建与DC兼容的位图
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, _
    ByVal handle As Long, ByVal dw As Long) As Long
'删除DC
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'删除Bmp
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'绘图
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'复制内存
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

'位图信息头
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

'RGB颜色表
Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

'位图信息
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Const DIB_RGB_COLORS = 0                        '颜色表
Private Const BITBLT_TRANSPARENT_WINDOWS = &H40000000   '截取透明窗体

Private bi          As BITMAPINFO                       '位图信息
Private hhDC        As Long                             '内存DC句柄
Private hhBmp       As Long                             '内存Bmp句柄
Private lpData      As Long                             '指向存储位图数据的指针
Private bSize       As Long                             '位图占用内存大小（字节）

'获取位图的宽度
Public Property Get iWidth() As Long
    iWidth = bi.bmiHeader.biWidth
End Property

'获取位图的高度
Public Property Get iHeight() As Long
    iHeight = bi.bmiHeader.biHeight
End Property

'获取当前设置的颜色位数
Public Property Get iBitCount() As Integer
    iBitCount = bi.bmiHeader.biBitCount
End Property

'获取当前位图内存占用大小（字节）
Public Property Get iImageSize() As Long
    iImageSize = bi.bmiHeader.biSizeImage
End Property

'获取当前的DC句柄
Public Property Get hDC() As Long
    hDC = hhDC
End Property

'获取当前的位图句柄
Public Property Get hBmp() As Long
    hBmp = hhBmp
End Property

'获取指向位图数据的地址
Public Property Get lpBitData() As Long
    lpBitData = lpData
End Property

'描述:   创建内存DC
'参数:   Width, Height: 分别是内存DC的宽度和高度（像素）；
'        BitCount: 颜色位数，可以为0, 1, 4, 8, 16, 24, 32。其中0为jpg或png格式的颜色位数；
'        hDCfrom: 从指定的句柄创建兼容的DC，默认为0
'
'返回值: 创建DC是否成功
Public Function CreateMemDC(ByVal iWidth As Long, ByVal iHeight As Long, _
    Optional ByVal iBitCount As Integer = 16, Optional ByVal FromHdc As Long = 0) As Boolean
    
    '如果之前创建了内存图 则删除掉之前的内存图先
    If hhDC <> 0 Or hhBmp <> 0 Then
        Call DeleteMemDC
    End If
    
    '设置位图信息
    With bi.bmiHeader
        .biBitCount = iBitCount
        .biWidth = iWidth
        .biHeight = iHeight
        .biSize = Len(bi)
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * .biBitCount / 8
        bSize = .biSizeImage
    End With
    
    '创建内存DC
    hhDC = CreateCompatibleDC(FromHdc)
    
    '创建内存位图
    hhBmp = CreateDIBSection(hhDC, bi, DIB_RGB_COLORS, ByVal VarPtr(lpData), 0, 0)
    
    '绑定Bmp和DC
    SelectObject hhDC, hhBmp
    
    CreateMemDC = (hhBmp <> 0)
End Function

'描述:   删除创建的内存DC和Bmp
Public Sub DeleteMemDC()
    If hhDC <> 0 Then
        DeleteDC hhDC
    End If
    If hhBmp <> 0 Then
        DeleteObject hhBmp
    End If
End Sub

'描述:   从指定的DC绘图到创建的内存DC
'参数:   FromHdc: 指定的DC
'        FromX, FromY: 分别是原图上的X, Y坐标
'        ToX, ToY: 分别是内存图上的X, Y坐标
'        iWidth, iHeight: 分别是位图的宽度和高度
'        DrawMode: 绘图模式，默认为vbSrcCopy
'返回值: 操作是否成功
Public Function BitBltFrom(FromHdc As Long, FromX As Long, FromY As Long, _
    ToX As Long, ToY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltFrom = BitBlt(hhDC, ToX, ToY, iWidth, iHeight, FromHdc, FromX, FromY, DrawMode)
    Else
        BitBltFrom = False
    End If
End Function

'描述:   把内存图绘图到指定的DC
'参数:   ToHdc: 指定的DC
'        ToX, ToY: 分别是指定DC上的X, Y坐标
'        FromX, FromY: 分别是内存图上的X, Y坐标
'        iWidth, iHeight: 分别是位图的宽度和高度
'        DrawMode: 绘图模式，默认为vbSrcCopy
'返回值: 操作是否成功
Public Function BitBltTo(ToHdc As Long, ToX As Long, ToY As Long, _
    FromX As Long, FromY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltTo = BitBlt(ToHdc, ToX, ToY, iWidth, iHeight, hhDC, FromX, FromY, DrawMode)
    Else
        BitBltTo = False
    End If
End Function

'描述:   从指定的数组复制数据到位图的内存中
'参数:   FromArray: 指定的数组
Public Sub CopyDataFrom(FromArray() As Byte)
    '保护机制：在复制内存前先判断数组大小是否已经超出位图占用内存可接受的大小
    If UBound(FromArray) + 1 < bi.bmiHeader.biSize Then
        CopyMemory ByVal lpData, FromArray(0), ByVal UBound(FromArray) + 1
    Else
        CopyMemory ByVal lpData, FromArray(0), ByVal bi.bmiHeader.biSizeImage
    End If
End Sub

'描述:   把位图的内存数据复制到指定的数组中
'参数:   ToArray: 指定的数组。注意，数组的大小必须足以装下位图的数据。
'返回值: 操作是否成功
Public Function CopyDataTo(ToArray() As Byte) As Boolean
    '保护机制：在复制内存前先检测数组是否能容纳
    If UBound(ToArray) + 1 < bi.bmiHeader.biSizeImage Then
        CopyDataTo = False
        Exit Function
    End If
    
    CopyMemory ToArray(0), ByVal lpData, ByVal bi.bmiHeader.biSizeImage
    CopyDataTo = True
End Function

'描述:   在类删除前释放DC和位图
Private Sub Class_Terminate()
    Call DeleteMemDC
End Sub
