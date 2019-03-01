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

'��λͼ��DC
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'�������豸���ݵ�DC
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'������DC���ݵ�λͼ
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, _
    ByVal handle As Long, ByVal dw As Long) As Long
'ɾ��DC
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'ɾ��Bmp
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'��ͼ
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'�����ڴ�
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

'λͼ��Ϣͷ
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

'RGB��ɫ��
Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

'λͼ��Ϣ
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Const DIB_RGB_COLORS = 0                        '��ɫ��
Private Const BITBLT_TRANSPARENT_WINDOWS = &H40000000   '��ȡ͸������

Private bi          As BITMAPINFO                       'λͼ��Ϣ
Private hhDC        As Long                             '�ڴ�DC���
Private hhBmp       As Long                             '�ڴ�Bmp���
Private lpData      As Long                             'ָ��洢λͼ���ݵ�ָ��
Private bSize       As Long                             'λͼռ���ڴ��С���ֽڣ�

'��ȡλͼ�Ŀ���
Public Property Get iWidth() As Long
    iWidth = bi.bmiHeader.biWidth
End Property

'��ȡλͼ�ĸ߶�
Public Property Get iHeight() As Long
    iHeight = bi.bmiHeader.biHeight
End Property

'��ȡ��ǰ���õ���ɫλ��
Public Property Get iBitCount() As Integer
    iBitCount = bi.bmiHeader.biBitCount
End Property

'��ȡ��ǰλͼ�ڴ�ռ�ô�С���ֽڣ�
Public Property Get iImageSize() As Long
    iImageSize = bi.bmiHeader.biSizeImage
End Property

'��ȡ��ǰ��DC���
Public Property Get hDC() As Long
    hDC = hhDC
End Property

'��ȡ��ǰ��λͼ���
Public Property Get hBmp() As Long
    hBmp = hhBmp
End Property

'��ȡָ��λͼ���ݵĵ�ַ
Public Property Get lpBitData() As Long
    lpBitData = lpData
End Property

'����:   �����ڴ�DC
'����:   Width, Height: �ֱ����ڴ�DC�Ŀ��Ⱥ͸߶ȣ����أ���
'        BitCount: ��ɫλ��������Ϊ0, 1, 4, 8, 16, 24, 32������0Ϊjpg��png��ʽ����ɫλ����
'        hDCfrom: ��ָ���ľ���������ݵ�DC��Ĭ��Ϊ0
'
'����ֵ: ����DC�Ƿ�ɹ�
Public Function CreateMemDC(ByVal iWidth As Long, ByVal iHeight As Long, _
    Optional ByVal iBitCount As Integer = 16, Optional ByVal FromHdc As Long = 0) As Boolean
    
    '���֮ǰ�������ڴ�ͼ ��ɾ����֮ǰ���ڴ�ͼ��
    If hhDC <> 0 Or hhBmp <> 0 Then
        Call DeleteMemDC
    End If
    
    '����λͼ��Ϣ
    With bi.bmiHeader
        .biBitCount = iBitCount
        .biWidth = iWidth
        .biHeight = iHeight
        .biSize = Len(bi)
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * .biBitCount / 8
        bSize = .biSizeImage
    End With
    
    '�����ڴ�DC
    hhDC = CreateCompatibleDC(FromHdc)
    
    '�����ڴ�λͼ
    hhBmp = CreateDIBSection(hhDC, bi, DIB_RGB_COLORS, ByVal VarPtr(lpData), 0, 0)
    
    '��Bmp��DC
    SelectObject hhDC, hhBmp
    
    CreateMemDC = (hhBmp <> 0)
End Function

'����:   ɾ���������ڴ�DC��Bmp
Public Sub DeleteMemDC()
    If hhDC <> 0 Then
        DeleteDC hhDC
    End If
    If hhBmp <> 0 Then
        DeleteObject hhBmp
    End If
End Sub

'����:   ��ָ����DC��ͼ���������ڴ�DC
'����:   FromHdc: ָ����DC
'        FromX, FromY: �ֱ���ԭͼ�ϵ�X, Y����
'        ToX, ToY: �ֱ����ڴ�ͼ�ϵ�X, Y����
'        iWidth, iHeight: �ֱ���λͼ�Ŀ��Ⱥ͸߶�
'        DrawMode: ��ͼģʽ��Ĭ��ΪvbSrcCopy
'����ֵ: �����Ƿ�ɹ�
Public Function BitBltFrom(FromHdc As Long, FromX As Long, FromY As Long, _
    ToX As Long, ToY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltFrom = BitBlt(hhDC, ToX, ToY, iWidth, iHeight, FromHdc, FromX, FromY, DrawMode)
    Else
        BitBltFrom = False
    End If
End Function

'����:   ���ڴ�ͼ��ͼ��ָ����DC
'����:   ToHdc: ָ����DC
'        ToX, ToY: �ֱ���ָ��DC�ϵ�X, Y����
'        FromX, FromY: �ֱ����ڴ�ͼ�ϵ�X, Y����
'        iWidth, iHeight: �ֱ���λͼ�Ŀ��Ⱥ͸߶�
'        DrawMode: ��ͼģʽ��Ĭ��ΪvbSrcCopy
'����ֵ: �����Ƿ�ɹ�
Public Function BitBltTo(ToHdc As Long, ToX As Long, ToY As Long, _
    FromX As Long, FromY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
    
    If hhDC <> 0 And hhBmp <> 0 Then
        BitBltTo = BitBlt(ToHdc, ToX, ToY, iWidth, iHeight, hhDC, FromX, FromY, DrawMode)
    Else
        BitBltTo = False
    End If
End Function

'����:   ��ָ�������鸴�����ݵ�λͼ���ڴ���
'����:   FromArray: ָ��������
Public Sub CopyDataFrom(FromArray() As Byte)
    '�������ƣ��ڸ����ڴ�ǰ���ж������С�Ƿ��Ѿ�����λͼռ���ڴ�ɽ��ܵĴ�С
    If UBound(FromArray) + 1 < bi.bmiHeader.biSize Then
        CopyMemory ByVal lpData, FromArray(0), ByVal UBound(FromArray) + 1
    Else
        CopyMemory ByVal lpData, FromArray(0), ByVal bi.bmiHeader.biSizeImage
    End If
End Sub

'����:   ��λͼ���ڴ����ݸ��Ƶ�ָ����������
'����:   ToArray: ָ�������顣ע�⣬����Ĵ�С��������װ��λͼ�����ݡ�
'����ֵ: �����Ƿ�ɹ�
Public Function CopyDataTo(ToArray() As Byte) As Boolean
    '�������ƣ��ڸ����ڴ�ǰ�ȼ�������Ƿ�������
    If UBound(ToArray) + 1 < bi.bmiHeader.biSizeImage Then
        CopyDataTo = False
        Exit Function
    End If
    
    CopyMemory ToArray(0), ByVal lpData, ByVal bi.bmiHeader.biSizeImage
    CopyDataTo = True
End Function

'����:   ����ɾ��ǰ�ͷ�DC��λͼ
Private Sub Class_Terminate()
    Call DeleteMemDC
End Sub