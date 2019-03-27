# VB6 MemoryDC
 [English Version](README_EN.md)

一个VB6编写的内存图类，让你可以更简单的操作内存图。

# 使用方法

## 初始化内存图类

声明一个 MemoryDC 类型的变量，只需要
```VBS
Dim memDC As New clsMemDC
```

接着，在使用内存图前创建它:
```VBS
memDC.CreateMemDC image_width，image_height
```
对于函数 `CreateMemDC`，请参考“函数说明”部分中对于它的描述。

## 删除内存图类

当你不再需要内存图类的时候，只需要
```VBS
memDC.DeleteMemDC
Set memDC = Nothing
```

## 从其它的DC绘图过来

从其他的DC绘制图像过来，只需要
```VBS
memDC.BitBltFrom YourDC, FromX, FromY, ToX, ToY, image_width, image_height
```

对于函数 `BitBltFrom`，请参考“函数说明”部分中对于它的描述。

## 绘图到其他的DC

从内存图绘制图像到其他DC，只需要
```VBS
memDC.BitBltTo YourDC, ToX, ToY, FromX, FromY, image_width, image_height
```

对于函数 `BitBltTo`，请参考“函数说明”部分中对于它的描述。

## 复制字节数据到Byte数组

把内存图的数据复制到一个Byte数组里，只需要
```VBS
memDC.CopyDataTo data     'data 是一个 Byte() 类型的数组变量
```

对于函数 `CopyDataTo`，请参考“函数说明”部分中对于它的描述。

## 从Byte数组复制字节数据到内存图像

把Byte数组中的数据复制到内存图像，只需要
```VBS
memDC.CopyDataFrom data     'data 是一个 Byte() 类型的数组变量
```

对于函数 `CopyDataFrom`，请参考“函数说明”部分中对于它的描述。

## 其他功能请参考下面的部分。

# 函数说明

## `CreateMemDC` 函数

`CreateMemDC` 函数先设置内存图像的属性（宽度，高度，颜色位数），再创建它。

### 定义

```VBS
Public Function CreateMemDC(ByVal iWidth As Long, ByVal iHeight As Long, _
    Optional ByVal iBitCount As Integer = 16, Optional ByVal FromHdc As Long = 0) As Boolean
```

### 参数

`iWidth`: Long, 需要创建的内存图像的宽度。

`iHeight`: Long, 需要创建的内存图像的高度。

`iBitCount`: 可选的, Integer, 需要创建的内存图像的颜色位数。默认是16位。

`FromHdc`: 可选的, Long, 源DC句柄。默认为0。

### 返回值

类型: `Boolean`

如果内存图像成功创建，则函数返回`True`。否则，函数返回`False`。

### 例子

```VBS
memDC.CreateMemDC 1920, 1080      '创建一个 1920 * 1080 的内存图像
```

```VBS
memDC.CreateMemDC 1920, 1080, 8   '创建一个 1920 * 1080，8位 的内存图像
```

```VBS
memDC.CreateMemDC 1920, 1080, , frmMain.hDC   '从frmMain.hDC 创建一个 1920 * 1080 的内存图像
```

**注意：当您使用 `CreateMemDC` 函数的时候，这个函数会首先自动删除掉之前创建的内存图像**

## `DeleteMemDC` 函数

`DeleteMemDC` 函数可以删掉创建的内存图像。

### 定义

```VBS
Public Sub DeleteMemDC()
```

### 例子

```VBS
memDC.DeleteMemDC
```

**注意：当您不再需要内存图像的时候应调用这个函数。当类被销毁的时候，这个函数会自动调用。**

## `BitBltFrom` 函数

`BitBltFrom` 函数可以从其他DC绘制图像到内存图里。

### 定义

```VBS
Public Function BitBltFrom(FromHdc As Long, FromX As Long, FromY As Long, _
    ToX As Long, ToY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
```
### 参数

`FromHdc`: Long, 指定来源图像的DC句柄。

`FromX`: Long, 原图像的X位置。

`FromY`: Long, 原图像的Y位置。

`ToX`: Long, 目标图像的X位置。

`ToY`: Long, 目标图像的Y位置。

`iWidth`: Long, 需要绘制的图像的宽度。

`iHeight`: Long, 需要绘制的图像的高度。

`DrawMode`: 可选的, RasterOpConstants, 指定绘图的模式。 默认是 `vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS`。 绘图模式可以使以下常数的组合:

| RasterOpConstants |
|-------------------|
| vbDstInvert       |
| vbMergeCopy       |
| vbMergePaint      |
| vbNotSrcCopy      |
| vbNotSrcErase     |
| vbPatCopy         |
| vbPatInvert       |
| vbPatPaint        |
| vbSrcAnd          |
| vbSrcCopy         |
| vbSrcErase        |
| vbSrcInvert       |
| vbSrcPaint        |

### 返回值

类型: `Boolean`

如果图像成功绘制，则函数返回`True`。否则，函数返回`False`。

### 例子

```VBS
memDC.BitBltFrom frmMain.hDC, 0, 0, 0, 0, 100, 100      '从 frmMain.hDC 句柄, 窗体的(0, 0)绘制到内存图像的(0, 0)，大小为 100 * 100
```

```VBS
memDC.BitBltFrom GetDC(0), 100, 200, 150, 250, 300, 400, vbSrcInvert      '从屏幕DC句柄，屏幕的(100, 200)绘制到内存图像的(150, 250)，大小为 300 * 400，使用 vbSrcInvert 绘图模式
```

## `BitBltTo` 函数

`BitBltTo` 函数可以从内存图绘制图像到其他DC里。

### 定义

```VBS
Public Function BitBltTo(ToHdc As Long, ToX As Long, ToY As Long, _
    FromX As Long, FromY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
```
### 参数

`ToHdc`: Long, 指定绘画图像的目标DC句柄。

`ToX`: Long, 目标图像的X位置。

`ToY`: Long, 目标图像的Y位置。

`FromX`: Long, 原图像的X位置。

`FromY`: Long, 原图像的Y位置。

`iWidth`: Long, 需要绘制的图像的宽度。

`iHeight`: Long, 需要绘制的图像的高度。

`DrawMode`: 可选的，RasterOpConstants，指定绘图的模式。 默认是 `vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS`。对于DrawMode可用的常数值，请参考"`BitBltFrom` 函数"部分。

### 返回值

类型: `Boolean`

如果图像成功绘制，则函数返回`True`。否则，函数返回`False`。

### Examples

```VBS
memDC.BitBltTo frmMain.hDC, 0, 0, 0, 0, 100, 100      '绘制图像到 frmMain.hDC 句柄，从内存图像的(0, 0)绘制到窗体的(0, 0)，大小为 100 * 100
```

```VBS
memDC.BitBltTo GetDC(0), 100, 200, 150, 250, 300, 400, vbSrcInvert      '绘制图像到屏幕DC句柄，从内存图像的(100, 200)绘制到屏幕的(150, 250) ，大小为 300 * 400, 使用 vbSrcInvert 绘图模式
```

## `CopyDataFrom` 函数

`CopyDataFrom` 函数从Byte数组复制数据到内存图像。

### 定义

```VBS
Public Sub CopyDataFrom(FromArray() As Byte)
```

### 参数

`FromArray`: Byte(), 指定一个Byte数组作为复制数据的来源。

### Examples

```VBS
memDC.CopyDataFrom data         '从 data 复制图像数据，其中 data 是一个 Byte 数组
```

**注意: `CopyDataFrom` 复制整个数组的数据到内存图像的数据区里。因此，它会检查数组的大小是否大于内存图像的数据区大小。如果内存图像的数据区比数组的大小要小，那么这个函数只会复制等同于数据区大小的数据。例如，如果数组大小是10字节，内存图像数据区大小是5字节，那么这个函数只会从数组复制5字节的数据。**

## `CopyDataTo` 函数

`CopyDataTo` 函数从内存图像的数据区复制数据到Byte数组里。

### 定义

```VBS
Public Function CopyDataTo(ToArray() As Byte) As Boolean
```

### 参数

`ToArray`: Byte(), 指定一个数组作为复制数据的目标。

### 返回值

类型: `Boolean`

如果数据成功复制，那么函数返回 `True`。否则，函数返回`False`。

### 例子

```VBS
memDC.CopyDataTo data         '复制图像数据到 data 里，其中 data 是一个 Byte 数组
```

**注意: `CopyDataTo` 函数会复制所有内存图像的数据区里的数据到数组中。因此，它会检查数组大小是否足够存储内存图像的数据。如果内存图像的数据区大小比数组大，这个函数会失败。例如，如果数组的大小是5字节，内存图像的数据区大小是10字节，这个函数会返回`False`**

# 属性说明

`iWidth`: Long, 内存图像的宽度。

`iHeight`: Long, 内存图像的高度。

`iBitCount`: Long, 内存图像的颜色位数。

`iImageSize`: Long, 内存图像的数据区大小。

`hDC`: Long, 内存图像的DC句柄。

`hBmp`: Long, 内存图像的位图句柄。该位图与内存DC一同创建。

`lpBitData`: Long, 内存图像的数据区地址。

# 开源协议

MIT
