# VB6 MemoryDC
A memory DC class written in VB6. You can use this class to create and operate memory DCs easily.

# Usage

## Initalizing MemoryDC class

To declare a MemoryDC typed variable, simply
```VBS
Dim memDC As New clsMemDC
```

Next, you should create the DC when you need it:
```VBS
memDC.CreateMemDC image_width, image_height
```
For function `CreateMemDC`, please refer its description in "Functions Description" section.

## Deleting MemoryDC class

When you no longer need the MemoryDC typed variable, simply
```VBS
memDC.DeleteMemDC
Set memDC = Nothing
```

## Painting from other DCs

To paint the image from other DCs, simply
```VBS
memDC.BitBltFrom YourDC, FromX, FromY, ToX, ToY, image_width, image_height
```

For function `BitBltFrom`, please refer its description in "Functions Description" section.

## Painting to other DCs

To paint the image to other DCs, simply
```VBS
memDC.BitBltTo YourDC, ToX, ToY, FromX, FromY, image_width, image_height
```

For function `BitBltTo`, please refer its description in "Functions Description" section.

## Copying bit data to a Byte array

To copy the bit data of the image in the DC to a Byte array, simply
```VBS
memDC.CopyDataTo data     'data is a Byte() typed array variable
```

For function `CopyDataTo`, please refer its description in "Functions Description" section.

## Copying bit data from a Byte array

To copy the bit data of the image in the DC from a Byte array, simply
```VBS
memDC.CopyDataFrom data     'data is a Byte() typed array variable
```

For function `CopyDataFrom`, please refer its description in "Functions Description" section.

## For more functionalities, please read the following sections.

# Functions Description

## `CreateMemDC` function

`CreateMemDC` function sets the memory DC information (width, height, bit count) and creates it.

### Definition

```VBS
Public Function CreateMemDC(ByVal iWidth As Long, ByVal iHeight As Long, _
    Optional ByVal iBitCount As Integer = 16, Optional ByVal FromHdc As Long = 0) As Boolean
```

### Parameters

`iWidth`: Long, the width of the memory DC being created.

`iHeight`: Long, the height of the memory DC being created.

`iBitCount`: Optional, Integer, the color bit count of the memory DC being created. Default is 16 bit.

`FromHdc`: Optional, Long, the source DC handle.

### Return value

Type: `Boolean`

If the memory DC is created successfully, the function returns `True`. Otherwise, the function returns `False`.

### Examples

```VBS
memDC.CreateMemDC 1920, 1080      'Creates a 1920 * 1080 memory DC
```

```VBS
memDC.CreateMemDC 1920, 1080, 8   'Creates a 1920 * 1080, 8 bit memory DC
```

```VBS
memDC.CreateMemDC 1920, 1080, , frmMain.hDC   'Creates a 1920 * 1080 memory DC from frmMain's hDC
```

**NOTE: When you use `CreateMemDC` function, this function deletes the previous memory DC automatically.**

## `DeleteMemDC` function

`DeleteMemDC` deletes the create memory DC.

### Definition

```VBS
Public Sub DeleteMemDC()
```

### Examples

```VBS
memDC.DeleteMemDC
```

**NOTE: Call `DeleteMemDC` when you don't need the memory DC anymore. This function is automatically called when the class is being terminated.**

## `BitBltFrom` function

`BitBltFrom` paints image from other DCs to the created memory DC.

### Definition

```VBS
Public Function BitBltFrom(FromHdc As Long, FromX As Long, FromY As Long, _
    ToX As Long, ToY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
```
### Parameters

`FromHdc`: Long, Specifics the DC handle to paint image from.

`FromX`: Long, X position of the original image.

`FromY`: Long, Y position of the original image.

`ToX`: Long, X position of the target image.

`ToY`: Long, Y position of the target image.

`iWidth`: Long, width of the image being painted

`iHeight`: Long, height of the image being painted

`DrawMode`: Optional, RasterOpConstants, specifics painting mode. Default is `vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS`. Painting mode can be the combination of the following constants:

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

### Return value

Type: `Boolean`

If the image is painted successfully, the function returns `True`. Otherwise, the function returns `False`.

### Examples

```VBS
memDC.BitBltFrom frmMain.hDC, 0, 0, 0, 0, 100, 100      'Paints the image from frmMain.hDC, from (0, 0) of the window to (0, 0) of the memory DC, sized 100 * 100
```

```VBS
memDC.BitBltFrom GetDC(0), 100, 200, 150, 250, 300, 400, vbSrcInvert      'Paints the image from the screen DC, from (100, 200) of the screen to (150, 250) of the memory DC, sized 300 * 400, using vbSrcInvert painting mode
```

## `BitBltTo` function

`BitBltTo` paints image from the created memory DC to other DCs.

### Definition

```VBS
Public Function BitBltTo(ToHdc As Long, ToX As Long, ToY As Long, _
    FromX As Long, FromY As Long, iWidth As Long, iHeight As Long, _
    Optional DrawMode As RasterOpConstants = vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS) As Boolean
```
### Parameters

`ToHdc`: Long, Specifics the DC handle to paint image to.

`ToX`: Long, X position of the target image.

`ToY`: Long, Y position of the target image.

`FromX`: Long, X position of the original image.

`FromY`: Long, Y position of the original image.

`iWidth`: Long, width of the image being painted

`iHeight`: Long, height of the image being painted

`DrawMode`: Optional, RasterOpConstants, specifics painting mode. Default is `vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS`. For constants available for DrawMode, please refer to "`BitBltFrom` function" section.

### Return value

Type: `Boolean`

If the image is painted successfully, the function returns `True`. Otherwise, the function returns `False`.

### Examples

```VBS
memDC.BitBltTo frmMain.hDC, 0, 0, 0, 0, 100, 100      'Paints the image to frmMain.hDC, from (0, 0) of the memory DC to (0, 0) of the window, sized 100 * 100
```

```VBS
memDC.BitBltTo GetDC(0), 100, 200, 150, 250, 300, 400, vbSrcInvert      'Paints the image to the screen DC, from (100, 200) of the memory DC to (150, 250) of the screen DC, sized 300 * 400, using vbSrcInvert painting mode
```

## `CopyDataFrom` function

`CopyDataFrom` function copies data from a Byte array to the memory DC.

### Definition

```VBS
Public Sub CopyDataFrom(FromArray() As Byte)
```

### Parameters

`FromArray`: Byte(), the Byte array to copy data from.

### Examples

```VBS
memDC.CopyDataFrom data         'Copy image data from data, where data is a Byte array
```

**NOTE: `CopyDataFrom` function copies all data from the array to the memory region of the memory DC. So it checks if the array size is larger than the size of memory region of the memory DC or not. If the memory region is smaller than the size of the array, it only copies data in the same size to the size of memory region of the memory DC. For example, if the array is 10 bytes, and the size of the memory region of the memory DC is 5 bytes, this function will only copy 5 bytes from the array.**

## `CopyDataTo` function

`CopyDataFrom` function copies data from the memory DC to a Byte array.

### Definition

```VBS
Public Function CopyDataTo(ToArray() As Byte) As Boolean
```

### Parameters

`ToArray`: Byte(), the Byte array to copy data to.

### Return value

Type: `Boolean`

If the data is copied successfully, the function returns `True`. Otherwise, the function returns `False`.

### Examples

```VBS
memDC.CopyDataTo data         'Copy image data to data, where data is a Byte array
```

**NOTE: `CopyDataTo` function copies all data from the memory region of the memory DC to the array. So it checks if the array size is large enough to receive the memory DC data or not. If the memory region is larger than the size of the array, this function fails. For example, if the array is 5 bytes, and the size of the memory region of the memory DC is 10 bytes, this function will return `False`**

# Properties Description

`iWidth`: Long, the width of the memory DC.

`iHeight`: Long, the height of the memory DC.

`iBitCount`: Long, the color bit count of the memory DC.

`iImageSize`: Long, the size of the memory region of the memory DC.

`hDC`: Long, the handle to the created memory DC.

`hBmp`: Long, the handle to the bitmap created. This bitmap is created together with the memory DC.

`lpBitData`: Long, the address of the memory region of the memory DC.

# License

MIT
