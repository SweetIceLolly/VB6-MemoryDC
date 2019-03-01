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

# Property Description
