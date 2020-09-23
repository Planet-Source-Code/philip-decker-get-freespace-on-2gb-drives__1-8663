<div align="center">

## Get Freespace on \> 2GB drives


</div>

### Description

This code will return the freespace on a drive even if it exceeds 2GB.
 
### More Info
 
Drive ie. "C:"

I saw lots of calculations for drive sizes but this one works for me and does not require calculations at all. It will return what Windows shows under properties on a drive. It even shows mapped network drives properly. Hope this code helps someone.

Drive size.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Philip Decker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/philip-decker.md)
**Level**          |Intermediate
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/philip-decker-get-freespace-on-2gb-drives__1-8663/archive/master.zip)

### API Declarations

```
Dim FB, BT, FBT As Currency
Dim DriveSize As String
Const Gigabyte = 1073741824
Const Megabyte = 1048576
Dim retval As Long
Private Declare Function GetDiskFreeSpace_FAT32 _
Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
(ByVal lpRootPathName As String, _
FreeBytesToCaller As Currency, BytesTotal _
As Currency, FreeBytesTotal As Currency) _
As Long
```


### Source Code

```
Public Function GetDriveInfo(DriveName As String) As String
  retval = GetDiskFreeSpace_FAT32(Left(DriveName, 2), FB, BT, FBT)
FBT = FBT * 10000 'convert result to actual size in bytes
  If FBT / Gigabyte < 1 Then 'If less than 1GB then show as MB
    DriveSize = Format(FBT / Megabyte, "####,###,###") & " MB free"
  Else 'Show as GB
    DriveSize = Format(FBT / Gigabyte, "####,###,###.00") & " GB free"
  End If
    GetDriveInfo = "[" & DriveSize & "]"
End Function
```

