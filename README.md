<div align="center">

## Dotless IP Address


</div>

### Description

This will convert an existing IP address into an IP address without any periods (dotless IP). Although the dotless IP code is useless, the CountWords and GetWord functions are very useful. CountWords will count the number of words in a string and GetWord will get that specified word. You can also choose what character seperates the words, such as a space or period. GetWord was created by James Lewis with some modification.
 
### More Info
 
IP_Dotless#(

Byval ipAddress$) 'IP address to convert

CountWords&(ByVal inWord$,_ 'Word to check

ByVal inSep$) 'Seperation chracter

GetWord$(ByVal inWord$,_ 'Word to check

ByVal inCount&,_ 'Position where word is

ByVal inSep$) 'Seperation chracter

A quick example to copy the ip address into the clipboard would look somewhat like this:

Sub Form_Load()

Clipboard.Clear

Clipboard.SetText _

Trim$(Str$(IP_Dotless#("216.46.226.13")))

End

End Sub

IP_Dotless# returns ip as double

CountWords& returns number of words as long

GetWord$ returns specified word as string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Trunks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/trunks.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/trunks-dotless-ip-address__1-6265/archive/master.zip)





### Source Code

```
' ----- for vb6 users -----
Function IP_Dotless#(ByVal ipAddress As String)
  Dim numArray As Variant
  numArray = Split(ipAddress$, ".")
  IP_Dotless = (numArray(0) * 256 ^ 3) + _
         (numArray(1) * 256 ^ 2) + _
         (numArray(2) * 256 ^ 1) + _
         numArray(3)
End Function
' ----- for vb5 and below users -----
Function IP_Dotless# (ByVal ipAddress As String)
IP_Dotless = (Val(GetWord$(ipAddress, 1, ".")) * 256 ^ 3) + (Val(GetWord$(ipAddress, 2, ".")) * 256 ^ 2) + (Val(GetWord$(ipAddress, 3, ".")) * 256 ^ 1) + (Val(GetWord$(ipAddress, 4, ".")))
End Function
Function CountWords& (ByVal inWord$, ByVal inSep$)
Dim strTempA$
Dim strTempB$
Dim lngTempA&
Dim lngTempB&
Dim lngRet&
On Error Resume Next
inWord$ = inWord$ + inSep$
For lngRet& = 1 To Len(inWord$)
strTempA$ = Mid$(inWord$, lngRet&, Len(inSep$))
strTempB$ = strTempB$ + strTempA$
If strTempA$ = inSep$ Then
lngTempA& = Len(strTempB$) - Len(inSep$)
strTempB$ = Left$(strTempB$, lngTempA&)
lngTempB& = lngTempB& + 1
strTempB$ = ""
End If
Next lngRet&
CountWords& = lngTempB&
End Function
Function GetWord$ (ByVal inWord$, ByVal inCount&, ByVal inSep$)
Dim strTempA$
Dim strTempB$
Dim lngTempA&
Dim lngTempB&
Dim lngRet&
On Error Resume Next
inWord$ = inWord$ + inSep$
For lngRet& = 1 To Len(inWord$)
strTempA$ = Mid$(inWord$, lngRet&, Len(inSep$))
strTempB$ = strTempB$ + strTempA$
If strTempA$ = inSep$ Then
lngTempA& = Len(strTempB$) - 1
strTempB$ = Left$(strTempB$, lngTempA&)
lngTempB& = lngTempB& + 1
If lngTempB& = inCount& Then
GetWord$ = strTempB$
Exit Function
End If
strTempB$ = ""
End If
Next lngRet&
GetWord$ = ""
End Function
```

