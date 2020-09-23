<div align="center">

## Credit Card Checksum Checker


</div>

### Description

Checks to see if a Credit Card Number is valid by performing the LUHN-10 check on it.
 
### More Info
 
CCNum as String

True if Valid, False if Invalid

May cause skin irritation


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Anderson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-anderson.md)
**Level**          |Unknown
**User Rating**    |5.9 (625 globes from 106 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-anderson-credit-card-checksum-checker__1-902/archive/master.zip)





### Source Code

```
Public Function IsValidCCNum(CCNum As String) As Boolean
  Dim i As Integer
  Dim total As Integer
  Dim TempMultiplier As String
  For i = Len(CCNum) To 2 Step -2
    total = total + CInt(Mid$(CCNum, i, 1))
    TempMultiplier = CStr((Mid$(CCNum, i - 1, 1)) * 2)
    total = total + CInt(Left$(TempMultiplier, 1))
    If Len(TempMultiplier) > 1 Then total = total + CInt(Right$(TempMultiplier, 1))
  Next
  If Len(CCNum) Mod 2 = 1 Then total = total + CInt(Left$(CCNum, 1))
  If total Mod 10 = 0 Then
    IsValidCCNum = True
  Else
    IsValidCCNum = False
  End If
End Function
```

