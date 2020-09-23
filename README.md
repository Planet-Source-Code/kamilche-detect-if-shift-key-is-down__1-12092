<div align="center">

## Detect if Shift Key is down


</div>

### Description

A function that returns whether or not the shift key is currently down.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kamilche](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kamilche.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kamilche-detect-if-shift-key-is-down__1-12092/archive/master.zip)

### API Declarations

```
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal uAction As Long) As Long
```


### Source Code

```
Private Function ShiftDown()
  Dim RetVal As Long
  RetVal = GetAsyncKeyState(16) 'SHIFT key
  If (RetVal And 32768) <> 0 Then
    ShiftDown = True
  Else
    ShiftDown = False
  End If
End Function
```

