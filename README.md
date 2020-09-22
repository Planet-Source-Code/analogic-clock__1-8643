<div align="center">

## Analogic Clock


</div>

### Description

An example for beginners, showing how to use the MoveToEx and LineTo API calls. Gets the current time, converts it into (x,y) coordinates and draw the "clock"...DON'T FORGET: DON'T EVEN THINK ABOUT VOTING FOR ME!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-06-05 16:19:18
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD6447652000\.zip](https://github.com/Planet-Source-Code/analogic-clock__1-8643/archive/master.zip)

### API Declarations

```
Public Declare Function MoveToEx Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" Alias "LineTo" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
```





