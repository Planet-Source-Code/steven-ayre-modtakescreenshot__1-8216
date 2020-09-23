<div align="center">

## modTakeScreenShot


</div>

### Description

Takes a screenshot of the screen using the Windows API. The subroutine ScreenShot takes a screenshot of the entire screen, the subroutine PartialScreenShot takes a screenshot of only part of the screen.
 
### More Info
 
Both subroutines require the DC of the destination for the image - e.g. Form1.hDC Picture1.hDC

In addition, the PartialScreenShot subroutine needs the Top, Left, Width, Height parameters which define an area of the screen to take a screenshot of.

The API call which takes the screenshot is BitBlt. This takes the image on one object and puts it on another object, using the objects' DCs to uniquely identify each object. The other two API calls are used to get the DC of the entire screen. The DC of the object that you are putting the image on can be found from its .hDC property (if it has one).

None, but puts an image on the object which has its hDC passed to one of the subroutines.

None! (afaik)


<span>             |<span>
---                |---
**Submitted On**   |2000-05-21 15:30:20
**By**             |[Steven Ayre](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-ayre.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD59635212000\.zip](https://github.com/Planet-Source-Code/steven-ayre-modtakescreenshot__1-8216/archive/master.zip)








