<div align="center">

## ElasticForms \(actually does the resizing\!\)


</div>

### Description

After seeing the "Elastic" post below, I thought I'd release my ElasticForms module 'cause this one actually *does* resize the components on the form (even lines). It's pretty tight, fast and it even allows you to set a min width and min height for a form. A zip containing the source and an example project can be found on my home page...
 
### More Info
 
Paste this code into a class module (clsElasticForms for example).

Note: It doesn't handle fonts - it easily could do, I just didn't need it to.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Leigh Bowers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/leigh-bowers.md)
**Level**          |Unknown
**User Rating**    |4.3 (166 globes from 39 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/leigh-bowers-elasticforms-actually-does-the-resizing__1-2114/archive/master.zip)





### Source Code

```
Option Explicit
Private fForm As Form
Private lOriginalWidth As Long
Private lOriginalHeight As Long
Private lMinWidth As Long
Private lMinHeight As Long
Private Type udtControl
  lLeft As Long
  lTop As Long
  lWidth As Long
  lHeight As Long
End Type
Private aControls() As udtControl
Public Property Let Form(ByVal fPassForm As Form)
Dim iCount As Integer
Dim cControl As Control
  Set fForm = fPassForm
  ' Store form's original Width & Height
  lOriginalWidth = fForm.Width
  lOriginalHeight = fForm.Height
  ' Use error trapping to ignore components that don't
  ' support certain properties being read at run-time
  On Error Resume Next
  ' Store the form's component's properties
  iCount = 0
  ReDim aControls(fForm.Controls.Count)
  For Each cControl In fForm.Controls
    iCount = iCount + 1
    With aControls(iCount)
      If TypeOf cControl Is Line Then
        .lLeft = cControl.X1
        .lTop = cControl.Y1
        .lWidth = cControl.X2
        .lHeight = cControl.Y2
      Else
        .lLeft = cControl.Left
        .lTop = cControl.Top
        .lWidth = cControl.Width
        .lHeight = cControl.Height
      End If
    End With
  Next
End Property
Public Sub FormResize()
  ' Resize the form
Dim iCount As Integer
Dim cControl As Control
Dim iTaskBarHeight As Integer
Dim sOriginalWidthUnit As Single
Dim sOriginalHeightUnit As Single
  If fForm Is Nothing Then Exit Sub
  ' Don't process minimized forms
  If fForm.WindowState = vbMinimized Then Exit Sub
  ' Check form size against minimums
  If fForm.Width < lMinWidth Then fForm.Width = lMinWidth
  If fForm.Height < lMinHeight Then fForm.Height = lMinHeight
  ' Perform calculations in advance (speed increase)
  iTaskBarHeight = 28 * Screen.TwipsPerPixelY ' Standard height
  sOriginalWidthUnit = lOriginalWidth / fForm.Width
  sOriginalHeightUnit = (lOriginalHeight - iTaskBarHeight) / (fForm.Height - iTaskBarHeight)
  ' Use error trapping to ignore components that don't
  ' support certain properties being set at run-time
  On Error Resume Next
  ' Resize...
  iCount = 0
  For Each cControl In fForm.Controls
    iCount = iCount + 1
    With cControl
      If TypeOf cControl Is Line Then
        .X1 = Int(aControls(iCount).lLeft / sOriginalWidthUnit)
        .Y1 = Int(aControls(iCount).lTop / sOriginalHeightUnit)
        .X2 = Int(aControls(iCount).lWidth / sOriginalWidthUnit)
        .Y2 = Int(aControls(iCount).lHeight / sOriginalHeightUnit)
      Else
        .Left = Int(aControls(iCount).lLeft / sOriginalWidthUnit)
        .Top = Int(aControls(iCount).lTop / sOriginalHeightUnit)
        .Width = Int(aControls(iCount).lWidth / sOriginalWidthUnit)
        .Height = Int(aControls(iCount).lHeight / sOriginalHeightUnit)
      End If
    End With
  Next
End Sub
Private Sub Class_Terminate()
  Set fForm = Nothing
End Sub
Public Property Let MinWidth(ByVal lPassMinWidth As Long)
  lMinWidth = lPassMinWidth
End Property
Public Property Let MinHeight(ByVal lPassMinheight As Long)
  lMinHeight = lPassMinheight
End Property
```

