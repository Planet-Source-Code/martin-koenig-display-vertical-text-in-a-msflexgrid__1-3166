<div align="center">

## Display vertical Text in a MSFlexGrid


</div>

### Description

With this code you are able to display vertical Text in a Flexgrid. This is

very helpfull, when you need to display a column which only has a YES/NO

Value and would waste to much horizontal Column Space to display the Column

Header
 
### More Info
 
1.Start a simple Project

2.Add the Microsoft Flex Grid component to the Project

3.Add a Picture to the Project, set the Index=0, Visible=FALSE and

AutoRedraw=TRUE

4.Copy the API Code and but it in a Module

5.Copy the normale Code in to the Form Code Module

I am not aware of any Side Effects, I am using this code in several

Application. But be aware that you set the AutoRedraw Property of the

PictureBox TRUE, otherwise you wont see anything. Be also carefull with the

Picture Alignement settings of the certain

MSFlexGrid1.CellPictureAlignement.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Martin Koenig](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/martin-koenig.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/martin-koenig-display-vertical-text-in-a-msflexgrid__1-3166/archive/master.zip)

### API Declarations

```
Option Explicit
'****************
' API Declaration
'****************
Public Declare Function CreateFontIndirect Lib "gdi32" Alias
"CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal
hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As
Long
'**********************
' API Const Declaration
'**********************
Public Const LF_FACESIZE = 32
Public Const ANTIALIASED_QUALITY = 4
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const DEFAULT_CHARSET = 1
Public Const OUT_TT_PRECIS = 4
Public Const VARIABLE_PITCH = 2
'*********************
' API Type Declaration
'*********************
Public Type LOGFONT
 lfHeight As Long
 lfWidth As Long
 lfEscapement As Long
 lfOrientation As Long
 lfWeight As Long
 lfItalic As Byte
 lfUnderline As Byte
 lfStrikeOut As Byte
 lfCharSet As Byte
 lfOutPrecision As Byte
 lfClipPrecision As Byte
 lfQuality As Byte
 lfPitchAndFamily As Byte
 lfFaceName As String * LF_FACESIZE
End Type
```


### Source Code

```
Private Sub Form_Load()
'---------------------------------------------------------------------------
---------------
' Name  : Form_Load
' Purpose  : Event when Form is being loaded
' Parameters :
' Date  : Sonntag 22 August 1999 17:36
' Revised  :
'---------------------------------------------------------------------------
---------------
 'Draw the Text
 DrawText
End Sub
Private Sub DrawText()
'---------------------------------------------------------------------------
---------------
' Name  : DrawText
' Purpose  : This Function Draws the Text vertical
' Parameters :
' Date  : Sonntag 22 August 1999 17:36
' Revised  :
'---------------------------------------------------------------------------
---------------
 'Declaration
 Dim stText1 As String
 Dim stText2 As String
 Dim imaxWidth As Integer
 Dim picTmp As PictureBox
 'Define the Text, add some extra spaces before and after the Text
 stText1 = " This is my vertical Text "
 stText2 = " This is shorter "
 'Get the max Width of the Text which will be displayed
 If TextWidth(stText1) > imaxWidth Then imaxWidth = TextWidth(stText1)
 If TextWidth(stText2) > imaxWidth Then imaxWidth = TextWidth(stText2)
 'Start with
 With MSFlexGrid1
 'Set the Width of the Col's so that the Text will be
 'Displayed ok
 .ColWidth(0) = TextHeight("W") * 2
 .ColWidth(1) = TextHeight("W") * 2
 'Set Hight of the First Row, thats where we are going to display
 'the vertical Text
 .RowHeight(0) = imaxWidth
 'Set Row for the First Time
 .Row = 0
 'Save Rotated Text
 Set picTmp = GetRotatetText(stText1)
 'Set Col
 .Col = 0
 'Set Picture
 Set .CellPicture = picTmp.Image
 'Save Rotated Text
 Set picTmp = GetRotatetText(stText2)
 'Set Col
 .Col = 1
 'Set Picture
 Set .CellPicture = picTmp.Image
 'End with
 End With
End Sub
Public Function GetRotatetText(stText As String) As PictureBox
'---------------------------------------------------------------------------
---------------
' Name  : GetRotatetText
' Purpose  : This Function Returns the Picture, which contains the
verical drawed Text
' Parameters : stText Contains the Text which has to be draw
' Date  : Sonntag 22 August 1999 17:37
' Revised  :
'---------------------------------------------------------------------------
---------------
 'Declaration
 Dim iIndex As Integer
 'Check if the first Picture has been used allready
 If Picture1(0).Tag <> "" Then
 Load Picture1(Picture1.Count)
 Else
 Picture1(0).Tag = "used"
 End If
 'Save Index
 iIndex = Picture1.Count - 1
 'Start with
 With Picture1(iIndex)
 'Set the Heigth
 .Height = MSFlexGrid1.RowHeight(0)
 'Draws the Text on the PictureBox
 DrawRotatedText Picture1(iIndex), 0, .Height, 90, stText
 'Set Return
 Set GetRotatetText = Picture1(iIndex)
 'End with
 End With
End Function
Public Function DrawRotatedText(ByVal pTarget As Object, _
        ByVal X As Single, ByVal Y As Single, _
        ByVal dAngle As Double, _
        ByVal stText As String) As Boolean
'---------------------------------------------------------------------------
---------------
' Name  : DrawRotatedText
' Purpose  : This Function Draws the Text an the PictureBox which is
defined in the
'    parameters
' Parameters : pTarget An Object, in this case the PictureBox
'    X  The X Coordinate
'    Y  The Y Coordinate
'    dAngle The Angle which should be used to draw, any anlge is
possible
'    stText The Text which should be drawn on the PictureBox
' Date  : Sonntag 22 August 1999 17:38
' Revised  :
'---------------------------------------------------------------------------
---------------
 'Declaration
 Dim RotFont As LOGFONT, OldFont As Long, hFont As Long
 Dim OldX As Single, OldY As Single
 'Set Error Handling
 On Error GoTo ErrorRotatedText
 'Define the LogFont Type
 With RotFont
 .lfEscapement = CLng(dAngle * 10)
 .lfFaceName = pTarget.FontName
 .lfHeight = pTarget.FontSize * -20 / Screen.TwipsPerPixelY
 .lfWeight = IIf(pTarget.FontBold, FW_BOLD, FW_NORMAL)
 If pTarget.FontStrikethru Then .lfStrikeOut = 1
 If pTarget.FontUnderline Then .lfUnderline = 1
 If pTarget.FontItalic Then .lfItalic = 1
 .lfOutPrecision = OUT_TT_PRECIS
 .lfQuality = ANTIALIASED_QUALITY
 .lfCharSet = DEFAULT_CHARSET
 .lfPitchAndFamily = VARIABLE_PITCH
 End With
 'Generate and Asign the Font-Object
 hFont = CreateFontIndirect(RotFont)
 OldFont = SelectObject(pTarget.hDC, hFont)
 'Save the Coordinatees
 OldX = pTarget.CurrentX
 OldY = pTarget.CurrentY
 'Set the desired Coordinates
 pTarget.CurrentX = X
 pTarget.CurrentY = Y
 'Print the Text
 pTarget.Print stText
 'Set the Coordinates back
 pTarget.CurrentX = OldX
 pTarget.CurrentY = OldY
 'Set original Font back and destroy the Generated Font
 SelectObject pTarget.hDC, OldFont
 DeleteObject hFont
 'Set Return
 DrawRotatedText = True
ExitRotatedText:
 Exit Function
ErrorRotatedText:
 Resume ExitRotatedText
End Function
```

