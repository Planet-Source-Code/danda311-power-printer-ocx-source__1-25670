VERSION 5.00
Begin VB.UserControl PPrint 
   CanGetFocus     =   0   'False
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   ForeColor       =   &H00000000&
   InvisibleAtRuntime=   -1  'True
   MaskPicture     =   "PPrint.ctx":0000
   Picture         =   "PPrint.ctx":1CCA
   ScaleHeight     =   720
   ScaleWidth      =   705
   ToolboxBitmap   =   "PPrint.ctx":3994
End
Attribute VB_Name = "PPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'special stuff
Private M_Align As Integer
Private M_Copies As Integer
Private M_Font As String
Private M_FontSize As Single
Private M_Landscape As Boolean
Private M_Italic As Boolean
Private M_Strike As Boolean
Private M_Underline As Boolean
Private M_Bold As Boolean
Private M_Rmar As Double
Private M_Lmar As Double

'constant defaults
Const M_Def_Rmar = 0.75
Const M_Def_Lmar = 0.75
Const M_Def_Align = 1
Const M_Def_Copies = 1
Const M_Def_Font = "Arial"
Const M_Def_FontSize = 12



'My stuff Not Editable by the user anywhere
Private First As Boolean
Private TempLine(500) As String
Private Line2(500) As String
Private Amount As Integer
Private Amt2 As Integer

Private Sub UserControl_Initialize()
'resize the little picture
With UserControl
.Height = 720
.Width = 720
End With

End Sub

Private Sub UserControl_InitProperties()
'When Placed on Initialize Default Values
M_Rmar = M_Def_Rmar
M_Lmar = M_Def_Lmar
M_Align = M_Def_Align
M_Copies = M_Def_Copies
M_Font = M_Def_Font
M_FontSize = M_Def_FontSize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'printing Info
Align = PropBag.ReadProperty("Align", M_Def_Align)
Copies = PropBag.ReadProperty("Copies", M_Def_Copies)
Landscape = PropBag.ReadProperty("Landscape", False)
Font = PropBag.ReadProperty("Font", M_Def_Font)
FontSize = PropBag.ReadProperty("FontSize", M_Def_FontSize)
Bold = PropBag.ReadProperty("Bold", False)
Italic = PropBag.ReadProperty("Italic", False)
Strike = PropBag.ReadProperty("Strike", False)
Underline = PropBag.ReadProperty("Underline", False)
Rmar = PropBag.ReadProperty("Rmar", M_Def_Rmar)
Lmar = PropBag.ReadProperty("Lmar", M_Def_Lmar)
End Sub

Private Sub UserControl_Resize()
'keep usercontrol samewidth and height
With UserControl
.Height = 720
.Width = 720
End With
End Sub
Public Function Start()
'This is The Starting Command for the Printer OCX
First = True
With Printer
    If M_Landscape = True Then .Orientation = 2 Else .Orientation = 1
End With
'You must do this first.  Its the law.
Printer.Print
End Function


Public Function PPrint(ByVal Line As String)
Dim PrintAbleArea As Single
Dim x As Integer
Dim SizeOf As Double
Dim StartAt As Double
Dim SizePage As Single
Dim ret As Variant
Dim y As Integer
Dim SizeText As Single

With Printer
If First = True Then
.CurrentY = 500
First = False
End If

.FontSize = M_FontSize
.Font = M_Font
.FontBold = M_Bold
.FontUnderline = M_Underline
.FontStrikethru = M_Strike
.FontItalic = M_Italic

SizeText = .TextHeight("A")
If (SizeText + .CurrentY) > 14999 Then
.NewPage
.CurrentY = 500
End If

If M_Landscape = True Then
PrintAbleArea = 15930 - ((M_Lmar * 1500) + (M_Rmar * 1500))
SizePage = 15930
Else
PrintAbleArea = 12100 - ((M_Lmar * 1500) + (M_Rmar * 1500))
SizePage = 12100
End If


Select Case M_Align

Case 1
DoTextBox (Line)
For y = 1 To Amt2
'split the line to smaller parts
ret = SplitLine(CStr(Line2(y)), PrintAbleArea)
'where to start printing
If M_Lmar < 0.21 Then
StartAt = ((0.01 * 1500))
Else
StartAt = (((M_Lmar - 0.2) * 1500))
End If
For x = 1 To Amount
''''''''Check If I should Move Down a Page''''''''
SizeText = .TextHeight("A")
If (SizeText + .CurrentY) > 14999 Then
.NewPage
.CurrentY = 500
End If
''''''''Done Checking''''''''
.CurrentX = StartAt
DoEvents
Printer.Print Trim(TempLine(x))
TempLine(x) = ""
Next x
Next y

Case 2
DoTextBox (Line)
For y = 1 To Amt2
'split the line to smaller parts
ret = SplitLine(CStr(Line2(y)), PrintAbleArea)
'where to start printing
For x = 1 To Amount
SizeOf = .TextWidth(TempLine(x))
StartAt = ((SizePage - ((M_Rmar * 1500))) - SizeOf)
''''''''Check If I should Move Down a Page''''''''
SizeText = .TextHeight("A")
If (SizeText + .CurrentY) > 14999 Then
.NewPage
.CurrentY = 500
End If
''''''''Done Checking''''''''
.CurrentX = StartAt
DoEvents
Printer.Print Trim(TempLine(x))
TempLine(x) = ""
Next x
Next y

Case 3
DoTextBox (Line)
For y = 1 To Amt2
'split the line to smaller parts
ret = SplitLine(CStr(Line2(y)), PrintAbleArea)
'where to start printing
For x = 1 To Amount
SizeOf = .TextWidth(TempLine(x))
StartAt = ((SizePage - SizeOf) / 2) - (0.1 * 1500) + ((M_Lmar - 1) * 750) - ((M_Rmar - 1) * 750)
''''''''Check If I should Move Down a Page''''''''
SizeText = .TextHeight("A")
If (SizeText + .CurrentY) > 14999 Then
.NewPage
.CurrentY = 500
End If
''''''''Done Checking''''''''
.CurrentX = StartAt
DoEvents
Printer.Print Trim(TempLine(x))
TempLine(x) = ""
Next x
Next y
End Select

End With
End Function

Private Function SplitLine(Line As String, tobig As Single)
Dim SizeOf As Double
Dim StartAt As Double
Dim todo As Integer
Dim temp As String
Dim temp1, temp2 As String
Dim x, j As Integer
Dim dalen, origlen As Integer
Dim highest As Integer
Dim retry As Boolean
Dim done As Boolean

With Printer
'~~~~~~~~~~~~~~~~~~~~~~~~~
Amount = 1
'~~~~~~~~~~~~~~~~~~~~~~~~~~



retry = True
TempLine(1) = Line
todo = 1
Do
SizeOf = .TextWidth(TempLine(todo))
    If SizeOf > tobig Then
    dalen = Len(TempLine(todo))
    origlen = dalen
        Do
            For x = dalen To 1 Step -1
            
            temp = Mid$(TempLine(todo), x, 1)
                
                If x = 1 Then
                For j = 60 To 1 Step -1
                .FontSize = CSng(j)
                SizeOf = .TextWidth(TempLine(todo))
                If SizeOf < tobig Then
                highest = j
                j = 1
                End If
                Next j
                temp = " "
                Do
                Printer.FontSize = InputBox("You Need to use a smaller font size to fit this on one line without going off the edge.  Please enter one from " & highest & " or lower", "Error Too Big")
                Loop Until Printer.FontSize <= highest
                End If
                
                If temp = " " Then
                dalen = x - 1
                temp1 = Mid$(TempLine(todo), 1, x - 1)
                temp2 = Mid$(TempLine(todo), x + 1, origlen)
                SizeOf = .TextWidth(temp1)
                    If SizeOf < tobig Then
                    retry = False
                    x = 1
                    TempLine(todo) = temp1
                    TempLine(todo + 1) = temp2
                        If temp2 = "" Then
                        done = True
                        Else
                        Amount = Amount + 1
                        End If
                    
                    Else
                    retry = True
                    End If
                End If
            Next x
        Loop Until retry = False
    Else
    done = True
    End If
todo = todo + 1
Loop Until done = True
End With

End Function
Private Function DoTextBox(ByVal DaLine As String)
Dim Last As Integer
Dim x As Integer
Last = 1
Amt2 = 0
For x = 1 To Len(DaLine)
If Mid(DaLine, x, 2) = vbCrLf Then
Amt2 = Amt2 + 1
Line2(Amt2) = Mid(DaLine, Last, x - Last)
Last = x + 2
End If
Next x

If Amt2 = 0 Then
Amt2 = 1
Line2(Amt2) = DaLine
Else
If Trim(Mid(DaLine, Last + 2, 5)) <> "" Then
Amt2 = Amt2 + 1
Line2(Amt2) = Mid(DaLine, Last, Len(DaLine) - Last + 1)
End If
End If

End Function

Public Function Finish()
Printer.EndDoc
First = True
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Align", M_Align, M_Def_Align)
Call PropBag.WriteProperty("Copies", M_Copies, M_Def_Copies)
Call PropBag.WriteProperty("Font", M_Font, M_Def_Font)
Call PropBag.WriteProperty("FontSize", M_FontSize, M_Def_FontSize)
Call PropBag.WriteProperty("Landscape", M_Landscape, False)
Call PropBag.WriteProperty("Italic", M_Italic, False)
Call PropBag.WriteProperty("Strike", M_Strike, False)
Call PropBag.WriteProperty("Underline", M_Underline, False)
Call PropBag.WriteProperty("Bold", M_Bold, False)
Call PropBag.WriteProperty("Rmar", M_Rmar, M_Def_Rmar)
Call PropBag.WriteProperty("Lmar", M_Lmar, M_Def_Lmar)
End Sub

Public Property Get Align() As Integer
  Align = M_Align
End Property

Public Property Let Align(ByVal New_Align As Integer)
  M_Align = New_Align
  PropertyChanged "Align"
End Property


Public Property Get Copies() As Integer
  Copies = M_Copies
End Property

Public Property Let Copies(ByVal New_Copies As Integer)
  M_Copies = New_Copies
  PropertyChanged "Copies"
End Property

Public Property Get FontSize() As Integer
  FontSize = M_FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Integer)
  M_FontSize = New_FontSize
  PropertyChanged "FontSize"
End Property

Public Property Get Font() As String
  Font = M_Font
End Property

Public Property Let Font(ByVal New_Font As String)
  M_Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Get Landscape() As Boolean
  Landscape = M_Landscape
End Property

Public Property Let Landscape(ByVal New_Landscape As Boolean)
  M_Landscape = New_Landscape
  PropertyChanged "Landscape"
End Property

Public Property Get Italic() As Boolean
  Italic = M_Italic
End Property

Public Property Let Italic(ByVal New_Italic As Boolean)
  M_Italic = New_Italic
  PropertyChanged "Italic"
End Property

Public Property Get Strike() As Boolean
  Strike = M_Strike
End Property

Public Property Let Strike(ByVal New_Strike As Boolean)
  M_Strike = New_Strike
  PropertyChanged "Strike"
End Property

Public Property Get Underline() As Boolean
  Underline = M_Underline
End Property

Public Property Let Underline(ByVal New_Underline As Boolean)
  M_Underline = New_Underline
  PropertyChanged "Underline"
End Property

Public Property Get Bold() As Boolean
  Bold = M_Bold
End Property

Public Property Let Bold(ByVal New_Bold As Boolean)
  M_Bold = New_Bold
  PropertyChanged "Bold"
End Property

Public Property Get Lmar() As Double
  Lmar = M_Lmar
End Property

Public Property Let Lmar(ByVal New_Lmar As Double)
  M_Lmar = New_Lmar
  PropertyChanged "Lmar"
End Property

Public Property Get Rmar() As Double
  Rmar = M_Rmar
End Property

Public Property Let Rmar(ByVal New_Rmar As Double)
  M_Rmar = New_Rmar
  PropertyChanged "Rmar"
End Property
