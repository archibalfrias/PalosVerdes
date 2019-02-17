VERSION 5.00
Begin VB.UserControl b8LineVertical 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Line Line2 
      BorderColor     =   &H00F6F8F8&
      X1              =   4
      X2              =   4
      Y1              =   2
      Y2              =   152
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00F6F8F8&
      X1              =   4
      X2              =   4
      Y1              =   2
      Y2              =   144
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00ACD0D7&
      X1              =   4
      X2              =   4
      Y1              =   8
      Y2              =   152
   End
End
Attribute VB_Name = "b8LineVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub UserControl_InitProperties()
     'MsgBox "b8Line" & vbNewLine & "Code By: ROLLIE A. JABONERO"
End Sub

Private Sub UserControl_Resize()

'    UserControl.Height = Screen.TwipsPerPixelY * 4
    
'    Line1.X1 = 0
'    Line1.X2 = UserControl.Width / Screen.TwipsPerPixelX
'    Line1.Y1 = 2
'    Line1.Y2 = 2
    
'    Line2.X1 = 0
'    Line2.X2 = UserControl.Width / Screen.TwipsPerPixelX
'    Line2.Y1 = 1
'    Line2.Y2 = 1
    
'    Line3.X1 = 0
'    Line3.X2 = UserControl.Width / Screen.TwipsPerPixelX
'    Line3.Y1 = 3
'    Line3.Y2 = 3
    
    UserControl.Width = Screen.TwipsPerPixelX * 4
    
    Line1.X1 = 2 '0
    Line1.X2 = 2 'UserControl.Width / Screen.TwipsPerPixelX
    Line1.Y1 = 0 '2
    Line1.Y2 = UserControl.Height / Screen.TwipsPerPixelY '2
    
    Line2.X1 = 1 '0
    Line2.X2 = 1 'UserControl.Width / Screen.TwipsPerPixelX
    Line2.Y1 = 0 '1
    Line2.Y2 = UserControl.Height / Screen.TwipsPerPixelY '1
    
    Line3.X1 = 3 '0
    Line3.X2 = 3 'UserControl.Width / Screen.TwipsPerPixelX
    Line3.Y1 = 0 '3
    Line3.Y2 = UserControl.Height / Screen.TwipsPerPixelY '3
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderColor
Public Property Get BorderColor1() As OLE_COLOR
    BorderColor1 = Line1.BorderColor
End Property

Public Property Let BorderColor1(ByVal New_BorderColor1 As OLE_COLOR)
    Line1.BorderColor() = New_BorderColor1
    PropertyChanged "BorderColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderColor
Public Property Get BorderColor2() As OLE_COLOR
    BorderColor2 = Line2.BorderColor
End Property

Public Property Let BorderColor2(ByVal New_BorderColor2 As OLE_COLOR)
    Line2.BorderColor() = New_BorderColor2
    Line3.BorderColor() = New_BorderColor2
    PropertyChanged "BorderColor2"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Line1.BorderColor = PropBag.ReadProperty("BorderColor1", 11325655)
    Line2.BorderColor = PropBag.ReadProperty("BorderColor2", 16185592)
    Line3.BorderColor = PropBag.ReadProperty("BorderColor3", 16185592)

    Line1.BorderStyle = PropBag.ReadProperty("BorderStyle1", 1)
    Line2.BorderStyle = PropBag.ReadProperty("BorderStyle2", 1)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BorderColor1", Line1.BorderColor, 11325655)
    Call PropBag.WriteProperty("BorderColor2", Line2.BorderColor, 16185592)
    Call PropBag.WriteProperty("BorderColor3", Line2.BorderColor, 16185592)

    Call PropBag.WriteProperty("BorderStyle1", Line1.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderStyle2", Line2.BorderStyle, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line1,Line1,-1,BorderStyle
Public Property Get BorderStyle1() As BorderStyleConstants
    BorderStyle1 = Line1.BorderStyle
End Property

Public Property Let BorderStyle1(ByVal New_BorderStyle1 As BorderStyleConstants)
    Line1.BorderStyle() = New_BorderStyle1
    PropertyChanged "BorderStyle1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Line2,Line2,-1,BorderStyle
Public Property Get BorderStyle2() As BorderStyleConstants
    BorderStyle2 = Line2.BorderStyle
    
End Property

Public Property Let BorderStyle2(ByVal New_BorderStyle2 As BorderStyleConstants)
    Line2.BorderStyle() = New_BorderStyle2
    Line3.BorderStyle() = New_BorderStyle2
    PropertyChanged "BorderStyle2"
End Property


