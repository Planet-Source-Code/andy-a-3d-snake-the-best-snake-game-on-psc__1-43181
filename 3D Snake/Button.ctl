VERSION 5.00
Begin VB.UserControl Button 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Text Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   2550
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   2850
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_FillColor = vbBlue
Const m_def_FillColorMouseOver = 255
Const m_def_FillColorMouseDown = 16581375
'Property Variables:
Dim m_FillColor As OLE_COLOR
Dim m_FillColorMouseOver As OLE_COLOR
Dim m_FillColorMouseDown As OLE_COLOR
Dim m_ForeColorInvert As Boolean
Dim m_FontChangeMouseDown As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Dim CtlState As Integer '0-Normal,1-Over,2-Down
Dim R As String, G As String, B As String 'For ForeColor Invert
Dim NormalFontSize As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
    Me.Reset
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    UserControl.Refresh
    Shape1.Refresh
    Label1.Refresh
End Sub

Public Sub Reset()
    Shape1.FillColor = Me.FillColor
    If m_ForeColorInvert = True Then
        ExtractColors Shape1.FillColor
        Label1.ForeColor = RGB(255 - R, 255 - G, 255 - B)
    End If
    On Error Resume Next
    Label1.FontSize = NormalFontSize
    UserControl_Resize
    CtlState = 0
End Sub

Public Sub ForceMouseOver(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    UserControl.SetFocus
    Shape1.FillColor = m_FillColorMouseOver
    If m_ForeColorInvert = True Then
        ExtractColors Shape1.FillColor
        Label1.ForeColor = RGB(255 - R, 255 - G, 255 - B)
    End If
    CtlState = 1
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub ForceMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Shape1.FillColor = m_FillColorMouseDown
    If m_ForeColorInvert = True Then
        ExtractColors Shape1.FillColor
        Label1.ForeColor = RGB(255 - R, 255 - G, 255 - B)
    End If
    NormalFontSize = Label1.FontSize
    Label1.Font.Size = Label1.Font.Size + m_FontChangeMouseDown: UserControl_Resize
    CtlState = 2
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub usercontrol_Click()
    RaiseEvent Click
End Sub

Private Sub usercontrol_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Me.Reset
End Sub

Private Sub usercontrol_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub usercontrol_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub usercontrol_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    Me.Reset
    MOverButton = False
End Sub

Private Sub usercontrol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ForceMouseDown Button, Shift, X, Y
End Sub

Private Sub usercontrol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CtlState = 0 Then Me.ForceMouseOver Button, Shift, X, Y
End Sub

Private Sub usercontrol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Reset
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Label1_Click()
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Label1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Label1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ForceMouseDown Button, Shift, X, Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CtlState = 0 Then Me.ForceMouseOver Button, Shift, X, Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Reset
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,Shape
Public Property Get Shape() As Integer
    Shape = Shape1.Shape
End Property

Public Property Let Shape(ByVal New_Shape As Integer)
    Shape1.Shape() = New_Shape
    UserControl_Resize
    Me.Refresh
    PropertyChanged "Shape"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_FillColor = New_FillColor
    Me.Reset
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
    UserControl.Size Width, Height
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,255
Public Property Get FillColorMouseOver() As OLE_COLOR
    FillColorMouseOver = m_FillColorMouseOver
End Property

Public Property Let FillColorMouseOver(ByVal New_FillColorMouseOver As OLE_COLOR)
    m_FillColorMouseOver = New_FillColorMouseOver
    PropertyChanged "FillColorMouseOver"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,16581375
Public Property Get FillColorMouseDown() As OLE_COLOR
    FillColorMouseDown = m_FillColorMouseDown
End Property

Public Property Let FillColorMouseDown(ByVal New_FillColorMouseDown As OLE_COLOR)
    m_FillColorMouseDown = New_FillColorMouseDown
    PropertyChanged "FillColorMouseDown"
End Property

Public Property Get ForeColorInvert() As Boolean
    ForeColorInvert = m_ForeColorInvert
End Property

Public Property Let ForeColorInvert(ByVal New_ForeColorInvert As Boolean)
    m_ForeColorInvert = New_ForeColorInvert
    PropertyChanged "ForeColorInvert"
End Property

Public Property Get FontChangeMouseDown() As Integer
    FontChangeMouseDown = m_FontChangeMouseDown
End Property

Public Property Let FontChangeMouseDown(ByVal New_FontChangeMouseDown As Integer)
    m_FontChangeMouseDown = New_FontChangeMouseDown
    PropertyChanged "FontChangeMouseDown"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FillColor = m_def_FillColor
    m_FillColorMouseOver = m_def_FillColorMouseOver
    m_FillColorMouseDown = m_def_FillColorMouseDown
    m_ForeColorInvert = False
    m_FontChangeMouseDown = 0
    Me.Reset
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80&)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label1.Caption = PropBag.ReadProperty("Caption", "Your Text Here")
    Shape1.Shape = PropBag.ReadProperty("Shape", 2)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_FillColorMouseOver = PropBag.ReadProperty("FillColorMouseOver", m_def_FillColorMouseOver)
    m_FillColorMouseDown = PropBag.ReadProperty("FillColorMouseDown", m_def_FillColorMouseDown)
    m_ForeColorInvert = PropBag.ReadProperty("ForeColorInvert")
    m_FontChangeMouseDown = PropBag.ReadProperty("FontChangeMouseDown")
End Sub

Private Sub UserControl_Show()
    Me.Reset
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80&)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Your Text Here")
    Call PropBag.WriteProperty("Shape", Shape1.Shape, 2)
    Call PropBag.WriteProperty("FillColor", m_FillColor) ', m_def_FillColor)
    Call PropBag.WriteProperty("FillColorMouseOver", m_FillColorMouseOver) ', m_def_FillColorMouseOver)
    Call PropBag.WriteProperty("FillColorMouseDown", m_FillColorMouseDown) ', m_def_FillColorMouseDown)
    Call PropBag.WriteProperty("ForeColorInvert", m_ForeColorInvert)
    Call PropBag.WriteProperty("FontChangeMouseDown", m_FontChangeMouseDown)
    Me.Reset
End Sub

Private Sub UserControl_Resize()
    Shape1.Move 15, 15, UserControl.Width - 30, UserControl.Height - 30
    Label1.Left = (Shape1.Width - Label1.Width) / 2
    Label1.Top = (Shape1.Height - Label1.Height) / 2
End Sub

Private Sub ExtractColors(TestColor As Long)
    '--------------------------------------------------------
    'This code splits the one color variable into three
    'separate variables into RGB mode.  Copy the code from
    'the first line of dashes to the last line of dashes.
    '--------------------------------------------------------
    R = (TestColor Mod 256)
    B = (Int(TestColor / 65536))
    G = ((TestColor - (B * 65536) - R) / 256)
    '--------------------------------------------------------
    If R < 0 Then R = 0
    If G < 0 Then G = 0
    If B < 0 Then B = 0
    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255
End Sub

