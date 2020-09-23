VERSION 5.00
Begin VB.UserControl ScrollLabel 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   480
   ScaleWidth      =   2265
   ToolboxBitmap   =   "ScrollLabel.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll Label"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "ScrollLabel.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2280
   End
End
Attribute VB_Name = "ScrollLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum Alignment
[Left Justify]
[Right Justify]
center
End Enum
Dim m_Caption As String
Dim m_Alignment As Alignment
Const m_def_Caption = "Scroll Label"
Const m_def_Alignment = 2
'Event Declarations:
Event DblClick()
Event Click()
Event Change()
Event OLECompleteDrag(Effect As Long)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()
Dim Move As Integer
Public Property Get Caption() As String
    Caption = Label1.Caption
   End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Appearance() As Integer
    Appearance = UserControl.Appearance
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label1.ForeColor
  End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   Label1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub Label1_Change()
RaiseEvent Change
End Sub

Private Sub Label1_Click()
RaiseEvent Click
End Sub



Private Sub Label1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub label1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property



Private Sub Label1_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub label1_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub label1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Public Property Get OLEDropMode() As Integer
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    Label1.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub label1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub label1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Public Property Get Picture() As Picture
    Set Picture = Image1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    PropertyChanged "Picture"
End Property



Private Sub UserControl_Resize()
    RaiseEvent Resize
        With Image1
        .Height = UserControl.ScaleHeight
        .Top = UserControl.ScaleTop
        .Left = UserControl.ScaleLeft
        .Width = UserControl.ScaleWidth
    End With
            With Label1
        .Height = UserControl.ScaleHeight
        .Top = UserControl.ScaleTop
        .Left = UserControl.ScaleLeft
        .Width = UserControl.ScaleWidth
    End With
With Image2
        .Height = UserControl.ScaleHeight
        .Top = UserControl.ScaleTop
        .Left = UserControl.ScaleLeft
        .Width = UserControl.ScaleWidth
    End With
    
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
 Set Font = Ambient.Font
  m_Caption = m_def_Caption
 m_Alignment = m_def_Alignment
UserControl.BorderStyle = 1
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Label1.Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   Label1.Alignment = PropBag.ReadProperty("Alignment", m_def_TextAlignment)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", Label1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", Label1.Caption, m_def_Caption)
End Sub

Sub Speed(Interval As Integer)
Timer1.Interval = Interval
End Sub

Sub StartIt()
Timer1.Enabled = True
End Sub
Sub StopIt()
Timer1.Enabled = False
End Sub


Private Sub Timer1_Timer()
Image2.Picture = Image1.Picture
Image1.Left = Image1.Left - Move
Image2.Left = Image2.Left - Move
If Image2.Left <= 0 Then Image1.Left = Image2.Left + Image2.Width
If Image2.Left <= 0 - Image2.Width Then Image2.Left = Image1.Left + Image1.Width
End Sub
Sub Movement(Interval As Integer)
Move = Interval
End Sub
Public Property Get Alignment() As Alignment
   Alignment = Label1.Alignment
    End Property

Public Property Let Alignment(ByVal New_Alignment As Alignment)
    Label1.Alignment = New_Alignment
    PropertyChanged "Alignment"
End Property

