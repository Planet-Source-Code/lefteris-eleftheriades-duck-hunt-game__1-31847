VERSION 5.00
Begin VB.UserControl Radio 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ScaleHeight     =   1005
   ScaleWidth      =   1755
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   0
      Picture         =   "Option.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   1140
      Picture         =   "Option.ctx":0342
      Top             =   210
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   900
      Picture         =   "Option.ctx":0684
      Top             =   210
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Radio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const winding = 2
Const alternate = 1
Const rgn_or = 2

Private Type pointapi
        X As Long
        Y As Long
End Type

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As pointapi, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Private Declare Function CreatePolyPolygonRgn& Lib "gdi32" (lpPoint As pointapi, ByVal nCount As Long, ByVal nPolyfillMode As Long, lpPolyCount As Long)
'Default Property Values:
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Boolean
'Event Declarations:
Event Click() 'MappingInfo=Image1(2),Image1,2,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Private Sub Image1_Click(Index As Integer)
If Index = 2 Then
 If m_Value Then
    Image1(2).Picture = Image1(0).Picture
 Else
    Image1(2).Picture = Image1(1).Picture
 End If
 m_Value = Not m_Value
 PropertyChanged "Value"
 RaiseEvent Click
End If
End Sub

Private Sub UserControl_Initialize()
 SetWindowRgn UserControl.hwnd, CreateEllipticRgn&(0, 0, UserControl.Width, UserControl.Height), True
End Sub

Private Sub UserControl_Resize()
If UserControl.Width > 930 Then UserControl.Width = 930
If UserControl.Height > 930 Then UserControl.Height = 930
Image1(2).Move 0, 0, UserControl.Width - 30, UserControl.Height - 30
SetWindowRgn UserControl.hwnd, CreateEllipticRgn&(0, 0, UserControl.Width / 15, UserControl.Height / 15), True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1(2),Image1,2,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Image1(2).Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Image1(2).Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image1(2),Image1,2,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Image1(2).ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Image1(2).ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
    If m_Value Then
      Image1(2).Picture = Image1(0).Picture
    Else
      Image1(2).Picture = Image1(1).Picture
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Image1(2).Enabled = PropBag.ReadProperty("Enabled", True)
    Image1(2).ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", Image1(2).Enabled, True)
    Call PropBag.WriteProperty("ToolTipText", Image1(2).ToolTipText, "")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

