VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Target"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Target"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarLeftPos As Long 'local copy
Private mvarTopPos As Long 'local copy
Private mvarWidthOfState As Long 'local copy
Private mvarHeightOfState As Long 'local copy
Private mvarSpinState As Long 'local copy
Private mvarTarget As Long 'local copy
'To fire this event, use RaiseEvent with the following syntax:
Public Event Shot(ByVal Skill As Long)

Public Enum TargetType
 DartBoard = 1
 Duck = 2
 Ammon = 3
 Piano = 5
 Bunny = 6
 Pokemon = 7
End Enum

'local variable(s) to hold property value(s)
Private mvarSpeed As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarTargetImages As Object 'local copy
Private mvarDataFilename As String 'local copy

Friend Property Let Data(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataFilename = 5
    mvarDataFilename = vData
End Property

Friend Property Get Data() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataFilename
    Data = mvarDataFilename
End Property

Friend Property Set Image(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TargetImages = Form1
    Set mvarTargetImages = vData
End Property

Friend Property Get Image() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TargetImages
    Set Image = mvarTargetImages
End Property

Friend Property Let Speed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Speed = 5
    mvarSpeed = vData
End Property

Friend Property Get Speed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speed
    Speed = mvarSpeed
End Property

Friend Sub Draw(Target As TargetType, State, x As Single, y As Single)

End Sub

Public Sub Action()
 'RaiseEvent Shot[(arg1, arg2, ... , argn)]
 If Target = DartBoard Then
    RaiseEvent Shot(1)
 End If
End Sub

Public Property Let Target(ByVal vData As TargetType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Target = 5
    mvarTarget = vData
End Property

Public Property Get Target() As TargetType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Target
    Target = mvarTarget
End Property

Friend Property Let SpinState(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpinState = 5
    mvarSpinState = vData
End Property

Friend Property Get SpinState() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpinState
    SpinState = mvarSpinState
End Property

Friend Property Let HeightOfState(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HeightOfState = 5
    mvarHeightOfState = vData
End Property

Friend Property Get HeightOfState() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HeightOfState
    HeightOfState = mvarHeightOfState
End Property

Friend Property Let WidthOfState(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WidthOfState = 5
    mvarWidthOfState = vData
End Property

Friend Property Get WidthOfState() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WidthOfState
    WidthOfState = mvarWidthOfState
End Property

Friend Property Let TopPos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TopPos = 5
    mvarTopPos = vData
End Property

Friend Property Get TopPos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopPos
    TopPos = mvarTopPos
End Property

Friend Property Let LeftPos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LeftPos = 5
    mvarLeftPos = vData
End Property

Friend Property Get LeftPos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LeftPos
    LeftPos = mvarLeftPos
End Property
