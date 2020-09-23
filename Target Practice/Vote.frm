VERSION 5.00
Begin VB.Form VoteFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form6"
   Picture         =   "Vote.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TargetShooter.Radio Radio1 
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   870
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin TargetShooter.Radio Radio1 
      Height          =   285
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   1170
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      Value           =   -1  'True
   End
   Begin TargetShooter.Radio Radio1 
      Height          =   285
      Index           =   2
      Left            =   150
      TabIndex        =   2
      Top             =   1470
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin TargetShooter.Radio Radio1 
      Height          =   285
      Index           =   3
      Left            =   150
      TabIndex        =   3
      Top             =   1770
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin TargetShooter.Radio Radio1 
      Height          =   285
      Index           =   4
      Left            =   150
      TabIndex        =   4
      Top             =   2040
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin VB.Image Image2 
      Height          =   525
      Left            =   2760
      Top             =   1770
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   2070
      Top             =   30
      Width           =   1725
   End
End
Attribute VB_Name = "VoteFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Dim K(5) As pointapi
 K(0).X = 0 / Screen.TwipsPerPixelX
 K(0).Y = 315 / Screen.TwipsPerPixelY
 
 K(1).X = 2100 / Screen.TwipsPerPixelX
 K(1).Y = 315 / Screen.TwipsPerPixelY
 
 K(2).X = 2385 / Screen.TwipsPerPixelX
 K(2).Y = 30 / Screen.TwipsPerPixelY
 
 K(3).X = Me.Width / Screen.TwipsPerPixelX
 K(3).Y = 30 / Screen.TwipsPerPixelY
 
 K(4).X = Me.Width / Screen.TwipsPerPixelX
 K(4).Y = Me.Height / Screen.TwipsPerPixelY
 
 K(5).X = 0 / Screen.TwipsPerPixelX
 K(5).Y = Me.Height / Screen.TwipsPerPixelY
 
 SetWindowRgn Me.hwnd, CreatePolygonRgn(K(0), 6, 0), True
 Radio1(1).Value = True
End Sub

Private Sub Image2_Click()
 Dim SelRadio As Long
 Dim i%
 For i% = 0 To 4
   If Radio1(i%).Value Then SelRadio = i%
 Next i%
 Select Case SelRadio
   Case 0: MsgBox "Thank you :-). It is good to be appriciated."
   Case 1: MsgBox "Thank you!"
   Case 2: MsgBox "Only? :-|"
   Case 3: MsgBox "Grrrr. Only Average?"
   Case 4: MsgBox "Fuck You."
 End Select
 Unload Me
End Sub

Private Sub Radio1_Click(Index As Integer)
 For Each Control In Controls
    If TypeOf Control Is Radio Then
        Control.Value = False
    End If
 Next Control
 Radio1(Index).Value = True
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0)
End Sub
