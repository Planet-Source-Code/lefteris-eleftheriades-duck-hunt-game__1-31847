VERSION 5.00
Begin VB.Form PauseFrm 
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Lucida Blackletter"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "Pause.frx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1680
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   540
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   120
      Picture         =   "Pause.frx":1D11A
      ScaleHeight     =   154
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   91
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image Image5 
      Height          =   405
      Left            =   1350
      Top             =   600
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   3090
      Top             =   630
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   1590
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   1410
      Picture         =   "Pause.frx":27764
      Top             =   1800
      Width           =   210
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   1530
      Top             =   2130
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   2100
      Top             =   30
      Width           =   1395
   End
End
Attribute VB_Name = "PauseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Weapon As New CaracterObject
Dim OX As Single
Dim WID As Long
Private Sub Form_Activate()
Label1.Caption = RelativeSpeedPublicVar
Image3.Left = (RelativeSpeedPublicVar * 144) + 1440
Picture2.Cls
Weapon.SpriteDataFile = App.Path & "\Guns.spr"
Weapon.Draw 1, 0, 0, Picture2.hdc, Picture1.hdc, 0
End Sub

Private Sub Form_Load()
 Dim K(5) As pointapi
 K(0).X = 0 / Screen.TwipsPerPixelX
 K(0).Y = 315 / Screen.TwipsPerPixelY
 
 K(1).X = 2100 / Screen.TwipsPerPixelX
 K(1).Y = 315 / Screen.TwipsPerPixelY
 
 K(2).X = 2385 / Screen.TwipsPerPixelX
 K(2).Y = 30 / Screen.TwipsPerPixelY
 
 K(3).X = 3480 / Screen.TwipsPerPixelX
 K(3).Y = 30 / Screen.TwipsPerPixelY
 
 K(4).X = 3480 / Screen.TwipsPerPixelX
 K(4).Y = Me.Height / Screen.TwipsPerPixelY
 
 K(5).X = 0 / Screen.TwipsPerPixelX
 K(5).Y = Me.Height / Screen.TwipsPerPixelY
 
 SetWindowRgn Me.hwnd, CreatePolygonRgn(K(0), 6, 0), True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DoEvents
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0)
End Sub

Private Sub Image2_Click()
  RelativeSpeedPublicVar = Round((Image3.Left - 1440) / 144)
  DoEvents
  Unload Me
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 OX = X
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 If (Image3.Left + (X - OX)) < 1440 Then Image3.Left = 1440: Exit Sub
 If (Image3.Left + (X - OX)) > 2880 Then Image3.Left = 2880: Exit Sub
 Image3.Left = Image3.Left + (X - OX)
 Label1.Caption = Round((Image3.Left - 1440) / 144)
End If
Label1.Caption = Round((Image3.Left - 1440) / 144)
End Sub

Private Sub Image4_Click()
 WID = WID + 1
 If WID = 6 Then WID = 1
 Picture2.Cls
 Weapon.Draw WID, 0, 0, Picture2.hdc, Picture1.hdc, 0
 Picture2.Refresh
End Sub

Private Sub Image5_Click()
 WID = WID - 1
 If WID <= 0 Then WID = 5
 Picture2.Cls
 Weapon.Draw WID, 0, 0, Picture2.hdc, Picture1.hdc, 0
 Picture2.Refresh
End Sub
