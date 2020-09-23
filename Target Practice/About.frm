VERSION 5.00
Begin VB.Form AboutFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form5"
   Picture         =   "About.frx":0000
   ScaleHeight     =   2370
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":1B04A
      BeginProperty Font 
         Name            =   "Lucida Blackletter"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   2130
      Top             =   30
      Width           =   1305
   End
End
Attribute VB_Name = "AboutFrm"
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
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0)
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
