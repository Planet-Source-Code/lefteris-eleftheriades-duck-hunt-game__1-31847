VERSION 5.00
Begin VB.Form LoadingFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   LinkTopic       =   "Form2"
   ScaleHeight     =   855
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   195
      Left            =   180
      ScaleHeight     =   135
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   480
      Width           =   3255
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   0
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   3435
      TabIndex        =   1
      Top             =   90
      Width           =   3435
      Begin VB.Label Label1 
         Caption         =   "Loading Dlls"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   90
         Width           =   3405
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3645
   End
End
Attribute VB_Name = "LoadingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Sub ProgressbarValue(Value&, Max&)
  Shape1.Width = (Picture2.Width / Max&) * Value&
End Sub

Sub Comment(Comments$)
 Label1.Caption = Comments$
End Sub
Private Sub Form_Activate()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
DoEvents
Load Form1
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
