VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12540
   DrawWidth       =   4
   BeginProperty Font 
      Name            =   "Lucida Blackletter"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   608
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   836
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox TargetRandarView 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   11520
      Picture         =   "Form1.frx":1594
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox TrgPts 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   5430
      Picture         =   "Form1.frx":4A6E
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox GhostIMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2130
      Picture         =   "Form1.frx":6B30
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   22
      Top             =   210
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox GhostMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2460
      Picture         =   "Form1.frx":702A
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox GhostBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4830
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   20
      Top             =   210
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox GhostSpr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   2760
      Picture         =   "Form1.frx":7524
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   182
      TabIndex        =   19
      Top             =   930
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Blackletter"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4710
      Visible         =   0   'False
      Width           =   5085
   End
   Begin VB.PictureBox CheatMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Index           =   0
      Left            =   -30
      Picture         =   "Form1.frx":E06E
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   366
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   5520
   End
   Begin VB.PictureBox CheatMenu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":2092C
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   366
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   5520
   End
   Begin VB.PictureBox NumMsk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   1920
      Picture         =   "Form1.frx":21666
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox NumMsk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":21B30
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox NumPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   1920
      Picture         =   "Form1.frx":21FFA
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox NumPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   0
      Left            =   510
      Picture         =   "Form1.frx":284BC
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   12
      Top             =   2970
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   0
      Left            =   720
      Picture         =   "Form1.frx":2E97E
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   8
      Top             =   7950
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1650
      Top             =   120
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   1
      Left            =   870
      Picture         =   "Form1.frx":5BBA0
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   1
      Left            =   -30
      Picture         =   "Form1.frx":88DC2
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   10
      Top             =   6690
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   0
      Left            =   30
      Picture         =   "Form1.frx":8AC20
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   9
      Top             =   7800
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   4650
      Picture         =   "Form1.frx":8CA7E
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   535
      TabIndex        =   6
      Top             =   3090
      Visible         =   0   'False
      Width           =   8025
   End
   Begin VB.PictureBox Buf 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2490
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   5
      Top             =   3270
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1650
      Top             =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7170
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   4410
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7680
      Left            =   7110
      ScaleHeight     =   512
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   4
      Top             =   4350
      Visible         =   0   'False
      Width           =   7665
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   7050
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   3
      Top             =   4290
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7200
      Left            =   6990
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   6930
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   4170
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox Msk 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   3750
      Picture         =   "Form1.frx":110F88
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   535
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   8025
   End
   Begin VB.Image TrgRadar 
      Height          =   540
      Index           =   3
      Left            =   2190
      Picture         =   "Form1.frx":116956
      Top             =   1500
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image TrgRadar 
      Height          =   1335
      Index           =   2
      Left            =   1260
      Picture         =   "Form1.frx":1179E8
      Top             =   1050
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image TrgRadar 
      Height          =   900
      Index           =   1
      Left            =   270
      Picture         =   "Form1.frx":11BB86
      Top             =   2040
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Image TrgRadar 
      Height          =   990
      Index           =   0
      Left            =   270
      Picture         =   "Form1.frx":122108
      Top             =   1050
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image TargetNo 
      Height          =   360
      Index           =   3
      Left            =   7800
      Picture         =   "Form1.frx":1255E2
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image TargetNo 
      Height          =   810
      Index           =   4
      Left            =   5340
      Picture         =   "Form1.frx":125D44
      Top             =   30
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image TargetNo 
      Height          =   1095
      Index           =   2
      Left            =   6240
      Picture         =   "Form1.frx":128606
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image TargetNo 
      Height          =   630
      Index           =   1
      Left            =   6240
      Picture         =   "Form1.frx":12B2C4
      Top             =   1110
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Image TargetNo 
      Height          =   780
      Index           =   0
      Left            =   7020
      Picture         =   "Form1.frx":12F15E
      Top             =   30
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image TrgRadar 
      Height          =   1005
      Index           =   4
      Left            =   2400
      Picture         =   "Form1.frx":131220
      Top             =   2040
      Visible         =   0   'False
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'————————————————————————————————————————————————————'
' This game is based on the class module called:     '
' "caracter object.cls" this module loads the sprite '
' data saved in XXXXXX.SPR files. which are created  '
' with Caracter Sprite Editor Made by Me. That       '
' Program's Code has also freely distributed source  '
' Code. The two Class Modules were created to        '
' fullfill the need of reading those files.          '
'————————————————————————————————————————————————————'

Dim Statusbar As Long                'What of the 2 status bar pics to show?
Dim Target1 As New CaracterObject    'Set Target1 as a Caractrer Object
Dim Target2 As New CaracterObject    'Set Target1 as a Caractrer Object
'VDL = Visual Display Label
Dim ScoreVDL As New DigitalCounter   '
Dim BulletsVDL As New DigitalCounter '
Dim TimeVDL As New DigitalCounter    '
Dim Ghost As New CaracterObject      '
Dim XX&, YY&                         '
Dim TXX&, TYY&                         '
Dim FreeShots As Long                '
Dim Tme As Long                      '
Dim Level As Long                    '
Dim Score As Long                    '
Dim IsInPause As Boolean             '
Dim IsInCheatMode As Boolean         '
Dim RelativeSpeed As Long            '
Dim AutoPilot As Boolean             '
Dim PMX As Integer                   '
Dim PMY As Integer                   '
Dim ACL As Single                    '
Dim ACT As Single                    '
Dim ACLL As Single                   '
Dim ACTT As Single                   '
Dim PStage As Long                   '
Dim Mx, My As Single                 '
Dim ShowRadar As Boolean             '
Dim MusicEnabled As Boolean          '
Dim SoundEnabled As Boolean          '
'Dim dx As New DirectX7
'Dim dd As DirectDraw7

Sub SS(ApplicationHWnd As Long)
  Set dd = dx.DirectDrawCreate("")
  Call dd.SetCooperativeLevel(ApplicationHWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
  dd.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
End Sub

Sub RandomSound()
    Randomize
    Select Case Int(Rnd * 13.99) 'Get a random number
     Case 0: PlayLargeSound App.Path & "\Game Sounds\Western.mid", MIDI_Sequence, "BckSnd"
     Case 1: PlayLargeSound App.Path & "\Game Sounds\Action.mid", MIDI_Sequence, "BckSnd"
     Case 2: PlayLargeSound App.Path & "\Game Sounds\mission.mid", MIDI_Sequence, "BckSnd"
     Case 3: PlayLargeSound App.Path & "\Game Sounds\win.mid", MIDI_Sequence, "BckSnd"
     Case 4: PlayLargeSound App.Path & "\Game Sounds\On a MotorBike.mid", MIDI_Sequence, "BckSnd"
     Case 5: PlayLargeSound App.Path & "\Game Sounds\JamesBond.mid", MIDI_Sequence, "BckSnd"
     Case 6: PlayLargeSound App.Path & "\Game Sounds\swimm.mid", MIDI_Sequence, "BckSnd"
     Case 7: PlayLargeSound App.Path & "\Game Sounds\boss.mid", MIDI_Sequence, "BckSnd"
     Case 8: PlayLargeSound App.Path & "\Game Sounds\Level2.mid", MIDI_Sequence, "BckSnd"
     Case 9: PlayLargeSound App.Path & "\Game Sounds\intro.mid", MIDI_Sequence, "BckSnd"
     Case 10: PlayLargeSound App.Path & "\Game Sounds\lastlevel.mid", MIDI_Sequence, "BckSnd"
     Case 11: PlayLargeSound App.Path & "\Game Sounds\jammin.mid", MIDI_Sequence, "BckSnd"
     Case 12: PlayLargeSound App.Path & "\Game Sounds\loose.mid", MIDI_Sequence, "BckSnd"
     Case 13: PlayLargeSound App.Path & "\Game Sounds\everybody.mid", MIDI_Sequence, "BckSnd"
    End Select
    DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim XXX&, YYY&
 Select Case KeyCode
   Case vbKeyEscape: StopLargeSound "BckSnd": DoEvents: End
   Case vbKeyF3
     Timer1.Enabled = False
     Timer2.Enabled = False
     'XXX& = ((Me.Width / Screen.TwipsPerPixelX) - pauseMenu(0).Width) / 2
     'YYY& = ((Me.Height / Screen.TwipsPerPixelY) - pauseMenu(0).Height) / 2
     'BitBlt Me.hdc, XXX&, YYY&, pauseMenu(0).Width, pauseMenu(0).Height, pauseMenu(1).hdc, 0, 0, SRCAND
     'BitBlt Me.hdc, XXX&, YYY&, pauseMenu(0).Width, pauseMenu(0).Height, pauseMenu(0).hdc, 0, 0, SRCINVERT
     'Me.Refresh
     IsInPause = True
     'Me.MousePointer = vbDefault
     RelativeSpeedPublicVar = RelativeSpeed
     PauseFrm.Show vbModal, Me
     RelativeSpeed = RelativeSpeedPublicVar
     DoEvents
     Timer1.Enabled = True
     Timer2.Enabled = True
     IsInPause = False
   Case vbKeyF12
     If Not IsInPause Then
       Timer1.Enabled = False
       Timer2.Enabled = False
       XXX& = ((Me.Width / Screen.TwipsPerPixelX) - CheatMenu(0).Width) / 2
       YYY& = ((Me.Height / Screen.TwipsPerPixelY) - CheatMenu(0).Height) / 2
       BitBlt Me.hdc, XXX&, YYY&, CheatMenu(0).Width, CheatMenu(0).Height, CheatMenu(1).hdc, 0, 0, SRCAND
       BitBlt Me.hdc, XXX&, YYY&, CheatMenu(0).Width, CheatMenu(0).Height, CheatMenu(0).hdc, 0, 0, SRCINVERT
       Me.Refresh
       IsInCheatMode = True
       Me.MousePointer = vbDefault
       Text1.Visible = True
       Text1.Move XXX& + 12, YYY& + 34
     End If
   Case vbKeyR: ShowRadar = Not ShowRadar
   Case vbKeyS: SoundEnabled = Not SoundEnabled
   Case vbKeyM
     MusicEnabled = Not MusicEnabled
     If MusicEnabled = False Then StopLargeSound "BckSnd"
   Case vbKeyF1
     Timer1.Enabled = False
     Timer2.Enabled = False
     Help.Show vbModal, Me
     Timer1.Enabled = True
     Timer2.Enabled = True
   Case vbKeyF10
     Timer1.Enabled = False
     Timer2.Enabled = False
     AboutFrm.Show vbModal, Me
     Timer1.Enabled = True
     Timer2.Enabled = True
   Case vbKeyF11
     Timer1.Enabled = False
     Timer2.Enabled = False
     VoteFrm.Show vbModal, Me
     Timer1.Enabled = True
     Timer2.Enabled = True
   Case vbKeyF2
     Timer1.Enabled = False
     Timer2.Enabled = False
     If MsgBox("Don't be stupid." & vbCrLf & "Restart Game?", vbExclamation Or vbYesNo Or vbDefaultButton2) = vbYes Then
        'Reset Score and Level
        Score = 0
        Level = 1
        'Load the first stage
        LoadStage 0
        'Reset the target's picture score
        'Red = 3 pts
        'White = 2 pts
        'Blue = 1 pt
        TrgPts.Picture = TargetNo(0).Picture
        'Reset target's picture.
        TargetRandarView.Picture = TrgRadar(0).Picture
     End If
     Timer1.Enabled = True
     Timer2.Enabled = True
 End Select
End Sub

Private Sub Form_Initialize()
 '''''''''''''''''''''''''''''''''''''''''''''''''''
 'Set all pictureboxes to autoredraw & vbpixels    '
 'AutoRedraw is needed for all API Paint Functions '
 'vbPixels is needed for all API Functions         '
 '''''''''''''''''''''''''''''''''''''''''''''''''''
 For Each Control In Me.Controls
     If TypeOf Control Is Picturebox Then
        Control.ScaleMode = vbPixels
        Control.AutoRedraw = True
     End If
 Next Control
 LoadingFrm.ProgressbarValue 1, 15
 '''''''''''''''''''''''''''''
 '   Set form's properties   '
 '''''''''''''''''''''''''''''
 Me.Cls                  'Clear the form
 Me.AutoRedraw = True    'Need 4 draw functions
 Me.ScaleMode = vbPixels 'Need API functions
 'Me.Move 0, 0, Me.Width * Screen.TwipsPerPixelX, Me.Height * Screen.TwipsPerPixelY
 Me.Move 0, 0, 800 * Screen.TwipsPerPixelX, 600 * Screen.TwipsPerPixelY
 '''''''''''''''''''''''''''''
 LoadingFrm.ProgressbarValue 2, 15
 Set_Default_Values
End Sub

Sub Set_Default_Values()
  Score = 0            'Reset Score to 0
  Level = 1            'Set enemy to target
  Tme = 500            'Set Begin Time (Reduces every sec.)
  PMX = 1              'Move UFO to the Right
  PMY = 1              'Move UFO to the Bottom
  XX& = 400            '1st Target Fall-Pos X
  YY& = -90            '1st Target Fall-Pos Y
  FreeShots = 2        'Load the gun
  RelativeSpeed = 5    'The fall/move speed of the game
  ShowRadar = True     'Show the radar in the begining
  MusicEnabled = False 'Do not play music unless the user says otherwize
  SoundEnabled = True  'Play Sound Fx unless the user says otherwize
End Sub
Private Sub Form_Load()
 Dim ResolutionNotification As String
 'On Error Resume Next
 LoadingFrm.ProgressbarValue 3, 15
 LoadingFrm.Comment "Loading Dlls"
 DoEvents
 LoadingFrm.ProgressbarValue 4, 15
 SetStageFilesIntoPictureBoxes
 LoadingFrm.Comment "Loading Game"
 'Me.Move 0, 0, Screen.Width, Screen.Height
 Me.Move 0, 0, 800 * 15, 600 * 15
 
 LoadingFrm.ProgressbarValue 10, 15
 DoEvents
 DoEvents
 LoadingFrm.Comment "Loading Stage"
 LoadAllSpriteData
 ''<Change screen Resolution>
 'If Screen.Width <> 800 And Screen.Height <> 600 Then
 '  Select Case Screen.Width / Screen.TwipsPerPixelX
 '  Case 640
 '  ResolutionNotification = _
 '     "This program can run under 640 X 480" & vbCrLf & _
 '     "But with low quality, It is recommended" & vbCrLf & _
 '     "that you let DirectX 7 to set your resolution at 800 X 600"
 '  Case 1024
 '     ResolutionNotification = _
 '     "This program can run under 1024 X 768" & vbCrLf & _
 '     "But it will be windowed not fullscreen," & vbCrLf & _
 '     "It is recommended that you let DirectX 7" & vbCrLf & _
 '     "to set your resolution at 800 X 600"
 '  End Select
 '  If MsgBox(ResolutionNotification & vbCrLf & "Note: If you dont have directX 7 select NO" & vbCrLf & "Chage screen resolusion to 800 X 600?", vbInformation Or vbYesNo) = vbYes Then
 '     SS Me.hwnd
 '  End If
 'End If
 ''</Change screen Resolution>
 LoadStage 0
 DoEvents
 DoEvents
 Timer1.Enabled = True
 Timer2.Enabled = True
 DoEvents
 LoadingFrm.ProgressbarValue 15, 15
End Sub

Sub SetStageFilesIntoPictureBoxes()
  LoadingFrm.Comment "Loading Background 1"
  Picture1.Picture = LoadPicture(App.Path & "\Back.jpg")
  LoadingFrm.ProgressbarValue 5, 15
  DoEvents
  
  LoadingFrm.Comment "Loading Background 2"
  Picture2.Picture = LoadPicture(App.Path & "\Exotic Background.jpg")
  LoadingFrm.ProgressbarValue 6, 15
  DoEvents
  
  LoadingFrm.Comment "Loading Background 3"
  Picture3.Picture = LoadPicture(App.Path & "\Mountains Background.jpg")
  LoadingFrm.ProgressbarValue 7, 15
  DoEvents
  
  LoadingFrm.Comment "Loading Background 4"
  Picture4.Picture = LoadPicture(App.Path & "\Mountains2 Background.jpg")
  LoadingFrm.ProgressbarValue 8, 15
  DoEvents
  
  LoadingFrm.Comment "Loading Background 5"
  Picture5.Picture = LoadPicture(App.Path & "\Sky.jpg")
  LoadingFrm.ProgressbarValue 9, 15
  DoEvents
End Sub

Sub LoadAllSpriteData()
  ScoreVDL.SpriteDataFile = App.Path & "\Digital.Spr"
  BulletsVDL.SpriteDataFile = App.Path & "\Digital.Spr"
  TimeVDL.SpriteDataFile = App.Path & "\Digital.Spr"
  Target1.SpriteDataFile = App.Path & "\Targets.Spr"
  Target2.SpriteDataFile = App.Path & "\Targets.Spr"
  Ghost.SpriteDataFile = App.Path & "\Ghosts.Spr"
 
  LoadingFrm.ProgressbarValue 14, 15
  DoEvents
  LoadingFrm.Comment "Loading Sprite Data"
End Sub

Sub LoadStage(Stage As Integer)
  Cls
  Select Case Stage
    Case 0: StretchBlt Me.hdc, 0, 0, Picture1.ScaleWidth * 2, Picture1.ScaleHeight * 2.57, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, SRCCOPY
    Case 1: StretchBlt Me.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, SRCCOPY
    Case 2: StretchBlt Me.hdc, 0, 0, Picture3.ScaleWidth * 1.3, Picture3.ScaleHeight * 1.3, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, SRCCOPY
    Case 3: StretchBlt Me.hdc, 0, 0, Picture4.ScaleWidth * 1.3, Picture4.ScaleHeight * 1.3, Picture4.hdc, 0, 0, Picture4.ScaleWidth, Picture4.ScaleHeight, SRCCOPY
    Case 4: StretchBlt Me.hdc, 0, 0, Picture5.ScaleWidth * 1.6, Picture5.ScaleHeight * 1.2, Picture5.hdc, 0, 0, Picture5.ScaleWidth, Picture5.ScaleHeight, SRCCOPY
  End Select
  Me.Picture = Me.Image
  Cls
  PStage = Stage
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     If Y > 520 And Screen.Width > 800 * 15 And Screen.Height > 600 * 15 Then
      Call ReleaseCapture
      Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0)
     End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mx = X
My = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim PoC As Long
 Dim Accuracy As Integer
 Dim FakeScore As Long
 
 If Not IsInPause Then
   Accuracy = 0
   If Button = 1 Then 'If Left (Shoot)
      If FreeShots > 0 Then 'If gun is loaded with bulets
          If X > XX& And X < XX& + 100 And Y > YY& And Y < YY& + 100 Then
              ACL = X - XX&
              ACT = Y - YY&
              PoC = TrgPts.Point(ACL, ACT)
              'MAYBE AN ERROR IF COLOR MODE = 24 bit COLORS
              If PoC = RGB(255, 0, 0) Then Accuracy = 3
              If PoC = RGB(255, 255, 255) Then Accuracy = 2
              If PoC = RGB(0, 0, 255) Then Accuracy = 1
              'If the hunter hit the target
              Score = Score + Accuracy ' Show the number of shooted items
              If Level = 1 And Accuracy <> 0 Then
                 YY& = -90       'Send target to the top
                 XX& = Rnd * 750 'And to a random location
              End If
              
              If Level = 2 And Accuracy <> 0 Then
                 XX& = -90
                 YY& = Rnd * 550 'send duck to a random location
              End If
              If Level = 3 And Accuracy <> 0 Then
                 YY& = -90       'Send worm to the top
                 XX& = Rnd * 750 'And to a random location
              End If
              If Level = 4 And Accuracy <> 0 Then
                 YY& = -90       'Send ghost to the top
                 XX& = Rnd * 750 'And to a random location
              End If
          End If
          'Play the sound and remove a bulet from the gun
          If SoundEnabled Then PlaySoundPart App.Path & "\shotgun.wav", 1, 200, WaveFiles, "TPSG"
          FreeShots = FreeShots - 1
      Else
          'Play the empty gun sound
          If SoundEnabled Then PlaySoundPart App.Path & "\Empty.wav", 1, 50, WaveFiles, "TPES"
      End If
   ElseIf Button = 2 Then
      'If Right (Reload) Button was pressed
      FreeShots = 2
      'Play reload sound
      If SoundEnabled Then PlaySoundPart App.Path & "\Reload.wav", 1, 300, WaveFiles, "TPRL"
   Else
    If Statusbar = 1 Then
      Statusbar = 0
     Else
      Statusbar = 1
     End If
   End If
   '''''''''''''''''''''''''''
   FakeScore = Score
   Do Until FakeScore < 50
    FakeScore = FakeScore - 50
   Loop
   
   If (FakeScore >= 10 And FakeScore < 20) And PStage <> 1 Then
        LoadStage 1 'Level 1 2
   End If
   If (FakeScore >= 20 And FakeScore < 30) And PStage <> 2 Then
      LoadStage 2 'Level 1 3
   End If
   If (FakeScore >= 30 And FakeScore < 40) And PStage <> 3 Then
      LoadStage 3 'Level 1 4
   End If
   If (FakeScore >= 40 And FakeScore < 50) And PStage <> 4 Then
      LoadStage 4 'Level 1 5
   End If
   '''''''''''''''''
   If Score >= 50 And Score < 60 And Level <> 2 Then
      'Level 2 0
      'Speaker.Speak "Excellent Job, Let's see now how will you do with the duck"
      Timer2.Enabled = False
      LoadStage 0
      Level = 2
      Tme = 750
      Step = 0
      TrgPts.Picture = TargetNo(1).Picture
      TargetRandarView.Picture = TrgRadar(1).Picture
      DoEvents
      Timer2.Enabled = True
   End If
   If Score >= 100 And Score < 110 And Level <> 3 Then
      'Speaker.Speak "Excellent Job, Let's see now how will you do with the worm"
      Timer2.Enabled = False
      LoadStage 0
      Level = 3
      Tme = 1000
      Step = 0
      TrgPts.Picture = TargetNo(2).Picture
      TargetRandarView.Picture = TrgRadar(2).Picture
      DoEvents
      Timer2.Enabled = True
   End If
   If Score >= 150 And Score < 160 And Level <> 4 Then
      'Speaker.Speak "Excellent Job, Let's see now how will you do with the Ghost"
      Timer2.Enabled = False
      LoadStage 0
      Level = 4
      Step = 0
      StopLargeSound "BckSnd"
      DoEvents
      If MusicEnabled Then PlayLargeSound App.Path & "\Game Sounds\Ghost Level.mid", MIDI_Sequence, "BckSnd"
      TrgPts.Picture = TargetNo(3).Picture
      TargetRandarView.Picture = TrgRadar(3).Picture
      DoEvents
      Timer2.Enabled = True
   End If
   If Score >= 200 And Score < 210 And Level <> 5 Then
      'Speaker.Speak "Excellent Job, Let's see now how will you do with the UFO"
      Timer2.Enabled = False
      LoadStage 0
      Level = 5
      Step = 0
      StopLargeSound "BckSnd"
      DoEvents
      If MusicEnabled Then PlayLargeSound App.Path & "\Game Sounds\x-files.mid", MIDI_Sequence, "BckSnd"
      TargetRandarView.Picture = TrgRadar(4).Picture
      TrgPts.Picture = TargetNo(4).Picture
      DoEvents
      Timer2.Enabled = True
   End If

   '''''''''''''''
End If
'If IsInPause Then
' XLocOfPM = ((Me.Width / Screen.TwipsPerPixelX) - pauseMenu(0).Width) / 2 + 100
' yLocOfPM = ((Me.Height / Screen.TwipsPerPixelY) - pauseMenu(0).Height) / 2 + 123
' 'Shape1.Move XLocOfPM, yLocOfPM, 40, 30
' If X > XLocOfPM And X < XLocOfPM + 40 And Y > yLocOfPM And Y < yLocOfPM + 30 Then
'  Timer1.Enabled = True
'  Timer2.Enabled = True
'  Me.MousePointer = vbCustom
'  IsInPause = False
' End If
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' StopLargeSound "TSBS"
End Sub

Private Sub GhostBuffer_Click()
 End
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = Asc(vbCr) Then
   Timer1.Enabled = True
   Timer2.Enabled = True
   Me.MousePointer = vbCustom
   Text1.Visible = False
   IsInCheatMode = False
   Me.Cls
   
   Target1.Draw Step, XX&, YY&, Me.hdc, Pic.hdc, Msk.hdc
   
   BitBlt Me.hdc, 0, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - Picture6(Statusbar).Height + 5, Picture6(Statusbar).Width, Picture6(Statusbar).Height, Picture7(Statusbar).hdc, 0, 0, SRCAND
   BitBlt Me.hdc, 0, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - Picture6(Statusbar).Height + 5, Picture6(Statusbar).Width, Picture6(Statusbar).Height, Picture6(Statusbar).hdc, 0, 0, SRCINVERT
   
   ScoreVDL.Value = LeadingZeros(Score, 6)
   ScoreVDL.Draw 93, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, CyanDC
   BulletsVDL.Value = LeadingZeros(FreeShots, 2)
   BulletsVDL.Draw 270, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, YellowDC
   TimeVDL.Value = LeadingZeros(Tme, 5)
   TimeVDL.Draw 420, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, RedDC
If ShowRadar Then
   TL = Me.ScaleWidth - TargetRandarView.Width
   TT = 0
   BitBlt Me.hdc, TL, TT, TargetRandarView.Width, TargetRandarView.Height, TargetRandarView.hdc, 0, 0, SRCCOPY
   ';;;;;;;;;;;;;;;;;;;;;;
   'BitBlt Me.hdc, TL + 7, TT + 7, TrgPts.Width, TrgPts.Height, TrgPts.hdc, 0, 0, SRCCOPY
   ';;;;;;;;;;;;;;;;;;;;;;
   ACLL = Mx - XX&
   ACTT = My - YY&
   If ACLL > 0 And ACTT > 0 And ACLL <= TrgPts.Width And ACTT <= TrgPts.Height Then _
      Me.Line (TL + ACLL + 7, TT + ACTT + 7)-(TL + ACLL + 1 + 7, TT + ACTT + 1 + 7), RGB(255, 255, 0)
   If ACL > 0 And ACT > 0 And ACL <= TrgPts.Width And ACT <= TrgPts.Height Then _
      Me.Line (TL + ACL + 7, TT + ACT + 7)-(TL + ACL + 1 + 7, TT + ACT + 1 + 7), RGB(0, 255, 0)
End If
   Me.Refresh
   DoEvents
   Select Case UCase(Mid(Text1.Text, 1, Len(Text1.Text) - 2))
    Case "BULLET STOCK": FreeShots = FreeShots + 25
    Case "PRACTICE MAKES PERFECT"
         Score = 0
         LoadStage 0
         Level = 1
         TrgPts.Picture = TargetNo(0).Picture
         TargetRandarView.Picture = TrgRadar(0).Picture
    Case "DUCK TALE"
         Score = 50
         LoadStage 0
         Level = 2
         TrgPts.Picture = TargetNo(1).Picture
         TargetRandarView.Picture = TrgRadar(1).Picture
    Case "TOTAL WORMAGE"
         Score = 100
         LoadStage 0
         Level = 3
         TrgPts.Picture = TargetNo(2).Picture
         TargetRandarView.Picture = TrgRadar(2).Picture
    Case "BOO"
         Score = 150
         LoadStage 0
         Level = 4
         TrgPts.Picture = TargetNo(3).Picture
         TargetRandarView.Picture = TrgRadar(3).Picture
    Case "UFO"
         Score = 200
         LoadStage 0
         Level = 5
         TrgPts.Picture = TargetNo(4).Picture
         TargetRandarView.Picture = TrgRadar(4).Picture
    Case "AUTO-PILOT": AutoPilot = Not AutoPilot
   End Select
   Text1.Text = ""
 End If
End Sub

Sub Level1_Timer(Step As Long)
  Target1.Draw Step, XX&, YY&, Me.hdc, Pic.hdc, Msk.hdc
  'Target2.Draw Step, TXX&, TYY&, Me.hdc, Pic.hdc, Msk.hdc
  'Move Bottom
  YY& = YY& + RelativeSpeed * 3
  'TYY& = TYY& + RelativeSpeed * 3 + (Rnd * 4)
  
  If YY& > Me.ScaleHeight - 80 Then
   YY& = 0
   XX& = Rnd * 750
  End If
  'If TYY& > Me.ScaleHeight - 80 Then
  ' TYY& = 0
  ' TXX& = Rnd * 750
  'End If

End Sub

Sub Level2_Timer(Step As Long)
  Target1.Draw Step, XX&, YY&, Me.hdc, Pic.hdc, Msk.hdc
  'Move Rigth
  XX& = XX& + RelativeSpeed * 6
  If XX& > Me.ScaleWidth Then
   XX& = -90
   YY& = Rnd * 550
  End If
End Sub

Sub Level3_Timer(Step As Long)
Select Case Step
 Case 0: ST = 19
 Case 1: ST = 18
 Case 2: ST = 19
 Case 3: ST = 16
 Case 4: ST = 20
 Case 5: ST = 22
 Case 6: ST = 21
 Case 7: ST = 25
 Case 8: ST = 24
 Case 9: ST = 23
 Case 10: ST = 26
 Case 11: ST = 29
 Case 12: ST = 28
 Case 13: ST = 27
End Select
Target1.Draw ST, XX&, YY&, Me.hdc, Pic.hdc, Msk.hdc
  YY& = YY& + RelativeSpeed * 3
  If YY& > Me.ScaleHeight - 80 Then
   YY& = 0
   XX& = Rnd * 750
  End If
'19,18,17,16,20,22,21,25,24,23,26,29,28,29
End Sub

Sub Level4_Timer(Step As Long)
'The form is cleared only contains the background
'See Level4 Logic.txt
'Copy the background to the buffer
 GhostBuffer.Cls
 YY& = YY& + (2 * RelativeSpeed)
 If YY& > Me.ScaleHeight - 80 Then
  YY& = 0
  XX& = Rnd * 750
 End If
'Copy the background to the buffer
 BitBlt GhostBuffer.hdc, 0, 0, 25, 24, Me.hdc, XX&, YY&, SRCCOPY
'Blend the Buffrer and the sprite
 Alpha_Blend GhostBuffer.hdc, GhostSpr.hdc, 0, 0, Ghost.FrameLeft(Step), Ghost.FrameTop(Step), 25, 24, 25, 24, 40 '40 is the Alpha 80/255
'Bitblt the Invert mask to the buffer
 GhostBuffer.Refresh
 BitBlt GhostBuffer.hdc, 0, 0, 25, 24, GhostIMask.hdc, Ghost.FrameLeft(Step), Ghost.FrameTop(Step), SRCAND
 GhostBuffer.Refresh
'Bitblt the mask to the form
BitBlt Me.hdc, XX&, YY&, 25, 24, GhostMask.hdc, Ghost.FrameLeft(Step), Ghost.FrameTop(Step), SRCAND
'Bitblt the buffer to the form
BitBlt Me.hdc, XX&, YY&, 25, 24, GhostBuffer.hdc, 0, 0, SRCINVERT
End Sub

Sub Level5_Timer(Step As Long)
  YY& = YY& + ((2 * RelativeSpeed) * PMY)
  XX& = XX& + ((2 * RelativeSpeed) * PMX)
  
  If YY& > Me.ScaleHeight - (54 + (Picture6(Statusbar).Height / 3 * 2)) Then PMY = -1
  If YY& < 10 Then PMY = 1
  If XX& > Me.ScaleWidth - 65 Then PMX = -1
  If XX& < 10 Then PMX = 1
  
  Target1.Draw Step, XX&, YY&, Me.hdc, Pic.hdc, Msk.hdc
End Sub

Sub DrawStatusbar()
'Status bar
BitBlt Me.hdc, 0, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - Picture6(Statusbar).Height + 5, Picture6(Statusbar).Width, Picture6(Statusbar).Height, Picture7(Statusbar).hdc, 0, 0, SRCAND
BitBlt Me.hdc, 0, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - Picture6(Statusbar).Height + 5, Picture6(Statusbar).Width, Picture6(Statusbar).Height, Picture6(Statusbar).hdc, 0, 0, SRCINVERT
'Score LEDs
ScoreVDL.Value = LeadingZeros(Score, 6)
ScoreVDL.Draw 93, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, CyanDC
'Bullets LEDs
BulletsVDL.Value = LeadingZeros(FreeShots, 2)
BulletsVDL.Draw 270, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, YellowDC
'Time LEDs
TimeVDL.Value = LeadingZeros(Tme, 5)
TimeVDL.Draw 420, (Me.Height - Picture6(Statusbar).Height) / Screen.TwipsPerPixelY - 30, NumPic(1), NumMsk(1), Me, RedDC
End Sub

Sub DrawRadar()
   TL = Me.ScaleWidth - TargetRandarView.Width
   TT = 0
   BitBlt Me.hdc, TL, TT, TargetRandarView.Width, TargetRandarView.Height, TargetRandarView.hdc, 0, 0, SRCCOPY
   ';;;;;;;;;;;;;;;;;;;;;;
   'BitBlt Me.hdc, TL + 7, TT + 7, TrgPts.Width, TrgPts.Height, TrgPts.hdc, 0, 0, SRCCOPY
   ';;;;;;;;;;;;;;;;;;;;;;
   ACLL = Mx - XX&
   ACTT = My - YY&
   If ACLL > 0 And ACTT > 0 And ACLL <= TrgPts.Width And ACTT <= TrgPts.Height Then _
      Me.Line (TL + ACLL + 7, TT + ACTT + 7)-(TL + ACLL + 1 + 7, TT + ACTT + 1 + 7), RGB(255, 255, 0)
   If ACL > 0 And ACT > 0 And ACL <= TrgPts.Width And ACT <= TrgPts.Height Then _
      Me.Line (TL + ACL + 7, TT + ACT + 7)-(TL + ACL + 1 + 7, TT + ACT + 1 + 7), RGB(0, 255, 0)
End Sub

Sub Auto_Pilot_Repoint_Mouse()
  Select Case Level
   Case 1: SetCursorPos XX& + 25, YY& + 18
   Case 2: SetCursorPos XX& + 68, YY& + 20
   Case 3: SetCursorPos XX& + 28, YY& + 47
   Case 4: SetCursorPos XX& + 10, YY& + 10
   Case 5: SetCursorPos XX& + 30, YY& + 22
  End Select
End Sub
Private Sub Timer1_Timer()
Static Step As Long
Dim ST As Long
Dim TL, TT As Single
Const SRCAND = &H8800C6
Const SRCINVERT = &H660046
Me.Cls
'If enemy is target
If Level = 1 Then
   Level1_Timer Step
   If Step >= 12 Then Step = 0
End If
'If enemy is duck
If Level = 2 Then
  Level2_Timer Step
  If Step >= 15 Then Step = 12
  If Step < 12 Then Step = 12
End If
'If enemy is worm
If Level = 3 Then
If Step >= 14 Then Step = 0
Level3_Timer Step
End If
'If enemy is ghost
If Level = 4 Then
 If Step >= 15 Then Step = 1
 Level4_Timer Step
End If
'If enemy is UFO
If Level = 5 Then
  Level5_Timer Step
  If Step < 30 Then Step = 30
  If Step >= 36 Then Step = 29
End If
DrawStatusbar
If ShowRadar Then DrawRadar
Me.Refresh
Step = Step + 1
If AutoPilot Then Auto_Pilot_Repoint_Mouse
End Sub

Private Sub Timer2_Timer()
 Tme = Tme - 1
 If Tme = 0 Then Tme = 100
 If MusicEnabled Then
   If PersendagePlayedOfAnOpenedSound("BckSnd") = 100 Then '100% of sound played
      StopLargeSound "BckSnd"
      DoEvents
      RandomSound
   End If
 End If
End Sub

Function LeadingZeros(Value As Variant, Zeros As Long) As String
 Dim OutV As String
 Dim Done As Boolean
 OutV = Trim(Str(Value))
 Done = False
 Do
  If Len(OutV) < Zeros Then
     OutV = "0" & OutV
  Else
     Done = True
  End If
 Loop Until Done
 LeadingZeros = OutV
End Function

