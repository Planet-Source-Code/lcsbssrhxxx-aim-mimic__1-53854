VERSION 5.00
Object = "{F3435EBB-ADE5-4FD0-88FE-05704DAA4FD0}#2.0#0"; "TOCSOCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVBUTTONS.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMimic 
   BackColor       =   &H00E3DFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AIM Mimic V1.03"
   ClientHeight    =   4575
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3390
   Icon            =   "frmMimic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   3390
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab xTab 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14933984
      TabCaption(0)   =   "Log In"
      TabPicture(0)   =   "frmMimic.frx":1272
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fakeIdle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fakeAway"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Interact"
      TabPicture(1)   =   "frmMimic.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "IM"
      Tab(1).Control(1)=   "txtMessage"
      Tab(1).Control(2)=   "EditProfile"
      Tab(1).Control(3)=   "Start"
      Tab(1).Control(4)=   "Warn"
      Tab(1).Control(5)=   "xStop"
      Tab(1).Control(6)=   "Victim"
      Tab(1).Control(7)=   "log"
      Tab(1).Control(8)=   "Image3"
      Tab(1).ControlCount=   9
      Begin LVbuttons.LaVolpeButton IM 
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Send IM"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":12AA
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.TextBox txtMessage 
         Height          =   855
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2760
         Width           =   2895
      End
      Begin LVbuttons.LaVolpeButton EditProfile 
         Height          =   255
         Left            =   -73440
         TabIndex        =   15
         ToolTipText     =   "Stop AIM Mimic"
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Edit Profile"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":12C6
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton Start 
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         ToolTipText     =   "Start AIM Mimic"
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Start"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":12E2
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton Warn 
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         ToolTipText     =   "Warn User"
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Warn"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":12FE
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton xStop 
         Height          =   255
         Left            =   -73440
         TabIndex        =   12
         ToolTipText     =   "Stop AIM Mimic"
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Stop"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":131A
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.TextBox Victim 
         Height          =   285
         Left            =   -74880
         TabIndex        =   10
         ToolTipText     =   "Victim's Screen Name"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox log 
         Height          =   1095
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2895
         Begin VB.ComboBox Server 
            Height          =   315
            ItemData        =   "frmMimic.frx":1336
            Left            =   120
            List            =   "frmMimic.frx":1340
            TabIndex        =   8
            Text            =   "toc.oscar.aol.com"
            ToolTipText     =   "Server"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Image Image2 
            Height          =   210
            Left            =   600
            Picture         =   "frmMimic.frx":1366
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E3DFE0&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
         Begin VB.TextBox Pass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   3
            ToolTipText     =   "Password"
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox SN 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Text            =   "<New User>"
            ToolTipText     =   "Screen Name"
            Top             =   360
            Width           =   2655
         End
         Begin tocSock1.tocSock ts 
            Left            =   0
            Top             =   1800
            _ExtentX        =   2566
            _ExtentY        =   661
         End
         Begin VB.Label GetPass 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Forgot Password?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            MouseIcon       =   "frmMimic.frx":1F78
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   4
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Image Button1 
            Height          =   495
            Left            =   2160
            Picture         =   "frmMimic.frx":20CA
            Top             =   1560
            Width           =   570
         End
         Begin VB.Image Image1 
            Height          =   225
            Left            =   120
            Picture         =   "frmMimic.frx":3000
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label GetSN 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Get a Screen Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   855
         End
         Begin VB.Image button2 
            Height          =   495
            Left            =   2160
            Picture         =   "frmMimic.frx":4212
            Top             =   1560
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin LVbuttons.LaVolpeButton fakeAway 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Fake Away"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":5148
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton fakeIdle 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Fake Idle"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14933984
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmMimic.frx":5164
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMimic.frx":5180
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Status"
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSignOn 
         Caption         =   "Sign On"
      End
      Begin VB.Menu mnuSignOff 
         Caption         =   "Sign Off"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print AIM Convorsation"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuFakeAway 
         Caption         =   "Fake Away"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuIdle 
         Caption         =   "Fake Idle"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Edit Profile"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuWarn 
         Caption         =   "Warn"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSpacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Server"
      End
      Begin VB.Menu mnuChangeSN 
         Caption         =   "Screen Name"
      End
      Begin VB.Menu mnuChangeVSN 
         Caption         =   "Victim's Screen Name"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu toneSend 
         Caption         =   "Play Tone When Sending An IM"
      End
      Begin VB.Menu toneRecive 
         Caption         =   "Play Tone When Reciving An IM"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSpacer8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInet 
         Caption         =   "Internet Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuH_GetSN 
         Caption         =   "Get a Screen Name"
      End
      Begin VB.Menu mnuH_ForgotPass 
         Caption         =   "Forgot Password?"
      End
      Begin VB.Menu mnuH_Problems 
         Caption         =   "Problems Signing On?"
      End
      Begin VB.Menu mnuSpacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuH_Server 
         Caption         =   "Server"
      End
      Begin VB.Menu mnuH_SN 
         Caption         =   "Screen Name / Password"
      End
      Begin VB.Menu mnuSpacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuH_FakeAway 
         Caption         =   "Fake Away"
      End
      Begin VB.Menu mnuH_FakeIdle 
         Caption         =   "Fake Idle"
      End
      Begin VB.Menu mnuSpacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuH_Victim 
         Caption         =   "Victim"
      End
      Begin VB.Menu mnuH_StartStop 
         Caption         =   "Start / Stop"
      End
      Begin VB.Menu mnuH_ChangeProfile 
         Caption         =   "Change Profile"
      End
   End
End
Attribute VB_Name = "frmMimic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'           *************************************************
'           *                 AIM Mimic V1.03               *
'           *  By Mike Plaehn (LCSBSSRHXXX) 5/7/04 - 5/8/04 *
'           *************************************************

'Feel Free To Learn From This, But Please Don't Use Any Of My Code
'With Out My Permission Thank You!
'Enjoi!
' - LCSBSSRHXXX

'If you need TocSock go to:
'http://www.lenshell.com/files.htm

Const FLASHW_STOP = 0 'Stop flashing. The system restores the window to its original state.
Const FLASHW_CAPTION = &H1 'Flash the window caption.
Const FLASHW_TRAY = &H2 'Flash the taskbar button.
Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY) 'Flash both the window caption and taskbar button. This is equivalent to setting the FLASHW_CAPTION Or FLASHW_TRAY flags.
Const FLASHW_TIMER = &H4 'Flash continuously, until the FLASHW_STOP flag is set.
Const FLASHW_TIMERNOFG = &HC 'Flash until the window comes to the foreground.

Private Type FLASHWINFO
    cbSize As Long
    hwnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type

Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Dim Running As Boolean
Dim LoggedOn As Boolean

Dim nRet As Long
Dim sCPL As String

'############################ LOAD / UNLOAD #############################

'[LOAD]

Private Sub Form_Load()
    Randomize
End Sub

'[UNLOAD]

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub toneRecive_Click()
    If toneRecive.Checked = True Then
        toneRecive.Checked = False
    Else
        toneRecive.Checked = True
    End If
End Sub
Private Sub toneSend_Click()
    If toneSend.Checked = True Then
        toneSend.Checked = False
    Else
        toneSend.Checked = True
    End If
End Sub

'################################# TS ###################################

'[SIGNED IN]

Private Sub ts_loggedIn()
    LoggedOn = True
    Status.Caption = "Logged In " & SN.Text
    Start.Enabled = True
    Warn.Enabled = True
    EditProfile.Enabled = True
    fakeAway.Enabled = True
    fakeIdle.Enabled = True
    IM.Enabled = True
    
    mnuStart.Enabled = True
    mnuStop.Enabled = True
    mnuProfile.Enabled = True
    mnuWarn.Enabled = True
    mnuIdle.Enabled = True
    mnuFakeAway.Enabled = True

    mnuSignOff.Enabled = True
    mnuSignOn.Enabled = False
End Sub

'[SIGNED OFF]

Private Sub ts_loggedOut()
    LoggedOn = False
    Status.Caption = "Logged Out"
    Start.Enabled = False
    xStop.Enabled = False
    Warn.Enabled = False
    EditProfile.Enabled = False
    fakeAway.Enabled = False
    fakeIdle.Enabled = False
    IM.Enabled = False

    mnuStart.Enabled = False
    mnuStop.Enabled = False
    mnuProfile.Enabled = False
    mnuWarn.Enabled = False
    mnuIdle.Enabled = False
    mnuFakeAway.Enabled = False
    
    mnuSignOff.Enabled = False
    mnuSignOn.Enabled = True
End Sub

'[IF INCOMING IM]

Private Sub ts_incomingIM(strName As String, strMessage As String, boolAuto As Boolean)
    Dim FlashInfo As FLASHWINFO

    FlashInfo.hwnd = Me.hwnd
    FlashInfo.cbSize = Len(FlashInfo)
    FlashInfo.dwTimeout = 0
    FlashInfo.dwFlags = FLASHW_ALL Or FLASHW_TIMER
    FlashInfo.uCount = 5
    FlashWindowEx FlashInfo

    xTab.Tab = 1
    
    If toneRecive.Checked = True Then
        Beep 1000, 100
        Beep 500, 200
    End If
    
    Victim.Text = strName
    log.Text = log.Text & vbCrLf & strName & " (" & Time & ") : " & ts.removeHTML(strMessage)
    If Running = True Then
        ts.sendIM strName, strMessage, True
        log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & ts.removeHTML(strMessage)
        If toneSend.Checked = True Then
            Beep 500, 100
            Beep 1000, 200
        End If
    End If
End Sub

'[IF INCOMING ERROR]

Private Sub ts_incomingError(intErrorCode As Integer)
    MsgBox "ERROR", vbCritical, "ERROR : " & intErrorCode
End Sub

'[IF INCOMING WARN]

Private Sub ts_incomingWarn(strName As String, intPercent As Integer)
    ts.signOff
End Sub

'############################# MOUSE MOVE ###############################

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    button2.Visible = False
    Button1.Visible = True
End Sub

Private Sub Button1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    button2.Visible = True
    Button1.Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    button2.Visible = False
    Button1.Visible = True
End Sub

'############################# HYPER LINKS ##############################

'[GET A SCREEN NAME]

Private Sub GetSN_Click()
    frmWeb.Show
    frmWeb.WB.Navigate "http://my.screenname.aol.com/_cqr/login/login.psp?siteId=aimregistrationPROD&authLev=1&mcState=initialized&createSn=1&triedAimAuth=y"
End Sub

'[FORGOT PASSWORD?]

Private Sub GetPass_Click()
 frmWeb.Show
    frmWeb.WB.Navigate "http://www.aim.com/help_faq/forgot_password/password.adp"
End Sub

'################################ BUTTONS ###############################

'[SIGN ON]

Private Sub button2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Server.Text = "toc.oscar.aol.com" Then
        ts.loginUser SN.Text, Pass.Text, "gaystuff", "toc.oscar.aol.com", "5190"
    End If
    If Server.Text = "login.icq.com" Then
        ts.loginUser SN.Text, Pass.Text, "gaystuff", "login.icq.com", "5190"
    End If
    Status.Caption = "Logging In " & SN.Text & "..."
    log.Text = ""
End Sub

'[IM]

Private Sub IM_Click()
    If Victim.Text = "" Then
    MsgBox "Please enter the screen name of the victim.", vbCritical, "AIM Mimic"
    Else
        If txtMessage.Text = "" Then txtMessage.Text = " "
        
        ts.sendIM Victim.Text, txtMessage.Text, True
        log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & txtMessage.Text
        txtMessage.Text = ""
    End If
    If toneSend.Checked = True Then
        Beep 500, 100
        Beep 1000, 200
    End If
End Sub

'[START MIMIC]

Private Sub Start_Click()
Dim xRandom As Integer

    If Victim.Text = "" Then
        MsgBox "Please enter the screen name of the victim.", vbCritical, "AIM Mimic"
        Victim.SetFocus
        Exit Sub
    End If
    
    Start.Enabled = False
    xStop.Enabled = True

    Running = True
    xRandom = Int(Rnd * 5)
    
    Select Case xRandom
        Case 0
            ts.sendIM Victim.Text, "hey", True
            log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & "hey"
        Case 1
            ts.sendIM Victim.Text, "hi", True
            log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & "hi"
        Case 2
            ts.sendIM Victim.Text, "hey there", True
            log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & "hey there"
        Case 3
            ts.sendIM Victim.Text, "sup?", True
            log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & "sup?"
        Case 4
            ts.sendIM Victim.Text, "hello", True
            log.Text = log.Text & vbCrLf & SN.Text & " to " & Victim.Text & " (" & Time & ") : " & "hello"
    End Select
    If toneSend.Checked = True Then
        Beep 500, 100
        Beep 1000, 200
    End If
End Sub

'[STOP  MIMIC]

Private Sub xStop_Click()
    Running = False
    Start.Enabled = True
    xStop.Enabled = False
End Sub

'[WARN]

Private Sub Warn_Click()
    frmWarn.Show
End Sub

'[EDIT PROFILE]

Private Sub EditProfile_Click()
    frmProfileEdit.Show
End Sub

'[FAKE AWAY]

Private Sub FakeAway_Click()
    ts.setAway " "
End Sub

'[FAKE IDLE]

Private Sub FakeIdle_Click()
    ts.setIdle (Rnd * 10000)
End Sub

'################################### MENU ###############################

'[FILE]

Private Sub mnuSignOn_Click()
    log.Text = ""
    If Server.Text = "toc.oscar.aol.com" Then
        ts.loginUser SN.Text, Pass.Text, "gaystuff", "toc.oscar.aol.com", "5190"
    End If
    If Server.Text = "login.icq.com" Then
        ts.loginUser SN.Text, Pass.Text, "gaystuff", "login.icq.com", "5190"
    End If
    Status.Caption = "Logging In " & SN.Text & "..."
End Sub
Private Sub mnuSignOff_Click()
    If LoggedOn = True Then
        ts.signOff
    Else
        LoggedOn = False
        MsgBox "Unable to sign off, you are not currently signed on.", vbCritical, "AIM Mimic"
    End If
End Sub
Private Sub mnuPrint_Click()
    CD.ShowPrinter
    Printer.Font = "Times New Roman"
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.Print vbNewLine
    Printer.Print Space(6) & log.Text
    Printer.Print vbNewLine
    Printer.EndDoc
End Sub
Private Sub mnuExit_Click()
    End
End Sub

'[EDIT]

Private Sub mnuFakeAway_Click()
    Call FakeAway_Click
End Sub
Private Sub mnuIdle_Click()
    Call FakeIdle_Click
End Sub
Private Sub mnuStart_Click()
    Call Start_Click
End Sub
Private Sub mnuStop_Click()
    Call xStop_Click
End Sub
Private Sub mnuProfile_Click()
    Call EditProfile_Click
End Sub
Private Sub mnuWarn_Click()
    Call Warn_Click
End Sub
Private Sub mnuServer_Click()
    xTab.Tab = 0
    Server.SetFocus
End Sub
Private Sub mnuChangeSN_Click()
    xTab.Tab = 0
    SN.Text = ""
    SN.SetFocus
End Sub
Private Sub mnuChangeVSN_Click()
    xTab.Tab = 1
    Victim.Text = ""
    Victim.SetFocus
End Sub

'[OPTIONS]

Private Sub mnuInet_Click()
    sCPL = "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
    On Error Resume Next
    nRet = Shell(sCPL, 5)
End Sub

'[HELP]

Private Sub mnuH_ChangeProfile_Click()
    MsgBox "Change Profile:" & vbCrLf & "To change your profile click the button that says" & vbCrLf & "'Edit Profile' while your logged in. Then type in" & vbCrLf & "your desired profile text, and hit apply.", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_FakeAway_Click()
    MsgBox "Fake Away:" & vbCrLf & "Faking away makes it appear that you are away," & vbCrLf & "but it still allows you to send messages with" & vbCrLf & "out returning from the away state.", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_FakeIdle_Click()
    MsgBox "Fake Idle:" & vbCrLf & "Faking idle makes it apper that you are idle," & vbCrLf & "but it still allows you to send messages with" & vbCrLf & "out returning form the idle state.", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_ForgotPass_Click()
    Call GetPass_Click
End Sub
Private Sub mnuH_GetSN_Click()
    Call GetSN_Click
End Sub
Private Sub mnuH_Problems_Click()
    MsgBox "Problems Signing On?" & vbCrLf & "If your having problems signing on try going to tools" & vbCrLf & "then internet options, and delete your temporary" & vbCrLf & "internet files and cookies.", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_Server_Click()
    MsgBox "Server:" & vbCrLf & "The server is how AIM Mimic connects to AIM." & vbCrLf & "The default server is 'toc.oscar.aol.com'", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_SN_Click()
    MsgBox "Screen Name / Password:" & vbCrLf & "Before you can log on you need to type in your AIM Screen Name and Password" & vbCrLf & "If you do not have an AIM Screen Name click on 'Get a Screen Name'", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_StartStop_Click()
    MsgBox "Start / Stop:" & vbCrLf & "The start and stop buttons" & vbCrLf & "are used to turn mimic on or off", vbInformation, "AIM Mimic"
End Sub
Private Sub mnuH_Victim_Click()
    MsgBox "Victim:" & vbCrLf & "The Victim is the Screen Name that you can" & vbCrLf & "send messages to, mimic them, and warn them.", vbInformation, "AIM Mimic"
End Sub

