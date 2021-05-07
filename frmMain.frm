VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   5775
   ClientLeft      =   2640
   ClientTop       =   2325
   ClientWidth     =   5895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdAboutIdoru 
      Height          =   375
      Left            =   2280
      Picture         =   "frmMain.frx":0442
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   3
      ToolTipText     =   "About Idoru Software."
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer tmrCurrent 
      Interval        =   1000
      Left            =   120
      Top             =   5280
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      DownPicture     =   "frmMain.frx":04AD
      Height          =   375
      Left            =   3600
      Picture         =   "frmMain.frx":057D
      TabIndex        =   4
      ToolTipText     =   "Encrypt any type of file."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      DownPicture     =   "frmMain.frx":0654
      Height          =   375
      Left            =   4800
      Picture         =   "frmMain.frx":072B
      TabIndex        =   5
      ToolTipText     =   "Decrypt any type of file."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      DownPicture     =   "frmMain.frx":07FB
      Height          =   375
      Left            =   1080
      Picture         =   "frmMain.frx":08D7
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   2
      ToolTipText     =   "Resets stats."
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompInfo 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":09A5
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   1
      ToolTipText     =   "Get computer information."
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdResetPing 
      DownPicture     =   "frmMain.frx":0CF5
      Height          =   495
      Left            =   2400
      Picture         =   "frmMain.frx":0DD1
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   9
      ToolTipText     =   "Resets stats."
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      DownPicture     =   "frmMain.frx":0E9F
      Height          =   495
      Left            =   4560
      Picture         =   "frmMain.frx":0F56
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   11
      ToolTipText     =   "Quit"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdPing 
      Appearance      =   0  'Ã◊Øƒ
      Default         =   -1  'True
      DownPicture     =   "frmMain.frx":102D
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":1135
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   8
      ToolTipText     =   "Ping the chosen server."
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   495
      Left            =   3720
      Picture         =   "frmMain.frx":1285
      Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
      TabIndex        =   10
      ToolTipText     =   "Help"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame fraComputer 
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtNetData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "<Ethernet Data>"
         ToolTipText     =   "Tells the machine address of the ethernet card."
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtUserData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "<Username Data>"
         ToolTipText     =   "Tells the name of the user currently logged on."
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtCompIPData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "<Comp IP Data>"
         ToolTipText     =   "Tells the IP number of this computer"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtNameData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "<Comp Name Data>"
         ToolTipText     =   "Tells the name of this computer."
         Top             =   600
         Width           =   2175
      End
      Begin MSWinsockLib.Winsock socLocal 
         Left            =   1200
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtOSData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "<OS Data>"
         ToolTipText     =   "Tells this computer's operating system."
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblNet 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Ethernet Address:"
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
         TabIndex        =   44
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblOS 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Operating System:"
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
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Username:"
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
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblCompIP 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Computer IP:"
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
         TabIndex        =   41
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Computer Name:"
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
         TabIndex        =   40
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame fraHacker 
      Caption         =   "PW Hacker"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   4080
      TabIndex        =   35
      Top             =   0
      Width           =   1815
      Begin VB.Timer tmrHack 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   1200
      End
      Begin VB.CommandButton cmdHackHelp 
         Height          =   375
         Left            =   960
         Picture         =   "frmMain.frx":12F0
         Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
         TabIndex        =   7
         ToolTipText     =   "Instructions"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdReady 
         DownPicture     =   "frmMain.frx":135B
         Height          =   375
         Left            =   120
         Picture         =   "frmMain.frx":1403
         Style           =   1  '∏ﬁ◊Ã®Ø∏Ω
         TabIndex        =   6
         ToolTipText     =   "Click to initiate Password Hacker."
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "5"
         ToolTipText     =   "Amount of time."
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00000000&
         BorderStyle     =   1  'é¿ê¸
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Status"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblDelay 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Delay -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraServer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   5895
      Begin VB.TextBox txtEnterServer 
         Height          =   285
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   34
         Text            =   "<name of server>"
         ToolTipText     =   "Type the name of the server you want to ping here."
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox txtEnterIp 
         Height          =   285
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   29
         Text            =   "<ip number>"
         ToolTipText     =   "Type the IP address you want to ping here."
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtEchoData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "<echo status data>"
         ToolTipText     =   "Echo status."
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtIPData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "<ip address data>"
         ToolTipText     =   "The server's IP number."
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox txtServerData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "<server name data>"
         ToolTipText     =   "The name of the server that was entered."
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtSizeData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "<size data>"
         ToolTipText     =   "The size in bytes."
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtPortData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "<port address data>"
         ToolTipText     =   "The lag in milliseconds."
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtLagData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "<lag time data>"
         ToolTipText     =   "Tells the lag time."
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtStatusData 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "<status data>"
         ToolTipText     =   "Tells if the ping was successfull."
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblEnterServer 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Name of Server:"
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
         TabIndex        =   33
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblServerName 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Server Name:"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblPing 
         BackStyle       =   0  'ìßñæ
         Caption         =   "Ping IP Number:"
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
         TabIndex        =   30
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblLag 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Lag Time:"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Port Address:"
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
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Size:"
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
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblServerStatus 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Status:"
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
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblEcho 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "Echo:"
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
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'ìßñæ
         Caption         =   "IP Address:"
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
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMain.frm

'*************** All Declarations Begin Here **************

Option Explicit

Private Const NCBASTAT = &H33
Private Const NCBNAMSZ = 16
Private Const HEAP_ZERO_MEMORY = &H8
Private Const HEAP_GENERATE_EXCEPTIONS = &H4
Private Const NCBRESET = &H32

Private Type NCB
ncb_command          As Byte
ncb_retcode          As Byte
ncb_lsn              As Byte
ncb_num              As Byte
ncb_buffer           As Long
ncb_length           As Integer
ncb_callname         As String * NCBNAMSZ
ncb_name             As String * NCBNAMSZ
ncb_rto              As Byte
ncb_sto              As Byte
ncb_post             As Long
ncb_lana_num         As Byte
ncb_cmd_cplt         As Byte
ncb_reserve(9)       As Byte
ncb_event            As Long
End Type

Private Type ADAPTER_STATUS
adapter_address(5)   As Byte
rev_major            As Byte
reserved0            As Byte
adapter_type         As Byte
rev_minor            As Byte
duration             As Integer
frmr_recv            As Integer
frmr_xmit            As Integer
iframe_recv_err      As Integer
xmit_aborts          As Integer
xmit_success         As Long
recv_success         As Long
iframe_xmit_err      As Integer
recv_buff_unavail    As Integer
t1_timeouts          As Integer
ti_timeouts          As Integer
Reserved1            As Long
free_ncbs            As Integer
max_cfg_ncbs         As Integer
max_ncbs             As Integer
xmit_buf_unavail     As Integer
max_dgram_size       As Integer
pending_sess         As Integer
max_cfg_sess         As Integer
max_sess             As Integer
max_sess_pkt_size    As Integer
name_count           As Integer
End Type

Private Type NAME_BUFFER
name                 As String * NCBNAMSZ
name_num             As Integer
name_flags           As Integer
End Type

Private Type ASTAT
adapt                As ADAPTER_STATUS
NameBuff(30)         As NAME_BUFFER
End Type


Private Declare Function Netbios Lib "netapi32.dll" (pncb As NCB) As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As _
Long


Private Type OPENFILENAME
lStructSize         As Long
hwndOwner           As Long
hInstance           As Long
lpstrFilter         As String
lpstrCustomFilter   As String
nMaxCustFilter      As Long
nFilterIndex        As Long
lpstrFile           As String
nMaxFile            As Long
lpstrFileTitle      As String
nMaxFileTitle       As Long
lpstrInitialDir     As String
lpstrTitle          As String
Flags               As Long
nFileOffset         As Integer
nFileExtension      As Integer
lpstrDefExt         As String
lCustData           As Long
lpfnHook            As Long
lpTemplateName      As String
End Type


Dim process         As Boolean

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32" _
(ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias _
"GetClassNameA" (ByVal hwnd As Long, ByVal _
lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
X As Long
Y As Long
End Type

Private Const EM_SETPASSWORDCHAR = &HCC

Private Sub ShowPassword()

Dim curWindow   As Long
Dim sClassName  As String * 255
Dim sName       As String
Dim lpPoint     As POINTAPI

Call GetCursorPos(lpPoint)
curWindow = WindowFromPoint(lpPoint.X, lpPoint.Y)

Call GetClassName(curWindow, sClassName, 255)
sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))

If sName = "Edit" Or InStr(sName, "TextBox") > 0 Then
Call SendMessage(curWindow, EM_SETPASSWORDCHAR, 0, 0)
lblStatus.BackColor = vbGreen
lblStatus.Caption = "Ready"
Else
MsgBox "Sorry! Cannot get the password", vbExclamation, "Error !"
End If

tmrHack.Enabled = False

End Sub

Function getfile(Filter As String, title As String, Handle As Long, Flags As Long) As String
 
Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hwndOwner = Handle
ofn.hInstance = App.hInstance
ofn.lpstrFilter = Filter
ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = App.Path
ofn.lpstrTitle = title
ofn.Flags = Flags
                               
Dim a
a = GetOpenFileName(ofn)

If (a) Then
getfile = Trim$(ofn.lpstrFile)
                                
Else
getfile = "Cancel"
End If

End Function
'***************  All Declarations End Here ***************

'***************** All Commands Begin Here ****************

'Shows New Features About Ping Monitor
Private Sub cmdAbout_Click()
Dim strMessage As String
strMessage = " Sorry, help now is not available. We ( Hoang Lam, Nguyen Duc, Yen Thai, Le Minh) will complete is as soon as possible"
MsgBox strMessage, vbOKOnly + vbInformation, "Ping Pro Help"
End Sub

'Shows the About Idoru Form
Private Sub cmdAboutIdoru_Click()
frmAbout.Visible = True
End Sub

'Shows Password Hacker Help
Private Sub cmdHackHelp_Click()

Dim strMessage As String
strMessage = "Set time delay (in seconds). Then press Ready button. " _
& vbCrLf & "After that place the mouse pointer on password box" _
& vbCrLf & "of any application." _
& vbCrLf & "After the specified time, the status message will tell " _
& vbCrLf & "you that you're now ready to get the password." _
& vbCrLf & "Then click that password box, and" _
& vbCrLf & "it will apear in normal text." _
& vbCrLf & "(Does not work on some networked password boxes)"
MsgBox strMessage, vbOKOnly + vbInformation, "Password Hacker Help"
End Sub

'Pings the Entered IP Number
Private Sub cmdPing_Click()

Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Integer
  
Call Ping(txtEnterIp.Text, ECHO)
txtStatusData.Text = GetStatusCode(ECHO.status)
txtLagData.Text = ECHO.RoundTripTime & " (milliseconds)"
txtPortData.Text = ECHO.Address
txtSizeData.Text = ECHO.DataSize & " (bytes)"
txtEchoData.Text = ECHO.Data & " (Data)"
txtIPData.Text = txtEnterIp.Text
txtServerData.Text = txtEnterServer.Text

If Left$(ECHO.Data, 1) <> Chr$(0) Then
pos = InStr(ECHO.Data, Chr$(0))
txtEchoData.Text = Left$(ECHO.Data, pos - 1)
End If

'An Easter Egg!
If txtEnterIp.Text = "fhec" Then
Dim strMessage As String
strMessage = "The company name is dead, but the spirit lives on!!"
MsgBox strMessage, vbOKOnly + vbInformation, "Viva Fish Head Eates of Cambodia!!!"
End If
End Sub

'Tells Hacker to Get Password
Private Sub cmdReady_Click()
tmrHack.Interval = Val(txtDelay.Text) * 1000
tmrHack.Enabled = True
lblStatus.BackColor = vbRed
lblStatus.Caption = "Wait"
End Sub

'Resets all Ping Text Boxes
Private Sub cmdResetPing_Click()
txtStatusData.Text = ""
txtLagData.Text = ""
txtPortData.Text = ""
txtSizeData.Text = ""
txtServerData.Text = ""
txtIPData.Text = ""
txtEchoData.Text = ""
End Sub

'Gets Information About the Computer
Private Sub cmdCompInfo_Click()
txtOSData.Text = GetWindowsVersion()
txtNameData.Text = modComputerName.ComputerName()
txtCompIPData.Text = socLocal.LocalIP
txtUserData.Text = modUserName.UserName()
txtNetData.Text = EthernetAddress(0)
If txtNetData.Text = "000000000000" Then
txtNetData.Text = "[No Ethernet Card Detected]"
End If
End Sub

'Clears Computer Informaion
Private Sub cmdRefresh_Click()
txtOSData.Text = ""
txtNameData.Text = ""
txtCompIPData.Text = ""
txtUserData.Text = ""
txtNetData.Text = ""
End Sub

'Decrypts any Type of File
Private Sub cmdDecrypt_Click()

Dim file2open As String
Dim file2save As String
Dim tempstr As String * 1
Dim result As Long
Dim fref1 As Long
Dim fref2 As Long
Dim starttime As Long
Dim endtime As Long

file2open = getfile("All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "Select file to Decrypt ...", Me.hwnd, &H1000)

If file2open = "Cancel" Then Exit Sub

file2save = getfile("All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "Save Decrypted file as ...", Me.hwnd, 0)

If file2save = "Cancel" Then
result = MsgBox("Cancel the Decryption Process ?", vbQuestion + vbYesNo)

If result = vbYes Then
Exit Sub

ElseIf result = vbNo Then
cmdDecrypt_Click

End If

Else
Me.WindowState = vbMinimized
starttime = GetTickCount
process = True

Me.Caption = "Decrypting " & file2open & " ..."

fref1 = FreeFile

Open file2open For Binary As #fref1

fref2 = FreeFile

Open file2save For Binary As #fref2

For result = 1 To LOF(fref1)

Get #fref1, , tempstr
Put #fref2, , encdec(tempstr)

Next result

Close #fref2
Close #fref1

endtime = GetTickCount
process = False

MsgBox "Done Decryption.Decryption process took " & endtime - starttime & " milliseconds", vbInformation

Me.WindowState = vbNormal
Me.Caption = "Hoang Lam"

End If
End Sub

'Encyrpts any Type of File
Private Sub cmdEncrypt_Click()
Dim file2open   As String
Dim file2save   As String
Dim result      As Long
Dim fref1       As Long
Dim fref2       As Long
Dim tempstr     As String * 1
Dim starttime   As Long
Dim endtime     As Long

file2open = getfile("All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "Select file to Encrypt ...", Me.hwnd, &H1000)

If file2open = "Cancel" Then Exit Sub


file2save = getfile("All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0), "Save Encrypted file as ...", Me.hwnd, 0)

If file2save = "Cancel" Then
result = MsgBox("Cancel the Encryption Process ?", vbQuestion + vbYesNo)

If result = vbYes Then
Exit Sub
ElseIf result = vbNo Then
cmdEncrypt_Click
End If

Else
Me.WindowState = vbMinimized
starttime = GetTickCount
process = True
Me.Caption = "Encrypting " & file2open & " ..."

fref1 = FreeFile

Open file2open For Binary As #fref1

fref2 = FreeFile

Open file2save For Binary As #fref2

For result = 1 To LOF(fref1)

Get #fref1, , tempstr
Put #fref2, , encdec(tempstr)

Next result

Close #fref2
Close #fref1

Me.WindowState = vbMinimized
process = False
endtime = GetTickCount

MsgBox "Done Encryption. Enccyption process took " & endtime - starttime & " milliseconds", vbInformation

Me.WindowState = vbNormal
Me.Caption = "Idoru Encryptor"
End If
End Sub

'The Exit Command
Private Sub cmdExit_Click()
Unload Me
End
End Sub
'******************* All Commands End Here ****************


'Loads Title Bar Caption and Current Time
Private Sub tmrCurrent_Timer()
frmMain.Caption = "Ping Monitor Pro             Current Time: " & Time
End Sub

Private Sub Form_Load()

'Loads Computer Info
txtOSData.Text = GetWindowsVersion()
txtNameData.Text = modComputerName.ComputerName()
txtCompIPData.Text = socLocal.LocalIP
txtUserData.Text = modUserName.UserName()
txtNetData.Text = EthernetAddress(0)

If txtNetData.Text = "000000000000" Then
txtNetData.Text = "[No Ethernet Card Detected]"
End If

'Clears All Text Boxes of Notes
txtStatusData.Text = ""
txtLagData.Text = ""
txtPortData.Text = ""
txtSizeData.Text = ""
txtServerData.Text = ""
txtIPData.Text = ""
txtEchoData.Text = ""
txtServerData.Text = ""
txtEnterIp.Text = "255.255.255.255"
txtEnterServer.Text = "HanoiNTC"

'Tells Caption and Color to the PW Hacker Label
lblStatus.BackColor = vbRed
lblStatus.Caption = "Wait"
End Sub

'Calls the ShowPassword Declaration
Private Sub tmrHack_Timer()
Call ShowPassword
End Sub

'Status Label for Password Hacker (Error Messages)
Private Sub lblStatus_Click()
If Not IsNumeric(txtDelay.Text) Then
MsgBox "Please Enter a specified amount of time...", vbInformation, "Error !"
txtDelay = 5
End If
End Sub

'Properties and Error Messages for txtDelay
Private Sub txtDelay_LostFocus()

If txtDelay.Text = "00" Then
MsgBox "Please Enter a specified amount of time...", vbInformation, "Error !"
txtDelay = 5
End If

If txtDelay.Text = "0" Then
MsgBox "Please Enter a specified amount of time...", vbInformation, "Error !"
txtDelay = 5
End If

If Not IsNumeric(txtDelay.Text) Then
MsgBox "Please Enter a specified amount of time...", vbInformation, "Error !"
txtDelay = 5
End If
End Sub

Private Function EthernetAddress(LanaNumber As Long) As String
  
Dim udtNCB       As NCB
Dim bytResponse  As Byte
Dim udtASTAT     As ASTAT
Dim udtTempASTAT As ASTAT
Dim lngASTAT     As Long
Dim strOut       As String
Dim X            As Integer

udtNCB.ncb_command = NCBRESET
bytResponse = Netbios(udtNCB)
udtNCB.ncb_command = NCBASTAT
udtNCB.ncb_lana_num = LanaNumber
udtNCB.ncb_callname = "* "
udtNCB.ncb_length = Len(udtASTAT)
lngASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, udtNCB.ncb_length)
strOut = ""

If lngASTAT Then
udtNCB.ncb_buffer = lngASTAT
bytResponse = Netbios(udtNCB)
CopyMemory udtASTAT, udtNCB.ncb_buffer, Len(udtASTAT)
With udtASTAT.adapt
For X = 0 To 5
strOut = strOut & Right$("00" & Hex$(.adapter_address(X)), 2)
Next X
End With
HeapFree GetProcessHeap(), 0, lngASTAT
End If

EthernetAddress = strOut
End Function
