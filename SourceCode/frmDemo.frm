VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDemo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   """GROG PRO"" ACCESS"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10215
   Visible         =   0   'False
   Begin VB.Timer tmrRelay 
      Enabled         =   0   'False
      Left            =   8160
      Top             =   3240
   End
   Begin VB.Timer tmrPasswTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   960
      Top             =   1080
   End
   Begin VB.CheckBox chkDummy 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   9720
      TabIndex        =   32
      Top             =   7200
      Width           =   255
   End
   Begin VB.Frame fraFlag 
      Caption         =   "Language"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   6360
      TabIndex        =   28
      Tag             =   "0"
      Top             =   1320
      Width           =   3732
      Begin VB.OptionButton optEnglish 
         Caption         =   "English"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Language"
         Top             =   240
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton optLatvian 
         Caption         =   "Latvian"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1440
         TabIndex        =   30
         ToolTipText     =   "Language"
         Top             =   240
         Width           =   1212
      End
      Begin VB.OptionButton optRussian 
         Caption         =   "Russian"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2640
         TabIndex        =   29
         ToolTipText     =   "Language"
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.Data datBase 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   516
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.PictureBox picTools 
      Height          =   612
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   9915
      TabIndex        =   26
      ToolTipText     =   "Tools"
      Top             =   120
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Image imgAccessServ 
         Height          =   375
         Left            =   9480
         Picture         =   "frmDemo.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "AccessService-Correction"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgAccessInfo 
         Height          =   375
         Left            =   9000
         Picture         =   "frmDemo.frx":03C2
         Stretch         =   -1  'True
         ToolTipText     =   "AccessInfo"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgAccessOut 
         Height          =   375
         Left            =   8520
         Picture         =   "frmDemo.frx":079C
         Stretch         =   -1  'True
         ToolTipText     =   "AccessOutputData"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgAccessIn 
         Height          =   375
         Left            =   8040
         Picture         =   "frmDemo.frx":0BE2
         Stretch         =   -1  'True
         ToolTipText     =   "AccessInputData"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line19 
         X1              =   7920
         X2              =   7920
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Image imgPreprocessors 
         Height          =   375
         Left            =   5400
         Picture         =   "frmDemo.frx":1028
         Stretch         =   -1  'True
         ToolTipText     =   "Preprocessors"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgParkingServ 
         Height          =   375
         Left            =   7440
         Picture         =   "frmDemo.frx":1892
         Stretch         =   -1  'True
         ToolTipText     =   "ParkingService-Correction"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgParkingInfo 
         Height          =   375
         Left            =   6960
         Picture         =   "frmDemo.frx":1C54
         Stretch         =   -1  'True
         ToolTipText     =   "ParkingInfo"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line18 
         X1              =   5880
         X2              =   5880
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Image imgParkingOut 
         Height          =   375
         Left            =   6480
         Picture         =   "frmDemo.frx":202E
         Stretch         =   -1  'True
         ToolTipText     =   "ParkingOutputData"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgParkingIn 
         Height          =   375
         Left            =   6000
         Picture         =   "frmDemo.frx":2240
         Stretch         =   -1  'True
         ToolTipText     =   "ParkingInputData"
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgBookKeeperBase 
         Height          =   375
         Left            =   720
         Picture         =   "frmDemo.frx":2452
         Stretch         =   -1  'True
         ToolTipText     =   "Form BookKeeper Base"
         Top             =   120
         Width           =   375
      End
      Begin VB.Line Line17 
         X1              =   2760
         X2              =   2760
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Image imgProtocolBase 
         Height          =   375
         Left            =   1200
         Picture         =   "frmDemo.frx":27F4
         Stretch         =   -1  'True
         ToolTipText     =   "Form Protocol Base"
         Top             =   120
         Width           =   375
      End
      Begin VB.Line Line16 
         X1              =   600
         X2              =   600
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Image imgSaveProtocol 
         Height          =   375
         Left            =   2280
         Picture         =   "frmDemo.frx":3076
         Stretch         =   -1  'True
         ToolTipText     =   "Save Protocol"
         Top             =   120
         Width           =   375
      End
      Begin VB.Line Line15 
         X1              =   5280
         X2              =   5280
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line14 
         X1              =   2160
         X2              =   2160
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Image imgTerminal 
         Height          =   375
         Left            =   3840
         Picture         =   "frmDemo.frx":3D78
         Stretch         =   -1  'True
         ToolTipText     =   "Terminal"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgTime 
         Height          =   375
         Left            =   3360
         Picture         =   "frmDemo.frx":3E7A
         Stretch         =   -1  'True
         ToolTipText     =   "Time"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgPrint 
         Height          =   375
         Left            =   120
         Picture         =   "frmDemo.frx":5B1C
         Stretch         =   -1  'True
         ToolTipText     =   "Print"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgPersons 
         Height          =   375
         Left            =   2880
         Picture         =   "frmDemo.frx":5EBE
         Stretch         =   -1  'True
         ToolTipText     =   "Persons"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgProtocArchives 
         Height          =   375
         Left            =   1680
         Picture         =   "frmDemo.frx":5F88
         Stretch         =   -1  'True
         ToolTipText     =   "Protocol to Archives"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgCalendar 
         Height          =   375
         Left            =   4320
         Picture         =   "frmDemo.frx":608A
         Stretch         =   -1  'True
         ToolTipText     =   "Calendar"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSystem 
         Height          =   375
         Left            =   4800
         Picture         =   "frmDemo.frx":688C
         Stretch         =   -1  'True
         ToolTipText     =   "System"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Timer tmrButton 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   10000
      Left            =   8880
      Tag             =   "0"
      Top             =   7560
   End
   Begin VB.Timer tmrButton 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   10000
      Left            =   7200
      Tag             =   "0"
      Top             =   7560
   End
   Begin VB.Timer tmrButton 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10000
      Left            =   5400
      Tag             =   "0"
      Top             =   7560
   End
   Begin VB.Timer tmrButton 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10000
      Left            =   3480
      Tag             =   "0"
      Top             =   7560
   End
   Begin VB.Timer tmrTermContr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1320
      Top             =   2880
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   8000
      Left            =   8880
      Top             =   6720
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   8000
      Left            =   7200
      Top             =   6720
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   7000
      Left            =   5400
      Top             =   6720
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   6000
      Left            =   3480
      Top             =   6720
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   21
      Tag             =   """"""
      ToolTipText     =   "For Setup and Exit "
      Top             =   804
      Width           =   972
   End
   Begin VB.CheckBox chkPhoto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "Car photo"
      Top             =   3840
      Width           =   252
   End
   Begin VB.CheckBox chkPhoto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   1920
      TabIndex        =   19
      ToolTipText     =   "Car photo"
      Top             =   3240
      Width           =   252
   End
   Begin VB.CheckBox chkPhoto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   18
      ToolTipText     =   "Person photo"
      Top             =   2520
      Width           =   252
   End
   Begin VB.CheckBox chkPhoto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   1920
      TabIndex        =   17
      ToolTipText     =   "Person photo"
      Top             =   1920
      Width           =   252
   End
   Begin VB.Frame fraControl 
      Caption         =   "Control"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   360
      TabIndex        =   16
      Top             =   6840
      Width           =   1455
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton optAutomatic 
         Caption         =   "Automatic"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   720
      TabIndex        =   15
      ToolTipText     =   "End"
      Top             =   5040
      Width           =   972
   End
   Begin VB.CheckBox chkSetup 
      Caption         =   "Execute/Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   14
      ToolTipText     =   "Settings"
      Top             =   840
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "N_3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   8640
      TabIndex        =   13
      Tag             =   "0"
      ToolTipText     =   "Permission"
      Top             =   7200
      Width           =   972
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "N_2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   6840
      TabIndex        =   12
      Tag             =   "0"
      ToolTipText     =   "Permission"
      Top             =   7200
      Width           =   972
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "N_1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Tag             =   "0"
      ToolTipText     =   "Permission"
      Top             =   7200
      Width           =   972
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "N_0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Tag             =   "0"
      ToolTipText     =   "Permission"
      Top             =   7200
      Width           =   972
   End
   Begin VB.CheckBox chkTerm 
      Caption         =   "N_3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   3
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Gates control"
      Top             =   3840
      Width           =   732
   End
   Begin VB.CheckBox chkTerm 
      Caption         =   "N_2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Barriers control"
      Top             =   3240
      Width           =   732
   End
   Begin VB.CheckBox chkTerm 
      Caption         =   "N_1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Tourniquets control"
      Top             =   2520
      Width           =   732
   End
   Begin VB.CheckBox chkTerm 
      Caption         =   "N_0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Doors control"
      Top             =   1920
      Width           =   732
   End
   Begin MSCommLib.MSComm prtPortC 
      Index           =   0
      Left            =   2760
      Tag             =   "0"
      Top             =   7200
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin MSCommLib.MSComm prtPortC 
      Index           =   1
      Left            =   4560
      Tag             =   "0"
      Top             =   7200
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin MSCommLib.MSComm prtPortC 
      Index           =   2
      Left            =   6360
      Tag             =   "0"
      Top             =   7200
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin MSCommLib.MSComm prtPortC 
      Index           =   3
      Left            =   8160
      Tag             =   "0"
      Top             =   7200
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
   End
   Begin MSCommLib.MSComm prtPortDocument 
      Left            =   960
      Tag             =   "0"
      Top             =   6120
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin MSCommLib.MSComm prtPortBarCode 
      Left            =   240
      Tag             =   "0"
      Top             =   6120
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   1024
   End
   Begin MSCommLib.MSComm prtPortDisplay 
      Left            =   1680
      Tag             =   "0"
      Top             =   6120
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.Label lblErrorInpOut 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Input/Output ! ! !"
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   38
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblErrorInpOut 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Input/Output ! ! !"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   37
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblErrorInpOut 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Input/Output ! ! !"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMessageInput 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   35
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblErrorBarCodePrinter 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "BarCode Printer Error ! ! !  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   34
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgEmployeInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3120
      Picture         =   "frmDemo.frx":6C2E
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Employe ""i"""
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgEmployeOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3720
      Picture         =   "frmDemo.frx":7038
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Employe ""--"""
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgEmployeInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2520
      Picture         =   "frmDemo.frx":7442
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Employe ""+"""
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgAccessInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   9600
      Picture         =   "frmDemo.frx":784C
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""Info_?"""
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   7080
      Picture         =   "frmDemo.frx":7C26
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""Info_?"""
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   5280
      Picture         =   "frmDemo.frx":8000
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""Info_?"""
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   9120
      Picture         =   "frmDemo.frx":83DA
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""--""==>"
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   7560
      Picture         =   "frmDemo.frx":8820
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""--""==>"
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   5760
      Picture         =   "frmDemo.frx":8C66
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""--""==>"
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "frmDemo.frx":90AC
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""+""<=="
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   6600
      Picture         =   "frmDemo.frx":94F2
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""+""<=="
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   4800
      Picture         =   "frmDemo.frx":9938
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""+""<=="
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3480
      Picture         =   "frmDemo.frx":9D7E
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""Info_?"""
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3960
      Picture         =   "frmDemo.frx":A158
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""--""==>"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgAccessInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "frmDemo.frx":A59E
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Access ""+""<=="
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblErrorInpOut 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Input/Output ! ! !"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgParkingInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   6480
      Picture         =   "frmDemo.frx":A9E4
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""+""<=="
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   27
      Tag             =   "24"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image imgParkingInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   9120
      Picture         =   "frmDemo.frx":ABF6
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""Info_?"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   5280
      Picture         =   "frmDemo.frx":AFD0
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""Info_?"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3480
      Picture         =   "frmDemo.frx":B3AA
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""Info_?"""
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   9600
      Picture         =   "frmDemo.frx":B784
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""--""==>"
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   5760
      Picture         =   "frmDemo.frx":B996
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""--""==>"
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgParkingOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3960
      Picture         =   "frmDemo.frx":BBA8
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""--""==>"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   8640
      Picture         =   "frmDemo.frx":BDBA
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""+""<=="
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   4680
      Picture         =   "frmDemo.frx":BFCC
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""+""<=="
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgParkingInData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   3000
      Picture         =   "frmDemo.frx":C1DE
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""+""<=="
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingInfoData 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   7080
      Picture         =   "frmDemo.frx":C3F0
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""Info_?"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgParkingOutData 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   7560
      Picture         =   "frmDemo.frx":C7CA
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Parking ""--""==>"
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgViewClose 
      Height          =   1455
      Index           =   3
      Left            =   8400
      Picture         =   "frmDemo.frx":C9DC
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgViewClose 
      Height          =   1452
      Index           =   2
      Left            =   6600
      Picture         =   "frmDemo.frx":1E66A
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Image imgViewClose 
      Height          =   1452
      Index           =   1
      Left            =   4680
      Picture         =   "frmDemo.frx":339FC
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Image imgViewClose 
      Height          =   1452
      Index           =   0
      Left            =   3120
      Picture         =   "frmDemo.frx":48D8E
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label lblOpen 
      Caption         =   "Control buttons"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   25
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      Caption         =   "Password"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1212
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   10080
      X2              =   10080
      Y1              =   6720
      Y2              =   8040
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   6720
      Y2              =   8040
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   120
      X2              =   10080
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   120
      X2              =   10080
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   10080
      X2              =   10080
      Y1              =   2160
      Y2              =   6600
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   2760
      X2              =   10080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   2160
      Y2              =   6600
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2760
      X2              =   10080
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Image imgViewOpen 
      Height          =   1455
      Index           =   3
      Left            =   8280
      Picture         =   "frmDemo.frx":54E64
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgViewOpen 
      Height          =   1452
      Index           =   2
      Left            =   6480
      Picture         =   "frmDemo.frx":66AF2
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      Caption         =   "#####"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   8280
      TabIndex        =   11
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      Caption         =   "#####"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   6480
      TabIndex        =   10
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Image imgPhoto 
      Height          =   1935
      Index           =   3
      Left            =   8280
      Picture         =   "frmDemo.frx":7C014
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image imgPhoto 
      Height          =   1935
      Index           =   2
      Left            =   6480
      Picture         =   "frmDemo.frx":940E6
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image imgPhoto 
      Height          =   1935
      Index           =   1
      Left            =   4680
      Picture         =   "frmDemo.frx":AC1B8
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      Caption         =   "#####"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Image imgViewOpen 
      Height          =   1452
      Index           =   1
      Left            =   4560
      Picture         =   "frmDemo.frx":C428A
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Image imgPhoto 
      Height          =   2052
      Index           =   0
      Left            =   2880
      Picture         =   "frmDemo.frx":D97AC
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label lblInform 
      Alignment       =   2  'Center
      Caption         =   "#####"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Top             =   4560
      Width           =   1452
   End
   Begin VB.Image imgViewOpen 
      Height          =   1452
      Index           =   0
      Left            =   3000
      Picture         =   "frmDemo.frx":F187E
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   2400
      X2              =   2520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblPhoto 
      Alignment       =   2  'Center
      Caption         =   "Photo"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   612
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   1680
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   2520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1680
      Y2              =   4560
   End
   Begin VB.Label lblTerminals 
      Alignment       =   2  'Center
      Caption         =   "Terminals"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   972
   End
   Begin VB.Image imgEnglish 
      Height          =   252
      Left            =   6480
      Picture         =   "frmDemo.frx":FD954
      Stretch         =   -1  'True
      Top             =   840
      Width           =   492
   End
   Begin VB.Image imgRussian 
      Height          =   252
      Left            =   9000
      Picture         =   "frmDemo.frx":FFE46
      Stretch         =   -1  'True
      Top             =   840
      Width           =   492
   End
   Begin VB.Image imgLatvian 
      Height          =   252
      Left            =   7800
      Picture         =   "frmDemo.frx":102288
      Stretch         =   -1  'True
      Top             =   840
      Width           =   492
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Preview..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormBookKeeperBase 
         Caption         =   "Form BookKeeper Base"
      End
      Begin VB.Menu mnuFormProtocolBase 
         Caption         =   "Form Protocol Base"
      End
      Begin VB.Menu mnuProtocolToArchives 
         Caption         =   "Protocol to Archives..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAdjustment 
      Caption         =   "Adjustment"
      Begin VB.Menu mnuSaveProtocol 
         Caption         =   "Save Protocol"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSaveProtocolAs 
         Caption         =   "Save Protocol As..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystem 
         Caption         =   "System"
      End
      Begin VB.Menu mnuPersons 
         Caption         =   "Persons"
      End
      Begin VB.Menu mnuTime 
         Caption         =   "Time"
      End
      Begin VB.Menu mnuTerminal 
         Caption         =   "Terminal"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "Calendar"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreprocessors 
         Caption         =   "Preprocessors..."
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessagesEditor 
         Caption         =   "MessagesEditor..."
      End
   End
   Begin VB.Menu mnuParking 
      Caption         =   "Parking"
      Begin VB.Menu mnuParkingInData 
         Caption         =   "ParkiingInData..."
      End
      Begin VB.Menu mnuParkingOutData 
         Caption         =   "ParkingOutData..."
      End
      Begin VB.Menu mnuParkingInfoData 
         Caption         =   "ParkingInfoData..."
      End
      Begin VB.Menu mnuParkingServData 
         Caption         =   "ParkingServData..."
      End
   End
   Begin VB.Menu mnuAccess 
      Caption         =   "Access"
      Begin VB.Menu mnuAccessInData 
         Caption         =   "AccessInData..."
      End
      Begin VB.Menu mnuAccessOutData 
         Caption         =   "AccessOutData..."
      End
      Begin VB.Menu mnuAccessInfoData 
         Caption         =   "AccessInfoData..."
      End
      Begin VB.Menu mnuAccessservData 
         Caption         =   "AccessServData..."
      End
   End
   Begin VB.Menu mnuEmploye 
      Caption         =   "Employe"
      Begin VB.Menu mnuEmployeInData 
         Caption         =   "EmployeInData..."
      End
      Begin VB.Menu mnuEmployeOutData 
         Caption         =   "EmployeOutData..."
      End
      Begin VB.Menu mnuEmployeInfoData 
         Caption         =   "EmployeInfoData..."
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Переменная-объект СОБЫТИЕ для
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
Dim WithEvents qEvent As MSMQEvent
Attribute qEvent.VB_VarHelpID = -1
            'Строка "Таблицы персон"
Dim gPerson As PersonInfo
            'Строка "Системной таблицы"
Dim gSystem As SystemInfo
            'Строка "Таблицы терминалов"
Dim gTerminal As TerminalInfo
            'Смещение в файле ресурсов (для локализации приложения)
Dim lngResource As Long
           ' Объявить массив "всплывающих" подсказок
Dim aComment(3, 23) As String
            ' Объявить массив надписей
Dim aCaption(3, 23) As String
             'Новый пароль
Dim strPassword As String
            'Новый индекс языка общения
Dim intLang As Integer
            'Строка отсылаемого сообщения
Dim strMessage As String


            'Перехват нажатия комбинаций клавиш "Alt"+ {"i", "<-" , "->", "+" и "-"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            'Индекс Иконки на форме
Dim intIndex As Integer
            'Режим "Выполнение"
    If chkSetup.Value = 1 And frmDemo.Enabled = True Then
            'Альтернатива "щелчку" мыши на элементе "imgXXXXXXInData"
        If KeyCode = 37 And Shift = 4 Then
            For intIndex = 0 To 3
                If imgParkingInData(intIndex).Visible = True Then
                    Call imgParkingInData_Click(intIndex)
                    Exit Sub
                ElseIf imgAccessInData(intIndex).Visible = True Then
                    Call imgAccessInData_Click(intIndex)
                    Exit Sub
                End If
            Next
            'Альтернатива "щелчку" мыши на элементе "imgXXXXXXOutData"
        ElseIf KeyCode = 39 And Shift = 4 Then
            For intIndex = 0 To 3
                If imgParkingOutData(intIndex).Visible = True Then
                    Call imgParkingOutData_Click(intIndex)
                    Exit Sub
                ElseIf imgAccessOutData(intIndex).Visible = True Then
                    Call imgAccessOutData_Click(intIndex)
                    Exit Sub
                End If
            Next
            'Альтернатива "щелчку" мыши на элементе "imgXXXXXXInfoData"
        ElseIf KeyCode = 73 And Shift = 4 Then
            For intIndex = 0 To 3
                If imgParkingInfoData(intIndex).Visible = True Then
                    Call imgParkingInfoData_Click(intIndex)
                    Exit Sub
                ElseIf imgAccessInfoData(intIndex).Visible = True Then
                    Call imgAccessInfoData_Click(intIndex)
                    Exit Sub
                ElseIf imgEmployeInfoData.Visible = True Then
                    Call imgEmployeInfoData_Click
                    Exit Sub
                End If
            Next
            'Альтернатива "щелчку" мыши на элементе "imgEmployeInData"
        ElseIf KeyCode = 107 And Shift = 4 Then
            If imgEmployeInData.Visible = True Then
                Call imgEmployeInData_Click
                Exit Sub
            End If
            'Альтернатива "щелчку" мыши на элементе "imgEmployeOutData"
        ElseIf KeyCode = 109 And Shift = 4 Then
            If imgEmployeOutData.Visible = True Then
                Call imgEmployeOutData_Click
                Exit Sub
            End If
        End If
    End If
    
End Sub
            
            'Перехват случайного нажатия клавиши "Пробел"
Private Sub chkDummy_Click()
    chkDummy.Value = 0
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus

End Sub

            'Изменение размера окна формы
Private Sub Form_Resize()
            'Максимизация размера окна формы
    frmDemo.WindowState = 0
    
End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Обработка вызова подменю "Form BookKeeper Base" меню "File"
Private Sub mnuFormBookKeeperBase_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Количество строк в "Базе Протокола"
Dim lngProtocolBaseCount As Long
            'Номер файла Архива
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола" и DUMMY файла
Dim lngRecordLen As Long
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer
            'Полное имя DUMMY файла (с указанием "пути" к нему)
Dim strDummyFileName As String
            'Текущий номер строки таблицы DUMMY файла
Dim lngRowDummy As Long
            'Полное имя папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Номер дня (обратный отсчет, начиная с текущего дня),
            '  который просматривается системой при копировании
            '  Архивов Препроцессоа в DUMMY файл
Dim intDayArchive As Integer
            'Количество строк в копируемом файле (Архиве или "TableProtocol")
Dim intRowQuan As Integer
            'Текущий номер строки копируемого Архива
            '   или таблицы "TableProtocol"
Dim intRowNumArchive As Integer
            'Количество строк в "Базе Бухгалтерии"
Dim lngBookKeepingBaseCount As Long
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
Dim lngBookKeepingRowNum As Long
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmDemo.MousePointer = vbHourglass
            'Сделать недоступными элементы управления формы
    frmDemo.Enabled = False
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            'Полное имя файла "Таблица протокола "(с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
    gFileDummy = FreeFile
            'Полное имя DUMMY файла (с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            'Начальная позиция в полном имени DUMMY файла(за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить "старый" DUMMY файл, если он существует
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            'Обработка ошибок
    On Error GoTo UnDefError
            'Открыть DUMMY файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            'Текущий номер  свободной строки DUMMY файла
    gDummyRowNum = 1
            
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
            frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmDemo.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    в DUMMY файл
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmDemo.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол"
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "TableProtocol.dat"

            'Записать строку в файл "Таблицы протокола"
            frmDemo.WriteProtocol
                    
        End If
            
            ' Если это Препроцессор
    Else
            
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
            frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gPreprocName)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmDemo.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gPreprocName)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    в DUMMY файл
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmDemo.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол"
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "TableProtocol.dat"

            'Записать строку в файл "Таблицы протокола"
            frmDemo.WriteProtocol
                    
        End If
    
    End If
            
            'Определить действительный "путь" к каталогу
            '  выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            'Установка свойств элемента "Data" доступа к "Базе Бухгалтерии"
    frmDemo.datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
    frmDemo.datBase.RecordSource = "BookKeeping"
            
            'Определить количество записей в "Базе Бухгалтерии"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngBookKeepingBaseCount = frmDemo.datBase.Recordset.RecordCount
            'Обновить "Базу Бухгалтерии"
    frmDemo.datBase.Recordset.MoveFirst
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
    lngBookKeepingRowNum = 0
    For lngRowDummy = 0 To gDummyRowNum - 1 Step 1
            'Разрешить прерывания для обработки различных событий
        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
        frmDemo.MousePointer = vbHourglass
            'Создать первую обязательную "фиктивную" запись
        If lngRowDummy = 0 Then
            'Отредактировать текущую запись "Базы Бухгалтерии"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = "Fiktive Record"
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = "0000000000000000"
            frmDemo.datBase.Recordset.Fields("Status").Value = "00"
            frmDemo.datBase.Recordset.Fields("Time").Value = "00:00:00AM"
            frmDemo.datBase.Recordset.Fields("Date").Value = "01.01.2000"
            'Обновление записи в "Базе Бухгалтерии"
            frmDemo.datBase.Recordset.Update
        Else
            'Читать строку DUMMY файла в буфер
            Get gFileDummy, lngRowDummy, gProtocol
            'Отредактировать текущую запись "Базы Бухгалтерии"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = gProtocol.strProtocName
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = gProtocol.strProtocPersonCode
            frmDemo.datBase.Recordset.Fields("Status").Value = Left(Trim(gProtocol.strProtocStatus), 2)
            frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            'Событие протокола:
            '                                  - Вход/Выход ("18"/"19") Служащего или
            '                                  - АвтоРегистрация ("16) Служащего или
            '                                  - АвтоУдаление ("17") Служащего или
            '                                  - Регистрация ("12") платного Клиента Автостоянки или
            '                                  - Исключение ("13") платного Клиента Автостоянки
            '                                  - Регистрация ("14") платного Посетителя Предприятия или
            '                                  - Исключение ("15") платного Посетителя Предприятия
            If ((frmDemo.datBase.Recordset.Fields("Status").Value = "00" Or _
            frmDemo.datBase.Recordset.Fields("Status").Value = "01") And _
            (Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Or _
                Right(Trim(gProtocol.strProtocReserve), 6) = "Output") Or _
            (frmDemo.datBase.Recordset.Fields("Status").Value = "01") And _
            (Trim(gProtocol.strProtocReserve) = "AutoRegistration" Or _
                Trim(gProtocol.strProtocReserve) = "AutoDelete") Or _
                (frmDemo.datBase.Recordset.Fields("Status").Value = "05" Or _
                frmDemo.datBase.Recordset.Fields("Status").Value = "06") And _
                (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Or _
                Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark") Or _
                (frmDemo.datBase.Recordset.Fields("Status").Value = "08" Or _
                frmDemo.datBase.Recordset.Fields("Status").Value = "09") And _
                (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Or _
                Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce")) And _
            Left(gProtocol.strProtocName, 1) <> "@" Then
            'Событие  - АвтоРегистрация Служащего (Установить признак
            '  в "Базе Бухгалтерии")
                If Trim(gProtocol.strProtocReserve) = "AutoRegistration" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "16"
            'Событие  - АвтоУдаление Служащего (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Trim(gProtocol.strProtocReserve) = "AutoDelete" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "17"
            'Событие  - Вход Служащего на Предприятие (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "18"
            'Событие  - Выход Служащего с Предприятия (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 6) = "Output" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "19"
            'Событие  - Регистрация Клиента Автостоянки (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "12"
            'Событие  - Исключение Клиента Автостоянки (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "13"
            'Событие  - Регистрация Посетителя Предприятия (Установить
            '  признак в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "14"
            'Событие  - Исключение Посетителя Предприятия (Установить
            '  признак в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "15"
                End If
            'Обновление записи в "Базе Бухгалтерии"
                frmDemo.datBase.Recordset.Update
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
                lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            'Не последняя запись старой "Базы Бухгалтерии"
                If lngBookKeepingRowNum < lngBookKeepingBaseCount Then
                    frmDemo.datBase.Recordset.MoveNext
            'Последняя запись старой "Базы Бухгалтерии"
                Else
                    frmDemo.datBase.Recordset.AddNew
                    frmDemo.datBase.Recordset.Update
                    frmDemo.datBase.Recordset.MoveNext
                End If
            End If
        End If
    Next
            'Текущий номер записи "Базы Бухгалтерии"
    lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            'Удаление одной лишней записи из  "Базы Бухгалтерии"
    If lngBookKeepingRowNum > lngBookKeepingBaseCount Then
        frmDemo.datBase.Recordset.Delete
            'Удаление лишних записей из  "Базы Бухгалтерии",
            '  кроме единственной
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum = 1 Then
        frmDemo.datBase.Recordset.MoveFirst
        frmDemo.datBase.Recordset.MoveNext
        For lngBookKeepingRowNum = 2 To lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
        Next
            'Удаление лишних записей из  "Базы Бухгалтерии"
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum <> 1 Then
        For lngBookKeepingRowNum = lngBookKeepingRowNum To _
        lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
        Next
    End If
            
            'Протоколирование события - "Формирование Базы Протокола"
    gProtocol.strProtocName = "BookKeeperBase"
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
    gProtocol.strProtocReserve = "Creation"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol
            
    GoTo EndProcedure
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            'Закрыть DUMMY файл
    Close gFileDummy
    On Error GoTo 0
            
            'Восстановить стандартный курсор мыши
    frmDemo.MousePointer = 0
            'Сделать доступными элементы управления формы
    frmDemo.Enabled = True
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    

End Sub
            
            'Обработка вызова подменю "Form Protocol Base" меню "File"
Private Sub mnuFormProtocolBase_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Количество строк в "Базе Протокола"
Dim lngProtocolBaseCount As Long
            'Номер файла Архива
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола" и DUMMY файла
Dim lngRecordLen As Long
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer
            'Полное имя DUMMY файла (с указанием "пути" к нему)
Dim strDummyFileName As String
            'Текущий номер строки таблицы DUMMY файла
Dim lngRowDummy As Long
            'Полное имя папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Номер дня (обратный отсчет, начиная с текущего дня),
            '  который просматривается системой при копировании
            '  Архивов Препроцессоа в DUMMY файл
Dim intDayArchive As Integer
            'Количество строк в копируемом файле (Архиве или "TableProtocol")
Dim intRowQuan As Integer
            'Текущий номер строки копируемого Архива
            '   или таблицы "TableProtocol"
Dim intRowNumArchive As Integer
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmDemo.MousePointer = vbHourglass
            'Сделать недоступными элементы управления формы
    frmDemo.Enabled = False
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            'Полное имя файла "Таблица протокола "(с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
    gFileDummy = FreeFile
            'Полное имя DUMMY файла (с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            'Начальная позиция в полном имени DUMMY файла(за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить "старый" DUMMY файл, если он существует
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            'Обработка ошибок
    On Error GoTo UnDefError
            'Открыть DUMMY файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            'Текущий номер  свободной строки DUMMY файла
    gDummyRowNum = 1
            
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
            frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmDemo.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    в DUMMY файл
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmDemo.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол"
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "TableProtocol.dat"

            'Записать строку в файл "Таблицы протокола"
            frmDemo.WriteProtocol
                    
        End If
            
            ' Если это Препроцессор
    Else
            
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
            frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gPreprocName)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmDemo.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gPreprocName)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    в DUMMY файл
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                frmPreprocessors.WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmDemo.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол"
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "TableProtocol.dat"

            'Записать строку в файл "Таблицы протокола"
            frmDemo.WriteProtocol
                    
        End If
    
    End If
            
            'Определить действительный "путь" к каталогу
            '  выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            'Установка свойств элемента "Data" доступа к "Базе Протокола"
    frmDemo.datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
    frmDemo.datBase.RecordSource = "Protocol"
            'Определить количество записей в "Базе Протокола"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngProtocolBaseCount = frmDemo.datBase.Recordset.RecordCount
            'Обновить "Базу Протокола"
    frmDemo.datBase.Recordset.MoveFirst
            'Цикл по всем строкам DUMMY файла
    For lngRowDummy = 1 To gDummyRowNum - 1 Step 1
            'Разрешить прерывания для обработки различных событий
        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
        frmDemo.MousePointer = vbHourglass
            'Читать строку DUMMY файла в буфер
        Get gFileDummy, lngRowDummy, gProtocol
            'Обновить текущую запись "Базы Протокола"
        frmDemo.datBase.Recordset.Edit
        frmDemo.datBase.Recordset.Fields("Name").Value = gProtocol.strProtocName
        frmDemo.datBase.Recordset.Fields("CodeOrPassword").Value = _
        gProtocol.strProtocPersonCode
        frmDemo.datBase.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
        frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
        frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
        frmDemo.datBase.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
        frmDemo.datBase.Recordset.Update
            'Не последняя запись старой "Базы Протокола"
        If lngRowDummy < lngProtocolBaseCount Then
            frmDemo.datBase.Recordset.MoveNext
            'Последняя запись старой "Базы Протокола"
        Else
            frmDemo.datBase.Recordset.AddNew
            frmDemo.datBase.Recordset.Update
            frmDemo.datBase.Recordset.MoveNext
        End If
    Next
            'Удаление одной лишней записи из  "Базы Протокола"
    If lngRowDummy > lngProtocolBaseCount Then
        frmDemo.datBase.Recordset.Delete
            'Удаление лишних записей из  "Базы Протокола"
    Else
        For lngRowDummy = lngRowDummy To lngProtocolBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmDemo.MousePointer = vbHourglass
        Next
    End If
            
            'Протоколирование события - "Формирование Базы Протокола"
    gProtocol.strProtocName = "ProtocolBase"
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
    gProtocol.strProtocReserve = "Creation"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol
    
    GoTo EndProcedure
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            'Закрыть DUMMY файл
    Close gFileDummy
    On Error GoTo 0
            
            'Восстановить стандартный курсор мыши
    frmDemo.MousePointer = 0
            'Сделать доступными элементы управления формы
    frmDemo.Enabled = True
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    
End Sub

            ' Изменение состояния переключателя "Выполнение/Установки"
Private Sub chkSetup_Click()
            ' Определить состояние преключателя
    If chkSetup.Value = 0 Then
            ' Режим "Установка" - сделать доступными элементы управления
        
            ' Запретить прием/передачу информации для терминалов
        If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
        If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
        If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
        If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            'Запретить опрос "Controller'ов" по таймерам
        gTermContr = 0
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        If imgParkingInfoData(0).Visible = True Then
            imgParkingInData(0).Enabled = False
            imgParkingOutData(0).Enabled = False
            imgParkingInfoData(0).Enabled = False
        End If
        If imgParkingInfoData(1).Visible = True Then
            imgParkingInData(1).Enabled = False
            imgParkingOutData(1).Enabled = False
            imgParkingInfoData(1).Enabled = False
        End If
        If imgParkingInfoData(2).Visible = True Then
            imgParkingInData(2).Enabled = False
            imgParkingOutData(2).Enabled = False
            imgParkingInfoData(2).Enabled = False
        End If
        If imgParkingInfoData(3).Visible = True Then
            imgParkingInData(3).Enabled = False
            imgParkingOutData(3).Enabled = False
            imgParkingInfoData(3).Enabled = False
        End If
            
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
        If imgAccessInfoData(0).Visible = True Then
            imgAccessInData(0).Enabled = False
            imgAccessOutData(0).Enabled = False
            imgAccessInfoData(0).Enabled = False
        End If
        If imgAccessInfoData(1).Visible = True Then
            imgAccessInData(1).Enabled = False
            imgAccessOutData(1).Enabled = False
            imgAccessInfoData(1).Enabled = False
        End If
        If imgAccessInfoData(2).Visible = True Then
            imgAccessInData(2).Enabled = False
            imgAccessOutData(2).Enabled = False
            imgAccessInfoData(2).Enabled = False
        End If
        If imgAccessInfoData(3).Visible = True Then
            imgAccessInData(3).Enabled = False
            imgAccessOutData(3).Enabled = False
            imgAccessInfoData(3).Enabled = False
        End If
            
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Служащих, Информация)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = False
            imgEmployeOutData.Enabled = False
            imgEmployeInfoData.Enabled = False
        End If
            
            ' Сделать недоступными кнопки ручного управления терминалами
        cmdOpen(0).Enabled = False
        cmdOpen(1).Enabled = False
        cmdOpen(2).Enabled = False
        cmdOpen(3).Enabled = False
            'Протоколирование события - щелчок мышью на переключателе "Execute/Setup"
        gProtocol.strProtocName = "????????????????"
            'Системный пароль
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "SETUP option"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            'Установить фокус на поле пароля
        txtPassword.Enabled = True
        txtPassword.SetFocus
            'Установить контроль времени ввода пароля
        tmrPasswTimeOut.Enabled = True
            'Удержание фокуса клавиатуры на поле пароля до его ввода
        Do While txtPassword.Enabled = True
            DoEvents
        Loop
        
            'Событие "TimeOut" при вводе пароля - продолжить выполнение программы
        If tmrPasswTimeOut.Enabled = False Then
            ' Режим "Выполнение" - сделать недоступными элементы управления
            GoTo Execution
        End If
        
            'Сбросить контроль времени ввода пароля
        tmrPasswTimeOut.Enabled = False
        
            'Сделать доступным переключатель "Выполнение/Установки"
        chkSetup.Enabled = True
        
            
            'Сделать доступным поле пароля - для возможного ввода нового пароля
        txtPassword.Enabled = True
        
            'Сделать видимым меню
        mnuFile.Visible = True
        mnuAdjustment.Visible = True
        mnuParking.Visible = True
        mnuAccess.Visible = True
        mnuEmploye.Visible = True
            'Сделать видимой панель инструментов
        picTools.Visible = True
            ' Сделать доступным опции языка общения
        fraFlag.Enabled = True
        optEnglish.Enabled = True
        optLatvian.Enabled = True
        optRussian.Enabled = True
            ' Сделать доступными переключатели выбираемых терминалов
        chkTerm(0).Enabled = True
        chkTerm(1).Enabled = True
        chkTerm(2).Enabled = True
        chkTerm(3).Enabled = True
            ' "Проявить" этикетку выбираемых терминалов
        lblTerminals.Enabled = True
           ' Сделать доступными переключатели фотоизображений
        chkPhoto(0).Enabled = True
        chkPhoto(1).Enabled = True
        chkPhoto(2).Enabled = True
        chkPhoto(3).Enabled = True
            ' "Проявить" этикетку переключателя фотоизображений
        lblPhoto.Enabled = True
            ' Сделать доступными опции "Автоматическое/Ручное" управление терминалами
        fraControl.Enabled = True
        optAutomatic.Enabled = True
        optManual.Enabled = True
               ' Сделать недоступными кнопки ручного управления терминалами
        cmdOpen(0).Enabled = False
        cmdOpen(1).Enabled = False
        cmdOpen(2).Enabled = False
        cmdOpen(3).Enabled = False
                ' "Погасить" этикетку кнопок ручного управления терминалами
        lblOpen.Enabled = False


    Else
            ' Режим "Выполнение" - сделать недоступными элементы управления
Execution:
            'Сделать недоступным поле пароля
        txtPassword.Enabled = False
        
        'Протоколирование события - щелчок мышью на переключателе "Execute/Setup"
        gProtocol.strProtocName = "????????????????"
            'Системный пароль
        gProtocol.strProtocPersonCode = txtPassword.Tag
            'Статус
        gProtocol.strProtocStatus = "04 - Manager"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "EXECUTE option"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
        
            'Сделать невидимым меню
        mnuFile.Visible = False
        mnuAdjustment.Visible = False
        mnuParking.Visible = False
        mnuAccess.Visible = False
        mnuEmploye.Visible = False
            'Сделать невидимой панель инструментов
        picTools.Visible = False
    ' Сделать недоступным опции языка общения
        fraFlag.Enabled = False
        optEnglish.Enabled = False
        optLatvian.Enabled = False
        optRussian.Enabled = False
            ' Сделать недоступными переключатели выбираемых терминалов
        chkTerm(0).Enabled = False
        chkTerm(1).Enabled = False
        chkTerm(2).Enabled = False
        chkTerm(3).Enabled = False
            ' "Погасить" этикетку выбираемых терминалов
        lblTerminals.Enabled = False
           ' Сделать недоступными переключатели фотоизображений
        chkPhoto(0).Enabled = False
        chkPhoto(1).Enabled = False
        chkPhoto(2).Enabled = False
        chkPhoto(3).Enabled = False
            ' "Очистить" поля фотоизображений
        imgPhoto(0).Picture = LoadPicture("")
        imgPhoto(1).Picture = LoadPicture("")
        imgPhoto(2).Picture = LoadPicture("")
        imgPhoto(3).Picture = LoadPicture("")
            ' "Погасить" этикетку переключателя фотоизображений
        lblPhoto.Enabled = False
            ' Сделать недоступным опции "Автоматическое/Ручное" управление терминалами
        fraControl.Enabled = False
        optAutomatic.Enabled = False
        optManual.Enabled = False
    
            'Если выбрана опция "Автоматическое" управление терминалами
        If optAutomatic.Value = True Then
            'Восстановить исходное состояние и надпись электронной "Кнопки"
            cmdOpen(0).Tag = 0
            cmdOpen(0).Caption = chkTerm(0).Caption
            cmdOpen(1).Tag = 0
            cmdOpen(1).Caption = chkTerm(1).Caption
            cmdOpen(2).Tag = 0
            cmdOpen(2).Caption = chkTerm(2).Caption
            cmdOpen(3).Tag = 0
            cmdOpen(3).Caption = chkTerm(3).Caption
               ' Сделать недоступными кнопки ручного управления терминалами
            cmdOpen(0).Enabled = False
            cmdOpen(1).Enabled = False
            cmdOpen(2).Enabled = False
            cmdOpen(3).Enabled = False
                ' "Погасить" этикетку кнопок ручного управления терминалами
            lblOpen.Enabled = False
                'Если выбрана опция "Ручное" управление терминалами
        Else
            'Адреса "Controller'ов", управляемых от электронной "Кнопки"
            cmdOpen(0).Tag = CByte(CInt(Trim(gAddrManual(0))))
            cmdOpen(0).Caption = Trim(gAddrManual(0))
            cmdOpen(1).Tag = CByte(CInt(Trim(gAddrManual(1))))
            cmdOpen(1).Caption = Trim(gAddrManual(1))
            cmdOpen(2).Tag = CByte(CInt(Trim(gAddrManual(2))))
            cmdOpen(2).Caption = Trim(gAddrManual(2))
            cmdOpen(3).Tag = CByte(CInt(Trim(gAddrManual(3))))
            cmdOpen(3).Caption = Trim(gAddrManual(3))
               ' Сделать доступными кнопки ручного управления отмеченных терминалов
            If chkTerm(0).Value = 1 Then cmdOpen(0).Enabled = True
            If chkTerm(1).Value = 1 Then cmdOpen(1).Enabled = True
            If chkTerm(2).Value = 1 Then cmdOpen(2).Enabled = True
            If chkTerm(3).Value = 1 Then cmdOpen(3).Enabled = True
                ' "Проявить" этикетку кнопок ручного управления терминалами
            lblOpen.Enabled = True
            
        End If
        
            ' Разрешить прием/передачу информации для отмеченных терминалов
            ' Логические порты "свободны" - могут обрабатывать данные
            ' отмеченных терминалов
        If chkTerm(0).Value = 1 Then
            If prtPortC(0).PortOpen = False Then prtPortC(0).PortOpen = True
            prtPortC(0).Tag = 0
        End If
        If chkTerm(1).Value = 1 Then
            If prtPortC(1).PortOpen = False Then prtPortC(1).PortOpen = True
            prtPortC(1).Tag = 0
        End If
        If chkTerm(2).Value = 1 Then
            If prtPortC(2).PortOpen = False Then prtPortC(2).PortOpen = True
            prtPortC(2).Tag = 0
        End If
        If chkTerm(3).Value = 1 Then
            If prtPortC(3).PortOpen = False Then prtPortC(3).PortOpen = True
            prtPortC(3).Tag = 0
        End If
            'Разрешить опрос "Controller'ов" по таймерам
        gTermContr = 1
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        If imgParkingInfoData(0).Visible = True Then
            imgParkingInData(0).Enabled = True
            imgParkingOutData(0).Enabled = True
            imgParkingInfoData(0).Enabled = True
        End If
        If imgParkingInfoData(1).Visible = True Then
            imgParkingInData(1).Enabled = True
            imgParkingOutData(1).Enabled = True
            imgParkingInfoData(1).Enabled = True
        End If
        If imgParkingInfoData(2).Visible = True Then
            imgParkingInData(2).Enabled = True
            imgParkingOutData(2).Enabled = True
            imgParkingInfoData(2).Enabled = True
        End If
        If imgParkingInfoData(3).Visible = True Then
            imgParkingInData(3).Enabled = True
            imgParkingOutData(3).Enabled = True
            imgParkingInfoData(3).Enabled = True
        End If
    
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
        If imgAccessInfoData(0).Visible = True Then
            imgAccessInData(0).Enabled = True
            imgAccessOutData(0).Enabled = True
            imgAccessInfoData(0).Enabled = True
        End If
        If imgAccessInfoData(1).Visible = True Then
            imgAccessInData(1).Enabled = True
            imgAccessOutData(1).Enabled = True
            imgAccessInfoData(1).Enabled = True
        End If
        If imgAccessInfoData(2).Visible = True Then
            imgAccessInData(2).Enabled = True
            imgAccessOutData(2).Enabled = True
            imgAccessInfoData(2).Enabled = True
        End If
        If imgAccessInfoData(3).Visible = True Then
            imgAccessInData(3).Enabled = True
            imgAccessOutData(3).Enabled = True
            imgAccessInfoData(3).Enabled = True
        End If
            
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Служащих, Информация)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = True
            imgEmployeOutData.Enabled = True
            imgEmployeInfoData.Enabled = True
        End If
            
            'Сделать доступным переключатель "Выполнение/Установки"
        chkSetup.Enabled = True
            'Установить режим "Выполнение"
        chkSetup.Value = 1
            ' Разрешить выход из программы управления
        cmdExit.Enabled = True
            'Установить фокус на опции "Dummy"
        If frmDemo.Visible = True Then chkDummy.SetFocus

    End If

End Sub

            ' Нажата кнопка "Exit"
Private Sub cmdExit_Click()
            'Код возврата при сохранении "Таблицы персон"
Dim intSaveTablePerson As Integer
    
            ' Запретить прием/передачу информации для терминалов
    If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
    If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
    If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
    If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        If imgParkingInfoData(0).Visible = True Then
            imgParkingInData(0).Enabled = False
            imgParkingOutData(0).Enabled = False
            imgParkingInfoData(0).Enabled = False
        End If
        If imgParkingInfoData(1).Visible = True Then
            imgParkingInData(1).Enabled = False
            imgParkingOutData(1).Enabled = False
            imgParkingInfoData(1).Enabled = False
        End If
        If imgParkingInfoData(2).Visible = True Then
            imgParkingInData(2).Enabled = False
            imgParkingOutData(2).Enabled = False
            imgParkingInfoData(2).Enabled = False
        End If
        If imgParkingInfoData(3).Visible = True Then
            imgParkingInData(3).Enabled = False
            imgParkingOutData(3).Enabled = False
            imgParkingInfoData(3).Enabled = False
        End If
            
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
        If imgAccessInfoData(0).Visible = True Then
            imgAccessInData(0).Enabled = False
            imgAccessOutData(0).Enabled = False
            imgAccessInfoData(0).Enabled = False
        End If
        If imgAccessInfoData(1).Visible = True Then
            imgAccessInData(1).Enabled = False
            imgAccessOutData(1).Enabled = False
            imgAccessInfoData(1).Enabled = False
        End If
        If imgAccessInfoData(2).Visible = True Then
            imgAccessInData(2).Enabled = False
            imgAccessOutData(2).Enabled = False
            imgAccessInfoData(2).Enabled = False
        End If
        If imgAccessInfoData(3).Visible = True Then
            imgAccessInData(3).Enabled = False
            imgAccessOutData(3).Enabled = False
            imgAccessInfoData(3).Enabled = False
        End If
            
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Служащих, Информация)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = False
            imgEmployeOutData.Enabled = False
            imgEmployeInfoData.Enabled = False
        End If
            
            ' Сделать недоступными кнопки ручного управления терминалами
    cmdOpen(0).Enabled = False
    cmdOpen(1).Enabled = False
    cmdOpen(2).Enabled = False
    cmdOpen(3).Enabled = False
            'Протоколирование события - нажата кнопка "Exit"
    gProtocol.strProtocName = "????????????????"
            'Системный пароль
    gProtocol.strProtocPersonCode = ""
            'Статус
    gProtocol.strProtocStatus = ""
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "EXIT button"
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
    
            'Установить фокус на поле пароля
    txtPassword.Enabled = True
    txtPassword.SetFocus
            'Установить контроль времени ввода пароля
    tmrPasswTimeOut.Enabled = True
            'Удержание фокуса клавиатуры на поле пароля до его ввода
    Do While txtPassword.Enabled = True
        DoEvents
    Loop
    
            'Событие "TimeOut" при вводе пароля - продолжить выполнение программы
    If tmrPasswTimeOut.Enabled = False Then
            ' Режим "Выполнение" - сделать недоступными элементы управления
            
            ' "Очистить" поля фотоизображений
        imgPhoto(0).Picture = LoadPicture("")
        imgPhoto(1).Picture = LoadPicture("")
        imgPhoto(2).Picture = LoadPicture("")
        imgPhoto(3).Picture = LoadPicture("")
    
            'Если выбрана опция "Автоматическое" управление терминалами
        If optAutomatic.Value = True Then
               ' Сделать недоступными кнопки ручного управления терминалами
            cmdOpen(0).Enabled = False
            cmdOpen(1).Enabled = False
            cmdOpen(2).Enabled = False
            cmdOpen(3).Enabled = False
                ' "Погасить" этикетку кнопок ручного управления терминалами
            lblOpen.Enabled = False
                    'Если выбрана опция "Ручное" управление терминалами
        Else
               ' Сделать доступными кнопки ручного управления отмеченных терминалов
            If chkTerm(0).Value = 1 Then cmdOpen(0).Enabled = True
            If chkTerm(1).Value = 1 Then cmdOpen(1).Enabled = True
            If chkTerm(2).Value = 1 Then cmdOpen(2).Enabled = True
            If chkTerm(3).Value = 1 Then cmdOpen(3).Enabled = True
                ' "Проявить" этикетку кнопок ручного управления терминалами
            lblOpen.Enabled = True
        End If
        
            ' Разрешить прием/передачу информации для отмеченных терминалов
            ' Логические порты "свободны" - могут обрабатывать данные отмеченных терминалов
        If chkTerm(0).Value = 1 And prtPortC(0).PortOpen = False Then
            prtPortC(0).PortOpen = True
            prtPortC(0).Tag = 0
        End If
        If chkTerm(1).Value = 1 And prtPortC(1).PortOpen = False Then
            prtPortC(1).PortOpen = True
            prtPortC(1).Tag = 0
        End If
        If chkTerm(2).Value = 1 And prtPortC(2).PortOpen = False Then
            prtPortC(2).PortOpen = True
            prtPortC(2).Tag = 0
        End If
        If chkTerm(3).Value = 1 And prtPortC(3).PortOpen = False Then
            prtPortC(3).PortOpen = True
            prtPortC(3).Tag = 0
        End If
        
            'Разрешить опрос "Controller'ов" по таймерам
        gTermContr = 1
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        If imgParkingInfoData(0).Visible = True Then
            imgParkingInData(0).Enabled = True
            imgParkingOutData(0).Enabled = True
            imgParkingInfoData(0).Enabled = True
        End If
        If imgParkingInfoData(1).Visible = True Then
            imgParkingInData(1).Enabled = True
            imgParkingOutData(1).Enabled = True
            imgParkingInfoData(1).Enabled = True
        End If
        If imgParkingInfoData(2).Visible = True Then
            imgParkingInData(2).Enabled = True
            imgParkingOutData(2).Enabled = True
            imgParkingInfoData(2).Enabled = True
        End If
        If imgParkingInfoData(3).Visible = True Then
            imgParkingInData(3).Enabled = True
            imgParkingOutData(3).Enabled = True
            imgParkingInfoData(3).Enabled = True
        End If
            
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
        If imgAccessInfoData(0).Visible = True Then
            imgAccessInData(0).Enabled = True
            imgAccessOutData(0).Enabled = True
            imgAccessInfoData(0).Enabled = True
        End If
        If imgAccessInfoData(1).Visible = True Then
            imgAccessInData(1).Enabled = True
            imgAccessOutData(1).Enabled = True
            imgAccessInfoData(1).Enabled = True
        End If
        If imgAccessInfoData(2).Visible = True Then
            imgAccessInData(2).Enabled = True
            imgAccessOutData(2).Enabled = True
            imgAccessInfoData(2).Enabled = True
        End If
        If imgAccessInfoData(3).Visible = True Then
            imgAccessInData(3).Enabled = True
            imgAccessOutData(3).Enabled = True
            imgAccessInfoData(3).Enabled = True
        End If
            
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Служащих, Информация)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = True
            imgEmployeOutData.Enabled = True
            imgEmployeInfoData.Enabled = True
        End If
            
            'Сделать доступным переключатель "Выполнение/Установки"
        chkSetup.Enabled = True
            ' Разрешить выход из программы управления
        cmdExit.Enabled = True
            'Установить фокус на опции "Dummy"
        If frmDemo.Visible = True Then chkDummy.SetFocus
        
            'Пароль введен за допустимое время
    Else
            'Установлен признак необходимости сжатия "Таблицы персон":
            '   устанавливается всегда в "Host Computer'e" и в тех случаях
            '   в "Препроцессоре", когда последний использует свою
            '   собственную "Таблицу персон" - "ЗЕРКАЛЬНАЯ Таблицa персон"
        If gCompresTablPers = 1 Then
            'Установлен признак внесенных изменений в "Таблицу персон"
            '  - сохранить таблицу в умалчиваемом файле
            If gChangesTablePerson = True Then _
                Call frmTablePerson.SaveTablePerson
            ' Если имеются Препроцессоры в локальной сети
            If gNetPreprocNum > 0 Then
            'Строка передачи сообщения
                strMessage = "ExitApp"
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                Call SendMessage(strMessage)
            End If
        End If
            
            'Протоколирование события - "Разгрузка системы"
        gProtocol.strProtocName = "################"
            'Системный пароль
        gProtocol.strProtocPersonCode = txtPassword.Tag
            'Статус
        gProtocol.strProtocStatus = "04 - Manager"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "Unload the Acc. Syst."
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            'Нет признака запрета формирования баз Протокола и Бухгалтерии
        If gMSBase = 1 Then
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
            Call BasesConvert
        End If
            
            'Закрыть файл "Таблицы протокола"
        Close gProtocFileNum
    
            ' Освободить ссылку на объект "FlexGrid" ("Таблица Персон")
        Set gTablePerson = Nothing
            ' Освободить ссылку на объект ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' Освободить ссылку на объект MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set qQueueInput = Nothing
            ' Освободить ссылку на объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set evQueue = Nothing
            ' Освободить ссылку на объект ПРИНИМАЕМОЕ СООБЩЕНИE
        Set qMsgInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        Set qQueueOutput = Nothing
            ' Освободить ссылку на объект ПЕРЕДАВАЕМОЕ СООБЩЕНИE
        Set qMsgOutput = Nothing
            
            'Завершить программу
        End
    End If
    
End Sub

            ' Загрузить форму
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы протокола"
Dim lngRecordLen As Long
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Инициализация смещения в файле ресурсов
    lngResource = 101
            ' Инициализировать массив надписей
    aCaption(0, 0) = LoadResString(lngResource)
    aCaption(1, 0) = LoadResString(lngResource + 23)
    aCaption(2, 0) = LoadResString(lngResource + 46)
    aCaption(0, 1) = LoadResString(lngResource + 1)
    aCaption(1, 1) = LoadResString(lngResource + 23 + 1)
    aCaption(2, 1) = LoadResString(lngResource + 46 + 1)
    aCaption(0, 2) = LoadResString(lngResource + 2)
    aCaption(1, 2) = LoadResString(lngResource + 23 + 2)
    aCaption(2, 2) = LoadResString(lngResource + 46 + 2)
    aCaption(0, 3) = LoadResString(lngResource + 3)
    aCaption(1, 3) = LoadResString(lngResource + 23 + 3)
    aCaption(2, 3) = LoadResString(lngResource + 46 + 3)
    aCaption(0, 4) = LoadResString(lngResource + 4)
    aCaption(1, 4) = LoadResString(lngResource + 23 + 4)
    aCaption(2, 4) = LoadResString(lngResource + 46 + 4)
    aCaption(0, 5) = LoadResString(lngResource + 5)
    aCaption(1, 5) = LoadResString(lngResource + 23 + 5)
    aCaption(2, 5) = LoadResString(lngResource + 46 + 5)
    aCaption(0, 6) = LoadResString(lngResource + 6)
    aCaption(1, 6) = LoadResString(lngResource + 23 + 6)
    aCaption(2, 6) = LoadResString(lngResource + 46 + 6)
    
    aCaption(0, 7) = LoadResString(lngResource + 7)
    aCaption(1, 7) = LoadResString(lngResource + 23 + 7)
    aCaption(2, 7) = LoadResString(lngResource + 46 + 7)
    aCaption(0, 8) = LoadResString(lngResource + 8)
    aCaption(1, 8) = LoadResString(lngResource + 23 + 8)
    aCaption(2, 8) = LoadResString(lngResource + 46 + 8)
    aCaption(0, 9) = LoadResString(lngResource + 9)
    aCaption(1, 9) = LoadResString(lngResource + 23 + 9)
    aCaption(2, 9) = LoadResString(lngResource + 46 + 9)
    aCaption(0, 10) = LoadResString(lngResource + 10)
    aCaption(1, 10) = LoadResString(lngResource + 23 + 10)
    aCaption(2, 10) = LoadResString(lngResource + 46 + 10)
    aCaption(0, 11) = LoadResString(lngResource + 11)
    aCaption(1, 11) = LoadResString(lngResource + 23 + 11)
    aCaption(2, 11) = LoadResString(lngResource + 46 + 11)
    aCaption(0, 12) = LoadResString(lngResource + 12)
    aCaption(1, 12) = LoadResString(lngResource + 23 + 12)
    aCaption(2, 12) = LoadResString(lngResource + 46 + 12)
    aCaption(0, 13) = LoadResString(lngResource + 13)
    aCaption(1, 13) = LoadResString(lngResource + 23 + 13)
    aCaption(2, 13) = LoadResString(lngResource + 46 + 13)
    aCaption(0, 14) = LoadResString(lngResource + 14)
    aCaption(1, 14) = LoadResString(lngResource + 23 + 14)
    aCaption(2, 14) = LoadResString(lngResource + 46 + 14)
    aCaption(0, 15) = LoadResString(lngResource + 15)
    aCaption(1, 15) = LoadResString(lngResource + 23 + 15)
    aCaption(2, 15) = LoadResString(lngResource + 46 + 15)
    aCaption(0, 16) = LoadResString(lngResource + 16)
    aCaption(1, 16) = LoadResString(lngResource + 23 + 16)
    aCaption(2, 16) = LoadResString(lngResource + 46 + 16)
    aCaption(0, 17) = LoadResString(lngResource + 17)
    aCaption(1, 17) = LoadResString(lngResource + 23 + 17)
    aCaption(2, 17) = LoadResString(lngResource + 46 + 17)
    aCaption(0, 18) = LoadResString(lngResource + 18)
    aCaption(1, 18) = LoadResString(lngResource + 23 + 18)
    aCaption(2, 18) = LoadResString(lngResource + 46 + 18)
    aCaption(0, 19) = LoadResString(lngResource + 19)
    aCaption(1, 19) = LoadResString(lngResource + 23 + 19)
    aCaption(2, 19) = LoadResString(lngResource + 46 + 19)
    aCaption(0, 20) = LoadResString(lngResource + 20)
    aCaption(1, 20) = LoadResString(lngResource + 23 + 20)
    aCaption(2, 20) = LoadResString(lngResource + 46 + 20)
    aCaption(0, 21) = LoadResString(lngResource + 21)
    aCaption(1, 21) = LoadResString(lngResource + 23 + 21)
    aCaption(2, 21) = LoadResString(lngResource + 46 + 21)
    aCaption(0, 22) = LoadResString(lngResource + 22)
    aCaption(1, 22) = LoadResString(lngResource + 23 + 22)
    aCaption(2, 22) = LoadResString(lngResource + 46 + 22)
            'Сделать невидимым меню
    mnuFile.Visible = False
    mnuAdjustment.Visible = False
    mnuParking.Visible = False
    mnuAccess.Visible = False
    mnuEmploye.Visible = False
            'Сделать невидимой панель инструментов
    picTools.Visible = False
    
            ' Инициализировать массив "всплывающих" подсказок
    aComment(0, 0) = LoadResString(lngResource + 69)
    aComment(1, 0) = LoadResString(lngResource + 92)
    aComment(2, 0) = LoadResString(lngResource + 115)
    aComment(0, 1) = LoadResString(lngResource + 69 + 1)
    aComment(1, 1) = LoadResString(lngResource + 92 + 1)
    aComment(2, 1) = LoadResString(lngResource + 115 + 1)
    aComment(0, 2) = LoadResString(lngResource + 69 + 2)
    aComment(1, 2) = LoadResString(lngResource + 92 + 2)
    aComment(2, 2) = LoadResString(lngResource + 115 + 2)
    aComment(0, 3) = LoadResString(lngResource + 69 + 3)
    aComment(1, 3) = LoadResString(lngResource + 92 + 3)
    aComment(2, 3) = LoadResString(lngResource + 115 + 3)
    aComment(0, 4) = LoadResString(lngResource + 69 + 4)
    aComment(1, 4) = LoadResString(lngResource + 92 + 4)
    aComment(2, 4) = LoadResString(lngResource + 115 + 4)
    aComment(0, 5) = LoadResString(lngResource + 69 + 5)
    aComment(1, 5) = LoadResString(lngResource + 92 + 5)
    aComment(2, 5) = LoadResString(lngResource + 115 + 5)
    aComment(0, 6) = LoadResString(lngResource + 69 + 6)
    aComment(1, 6) = LoadResString(lngResource + 92 + 6)
    aComment(2, 6) = LoadResString(lngResource + 115 + 6)
    aComment(0, 7) = LoadResString(lngResource + 69 + 7)
    aComment(1, 7) = LoadResString(lngResource + 92 + 7)
    aComment(2, 7) = LoadResString(lngResource + 115 + 7)
    aComment(0, 8) = LoadResString(lngResource + 69 + 8)
    aComment(1, 8) = LoadResString(lngResource + 92 + 8)
    aComment(2, 8) = LoadResString(lngResource + 115 + 8)
    aComment(0, 9) = LoadResString(lngResource + 69 + 9)
    aComment(1, 9) = LoadResString(lngResource + 92 + 9)
    aComment(2, 9) = LoadResString(lngResource + 115 + 9)
    aComment(0, 10) = LoadResString(lngResource + 69 + 10)
    aComment(1, 10) = LoadResString(lngResource + 92 + 10)
    aComment(2, 10) = LoadResString(lngResource + 115 + 10)
    aComment(0, 11) = LoadResString(lngResource + 69 + 11)
    aComment(1, 11) = LoadResString(lngResource + 92 + 11)
    aComment(2, 11) = LoadResString(lngResource + 115 + 11)
    aComment(0, 12) = LoadResString(lngResource + 69 + 12)
    aComment(1, 12) = LoadResString(lngResource + 92 + 12)
    aComment(2, 12) = LoadResString(lngResource + 115 + 12)
    aComment(0, 13) = LoadResString(lngResource + 69 + 13)
    aComment(1, 13) = LoadResString(lngResource + 92 + 13)
    aComment(2, 13) = LoadResString(lngResource + 115 + 13)
    aComment(0, 14) = LoadResString(lngResource + 69 + 14)
    aComment(1, 14) = LoadResString(lngResource + 92 + 14)
    aComment(2, 14) = LoadResString(lngResource + 115 + 14)
    aComment(0, 15) = LoadResString(lngResource + 69 + 15)
    aComment(1, 15) = LoadResString(lngResource + 92 + 15)
    aComment(2, 15) = LoadResString(lngResource + 115 + 15)
    aComment(0, 16) = LoadResString(lngResource + 69 + 16)
    aComment(1, 16) = LoadResString(lngResource + 92 + 16)
    aComment(2, 16) = LoadResString(lngResource + 115 + 16)
    aComment(0, 17) = LoadResString(lngResource + 69 + 17)
    aComment(1, 17) = LoadResString(lngResource + 92 + 17)
    aComment(2, 17) = LoadResString(lngResource + 115 + 17)
    aComment(0, 18) = LoadResString(lngResource + 69 + 18)
    aComment(1, 18) = LoadResString(lngResource + 92 + 18)
    aComment(2, 18) = LoadResString(lngResource + 115 + 18)
    aComment(0, 19) = LoadResString(lngResource + 69 + 19)
    aComment(1, 19) = LoadResString(lngResource + 92 + 19)
    aComment(2, 19) = LoadResString(lngResource + 115 + 19)
    aComment(0, 20) = LoadResString(lngResource + 69 + 20)
    aComment(1, 20) = LoadResString(lngResource + 92 + 20)
    aComment(2, 20) = LoadResString(lngResource + 115 + 20)
    aComment(0, 21) = LoadResString(lngResource + 69 + 21)
    aComment(1, 21) = LoadResString(lngResource + 92 + 21)
    aComment(2, 21) = LoadResString(lngResource + 115 + 21)
    aComment(0, 22) = LoadResString(lngResource + 69 + 22)
    aComment(1, 22) = LoadResString(lngResource + 92 + 22)
    aComment(2, 22) = LoadResString(lngResource + 115 + 22)
            'Установить опцию языка общения и отобразить флаг
    If fraFlag.Tag = 0 Then
    optEnglish.Value = True
    imgEnglish.Visible = True
    imgLatvian.Visible = False
    imgRussian.Visible = False
    End If
    
    If fraFlag.Tag = 1 Then
    optLatvian.Value = True
    imgEnglish.Visible = False
    imgLatvian.Visible = True
    imgRussian.Visible = False
    End If
    
    If fraFlag.Tag = 2 Then
    optRussian.Value = True
    imgEnglish.Visible = False
    imgLatvian.Visible = False
    imgRussian.Visible = True
    End If
            'Новый индекс языка общения
    intLang = fraFlag.Tag
            ' Инициализировать массив надписей
    chkSetup.Caption = aCaption(intLang, 0)
    fraFlag.Caption = aCaption(intLang, 1)
    optEnglish.Caption = aCaption(intLang, 2)
    optLatvian.Caption = aCaption(intLang, 3)
    optRussian.Caption = aCaption(intLang, 4)
    lblTerminals.Caption = aCaption(intLang, 5)
    chkTerm(0).Caption = aCaption(intLang, 6)
    chkTerm(1).Caption = aCaption(intLang, 7)
    chkTerm(2).Caption = aCaption(intLang, 8)
    chkTerm(3).Caption = aCaption(intLang, 9)
    lblPhoto.Caption = aCaption(intLang, 10)
    cmdExit.Caption = aCaption(intLang, 11)
    optAutomatic.Caption = aCaption(intLang, 12)
    lblOpen.Caption = aCaption(intLang, 13)
    cmdOpen(0).Caption = aCaption(intLang, 14)
    cmdOpen(1).Caption = aCaption(intLang, 15)
    cmdOpen(2).Caption = aCaption(intLang, 16)
    cmdOpen(3).Caption = aCaption(intLang, 17)
    optManual.Caption = aCaption(intLang, 18)
    lblPassword.Caption = aCaption(intLang, 19)
    mnuFile.Caption = aCaption(intLang, 20)
    mnuAdjustment.Caption = aCaption(intLang, 21)
    fraControl.Caption = aCaption(intLang, 22)
            ' Инициализировать массив "всплывающих" подсказок
    txtPassword.ToolTipText = aComment(intLang, 0)
    chkSetup.ToolTipText = aComment(intLang, 1)
    optEnglish.ToolTipText = aComment(intLang, 2)
    optLatvian.ToolTipText = aComment(intLang, 3)
    optRussian.ToolTipText = aComment(intLang, 4)
    chkTerm(0).ToolTipText = aComment(intLang, 5)
    chkTerm(1).ToolTipText = aComment(intLang, 6)
    chkTerm(2).ToolTipText = aComment(intLang, 7)
    chkTerm(3).ToolTipText = aComment(intLang, 8)
    chkPhoto(0).ToolTipText = aComment(intLang, 9)
    chkPhoto(1).ToolTipText = aComment(intLang, 10)
    chkPhoto(2).ToolTipText = aComment(intLang, 11)
    chkPhoto(3).ToolTipText = aComment(intLang, 12)
    cmdExit.ToolTipText = aComment(intLang, 13)
    optAutomatic.ToolTipText = aComment(intLang, 14)
    optManual.ToolTipText = aComment(intLang, 15)
'    chkManual(1).ToolTipText = aComment(intLang, 16)
'    chkManual(2).ToolTipText = aComment(intLang, 17)
'    chkManual(3).ToolTipText = aComment(intLang, 18)
    cmdOpen(0).ToolTipText = aComment(intLang, 19)
    cmdOpen(1).ToolTipText = aComment(intLang, 20)
    cmdOpen(2).ToolTipText = aComment(intLang, 21)
    cmdOpen(3).ToolTipText = aComment(intLang, 22)
        
            
            ' Создание объекта MSMQQueueInfo для управления
            '  очередью ПРИНИМАЕМЫХ СООБЩЕНИЙ
    Set qInfoInput = New MSMQQueueInfo
            ' Установить путь к очереди ПРИНИМАЕМЫХ СООБЩЕНИЙ
    qInfoInput.PathName = ".\Private$\GeneralQueue"
            ' Присвоить имя очереди ПРИНИМАЕМЫХ СООБЩЕНИЙ
    qInfoInput.Label = "Input Message Queue"
    On Error Resume Next
            ' Создать очередь ПРИНИМАЕМЫХ СООБЩЕНИЙ
            '   информацией
    qInfoInput.Create
    On Error GoTo 0
            ' Открыть очередь сообщений с параметрами (для приема
            '   сообщений, доступ к очереди разрешен всем)
    Set qQueueInput = qInfoInput.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
            ' Создать экземпляр объекта MSMQEvent и его активизация
    Set qEvent = New MSMQEvent
    qQueueInput.EnableNotification qEvent
            
            
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла "Таблицы протокола"
    gProtocFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Файл "Таблицы протокола" требует архивирования
    If FileLen(strPathFileName) / lngRecordLen + 1 > 32760 Then
            'Издать звуковой сигнал
        BeepSound
        If optEnglish = True Then
            strResponse = MsgBox("The protocol overflow ?", vbYesNo + vbQuestion, "Cancel")
        Else
            strResponse = MsgBox("Protokols ir izpildits ?", vbYesNo + vbQuestion, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Номер последней строки "Таблицы протокола" (потеря записей)
            gProtocRowNum = 32760
            'Нажата кнопка "Да"
        Else
            'Архивирование файла "Таблицы протокола"
            WriteProtocolToArchives
        End If
            'Файл "Таблицы протокола" не требует архивирования
    Else
            'Номер первой свободной строки "Таблицы протокола"
        gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
    End If
            'Протоколирование события - "Старт системы"
    gProtocol.strProtocName = "****************"
            'Системный пароль
    gProtocol.strProtocPersonCode = txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "Restart the Acc. Syst."
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
     
End Sub

            'Вызов подменю "mnuCalendar" меню "Adjustment"
Private Sub imgCalendar_Click()
    mnuCalendar_Click
End Sub

            'Вызов подменю "mnuPersons" меню "Adjustment"
Private Sub imgPersons_Click()
    mnuPersons_Click
End Sub

            'Вызов подменю "mnuParkingInData" меню "Parking"
Private Sub imgParkingIn_Click()
    mnuParkingInData_Click
End Sub

            'Вызов подменю "mnuParkingOutData" меню "Parking"
Private Sub imgParkingOut_Click()
    mnuParkingOutData_Click
End Sub

            'Вызов подменю "mnuParkingInfoData" меню "Parking"
Private Sub imgParkingInfo_Click()
    mnuParkingInfoData_Click
End Sub

            'Вызов подменю "mnuParkingServData" меню "Parking"
Private Sub imgParkingServ_Click()
    mnuParkingServData_Click
End Sub

            'Вызов подменю "mnuAccessInData" меню "Access"
Private Sub imgAccessIn_Click()
    mnuAccessInData_Click
End Sub

            'Вызов подменю "mnuAccessOutData" меню "Access"
Private Sub imgAccessOut_Click()
    mnuAccessOutData_Click
End Sub

            'Вызов подменю "mnuAccessInfoData" меню "Access"
Private Sub imgAccessInfo_Click()
    mnuAccessInfoData_Click
End Sub

            'Вызов подменю "mnuAccessServData" меню "Access"
Private Sub imgAccessServ_Click()
    mnuAccessServData_Click
End Sub

            'Вызов подменю "mnuPrint..." меню "File"
Private Sub imgPrint_Click()
    mnuPrint_Click
End Sub

            'Вызов подменю "mnuFormProtocolBase" меню "File"
Private Sub imgProtocolBase_Click()
    mnuFormProtocolBase_Click
End Sub

            'Вызов подменю "mnuFormBookKeeperBase" меню "File"
Private Sub imgBookKeeperBase_Click()
    mnuFormBookKeeperBase_Click
End Sub

            'Вызов подменю "mnuProtocolToArchives..." меню "File"
Private Sub imgProtocArchives_Click()
    mnuProtocolToArchives_Click
End Sub

            'Вызов подменю "mnuSaveProtocol" меню "Adjustment"
Private Sub imgSaveProtocol_Click()
    mnuSaveProtocol_Click
End Sub

            'Вызов подменю "mnuSystem" меню "Adjustment"
Private Sub imgSystem_Click()
    mnuSystem_Click
End Sub

            'Вызов подменю "mnuTerminal" меню "Adjustment"
Private Sub imgTerminal_Click()
    mnuTerminal_Click
End Sub

            'Вызов подменю "mnuTime" меню "Adjustment"
Private Sub imgTime_Click()
    mnuTime_Click

End Sub

            'Вызов подменю "mnuPreprocessors" меню "Adjustment"
Private Sub imgPreprocessors_Click()
    mnuPreprocessors_Click
    
End Sub

            'Обработка вызова подменю "Preprocessors" меню "Adjustment"
Private Sub mnuPreprocessors_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer

            'Сервисная функция доступна только для "Host Computer'a"
    If gPreprocName <> "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If

             'Загрузить (не показывая) форму "frmPreprocessors"
    Load frmPreprocessors
    
            'Препроцессоры есть в системе (в "Системной таблице")
    If frmPreprocessors.cboFileName.ListCount <> 0 Then
            'Вывести на экран форму "frmPreprocessors"
            '   с уровнем модальности 1
        frmPreprocessors.Show 1
    End If
    
            'Выгрузить форму "frmPreprocessors"
    UnLoad frmPreprocessors
            'Освободить память, занимаемую выгруженной формой
    Set frmPreprocessors = Nothing
            'Восстановить стандартный курсор мыши
    frmDemo.MousePointer = 0
            'Сделать доступными элементы управления формы
    frmDemo.Enabled = True
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then frmDemo.chkDummy.SetFocus

End Sub

            'Обработка вызова подменю "Protocol to Archives..." меню "File"
Private Sub mnuProtocolToArchives_Click()
            'Полное имя файла архива (с указанием "пути" к нему)
Dim strPathFileName As String
            'Номер архивного файла
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола"
Dim lngRecordLen As Long
            'Текущий номер строки "Таблицы протокола"
Dim intRowNum As Integer
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer

            'Загрузить (не показывая) форму "frmGetFile"
    Load frmGetFile
            'Заполнить список комбинированного поля "cboFileType
    frmGetFile.cboFileType.AddItem "All files (*.*)"
    frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
    frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            'Выбрать элемент списка "Все файлы"
    frmGetFile.cboFileType.ListIndex = 0
    
            'Сформировать умалчиваемое имя архивируемого файла
            
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            'Полное имя файла (с указанием "пути" к нему)
        frmGetFile.txtFileName = gHost + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
            ' Если это Препроцессор
    Else
            'Полное имя файла (с указанием "пути" к нему)
        frmGetFile.txtFileName = gPreprocName + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
    End If
    
    
            'Вывести на экран форму "frmGetFile" с уровнем модальности 1
    frmGetFile.Show 1
            'Файл не выбран
    If frmGetFile.Tag = "" Then
            'Издать звуковой сигнал
        BeepSound
        MsgBox "The file isn't selected !"
            'Запись "Таблицы протокола" в архивный файл
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Таблицы протокола"
        lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Начальная позиция в полном имени файла (за символами "C:\")
        intSymbPos = 4
            'Найти начальную позицию собственно имени файла
        Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
        Loop
            'Удалить "старый" архивный файл, если он существует
        If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
            Kill strPathFileName
        End If
        
            'Обработка ошибок
                On Error GoTo UnDefError
            'Открыть выбранный архивный файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
    
            'Цикл по всем строкам "Таблицы протокола"
        For intRowNum = 1 To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, intRowNum, gProtocol
            'Вывести строку "Таблицы протокола" в архивный файл
            Put intFileNum, intRowNum, gProtocol
        Next
            'Закрыть выбранный  архивный файл
        Close intFileNum
            
            ' Если это Препроцессор
        If gPreprocName <> "" Then
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            strMessage = "Archive" + " " + Mid(strPathFileName, intSymbPos)
            'Отослать СООБЩЕНИЕ
            Call SendMessage(strMessage)
        End If
        
             'Закрыть "текущий" файл "Таблицы протокола"
        Close gProtocFileNum
           'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableProtocol.dat"
        
            'Начальная позиция в полном имени файла (за символами "C:\")
        intSymbPos = 4
            'Найти начальную позицию собственно имени файла
        Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
        Loop
            'Удалить  "текущий" файл "Таблицы протокола"
        If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
            Kill strPathFileName
        End If
            'Получить свободный номер файла для новой "Таблицы протокола"
        gProtocFileNum = FreeFile
            'Открыть новый файл "Таблицы протокола" для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Номер следующей "свободной" строки в новом файле "Таблицы протокола"
        gProtocRowNum = 1
            'Установить признак сохранения протокола событий в умалчиваемом файле
        mnuSaveProtocol.Checked = True
        mnuSaveProtocolAs.Checked = False
    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
    
    Exit Sub
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing

End Sub
            
            'Обработка вызова подменю "Print..." меню "File"
Private Sub mnuPrint_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
    Dim strPathFileName As String
            'Количество строк в "Базе Бухгалтерии"
    Dim intBookKeepingBaseCount As Integer
            'Количество строк в "Базе Протокола"
    Dim intProtocolBaseCount As Integer
            'Имя таблицы, выбранной для печати
    Dim strTableName As String
            'Текущий номер строки формы "frmPrintPreview"
    Dim intRowPrintNum As Integer
            'Количество строк на одной странице формы "frmPrintPreview"
    Dim intRowPrintQuan As Integer
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    Dim intRowNum As Integer
            'Текущий номер столбца таблиц ("TablePerson", "TableCalendar", "TableProtocol"
            '  "TableSystem", "TableTime", "TableTerminal")
    Dim intColNum As Integer
            'Буфер печати строки "Системной таблицы"
    Dim strTableSystem(5) As String
            'Буфер печати строки "Таблицы персон"
    Dim strTablePerson(6) As String
            'Буфер печати строки "Таблицы календаря"
    Dim strTableCalendar(8) As String
            'Буфер печати строки "Таблицы времени"
    Dim strTableTime(3) As String
            'Буфер печати строки "Таблицы терминалов"
    Dim strTableTerminal(4) As String
    
            'Загрузить (не показывая) форму "frmSelectRow"
    Load frmSelectRow
            'Инициализировать этикетку "lblColName" формы "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Table type"
    
             'Очистить список объектов
    frmSelectRow.lstSelectRow.Clear
            'Заполнение списка "lstSelectRow"
    frmSelectRow.lstSelectRow.AddItem "TableSystem"
    frmSelectRow.lstSelectRow.AddItem "TablePerson"
    frmSelectRow.lstSelectRow.AddItem "TableCalendar*"
    frmSelectRow.lstSelectRow.AddItem "TableProtocol"
    frmSelectRow.lstSelectRow.AddItem "TableTime*"
    frmSelectRow.lstSelectRow.AddItem "TableTerminal*"
    frmSelectRow.lstSelectRow.AddItem "BookKeepingBase"
    frmSelectRow.lstSelectRow.AddItem "ProtocolBase"
            'Выбрать элемент списка
    frmSelectRow.lstSelectRow.ListIndex = 0
            'Вывести на экран форму "frmSelectRow" с уровнем модальности 1
    frmSelectRow.Show 1
            'Строка не выбрана
    If frmSelectRow.Tag = "" Then
            'Издать звуковой сигнал
        BeepSound
        MsgBox "The table isn't selected !"
            'Выгрузить форму "frmSelectRow"
        UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
        Set frmSelectRow = Nothing
            'Выбрана таблица для просмотра
    Else
            'Имя таблицы, выбранной для печати
        strTableName = frmSelectRow.Tag
            'Выгрузить форму "frmSelectRow"
        UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
        Set frmSelectRow = Nothing
              
            'Текущий номер строки на странице печати
        intRowPrintNum = 1
            'Количество строк на одной странице печати
        intRowPrintQuan = gRowPrintQuan
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal") или записей баз("ProtocolBase",
            '  "BookKeepingBase")
        intRowNum = 1
            
            'Выбрать принтер "По Умолчанию - широкий"
        Set Printer = Printers(0)
            'Очистить принтер от "остатков" предыдущей печати
        Printer.EndDoc
            'Вывести 3-и пустые строки
        Printer.CurrentY = 4
            'Печать номера страницы
        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
        Printer.Print
    
            'Печать "Системной таблицы"
        If strTableName = "TableSystem" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(45); "Type"; _
            Tab(70); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
            For intRowNum = intRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
                frmTableSystem.grdTableSystem.Row = intRowNum
            'По всем столбцам "Системной таблицы"
                For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
                    frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
                    strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
                Next
            'Вывести на печать строку "Системной таблицы"
                Printer.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
                Tab(45); strTableSystem(2); Tab(70); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Системной таблицы"
                    If intRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(45); "Type"; _
                        Tab(70); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Системной таблицы"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
    
            'Печать "Таблицы персон"
        ElseIf strTableName = "TablePerson" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
            For intRowNum = intRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
                gTablePerson.Row = intRowNum
            'По всем столбцам "Таблицы персон"
                For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
                    gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
                    strTablePerson(intColNum) = gTablePerson.Text
                Next
            'Статус - Клиент Автостоянки
                If Left(Trim(strTablePerson(2)), 2) = "07" Or _
                Left(Trim(strTablePerson(2)), 2) = "05" Or _
                Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
                    strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
                ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
                Left(Trim(strTablePerson(2)), 2) = "08" Or _
                Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
                    strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
                End If
            'Вывести на печать строку "Таблицы персон"
                Printer.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
                Tab(45); strTablePerson(2); Tab(70); strTablePerson(3); Tab(95); strTablePerson(4); _
                Tab(115); strTablePerson(5)
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Таблицы персон"
                    If intRowNum < gTablePerson.Rows - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Таблицы персон"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
    
            'Печать "Таблицы календаря"
        ElseIf strTableName = "TableCalendar*" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
            Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
            Tab(115); "Sunday"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
            For intRowNum = intRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
                frmTableCalendar.grdTableCalendar.Row = intRowNum
            'По всем столбцам "Таблицы календаря"
                For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
                    frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
                    strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
                Next
            'Вывести на печать строку "Таблицы календаря"
                Printer.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
                Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
                Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Таблицы календаря"
                    If intRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
                        Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
                        Tab(115); "Sunday"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Таблицы календаря"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
        
            'Печать "Таблицы протокола"
        ElseIf strTableName = "TableProtocol" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем строкам "Таблицы протокола"
            For intRowNum = intRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                Get gProtocFileNum, intRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
                Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(45); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(85); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Таблицы протокола"
                    If intRowNum < gProtocRowNum - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Таблицы протокола"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
                    
            'Печать "Базы Протокола"
        ElseIf strTableName = "ProtocolBase" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            'Определить действительный "путь" к каталогу выполняемой программы
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            'Установка свойств элемента "Data" доступа к "Базе Протокола"
            datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
            datBase.RecordSource = "Protocol"
            
            'Определить количество записей в "Базе Протокола"
            datBase.Refresh
            datBase.Recordset.MoveLast
            intProtocolBaseCount = datBase.Recordset.RecordCount
            'Читать "Базу Протокола"
            datBase.Recordset.MoveFirst
            'Цикл по всем записям "Базы Протокола"
            For intRowNum = 1 To intProtocolBaseCount Step 1
            'Вывести текущую запись "Базы Протокола"
                Printer.Print Tab(3); datBase.Recordset.Fields("Name").Value; _
                Tab(25); datBase.Recordset.Fields("CodeOrPassword").Value; _
                Tab(45); datBase.Recordset.Fields("Status").Value; _
                Tab(70); datBase.Recordset.Fields("Time").Value; _
                Tab(85); datBase.Recordset.Fields("Date").Value; _
                Tab(100); datBase.Recordset.Fields("ReservOrNote").Value
                datBase.Recordset.MoveNext
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все записии "Базы Протокола"
                    If intRowNum < intProtocolBaseCount Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все записи "Базы Протокола"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
                    
                    'Печать "Таблицы времени"
        ElseIf strTableName = "TableTime*" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(45); "Expander"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
            For intRowNum = intRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
                frmTableTime.grdTableTime.Row = intRowNum
            'По всем столбцам "Таблицы времени"
                For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
                    frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
                    strTableTime(intColNum) = frmTableTime.grdTableTime.Text
                Next
            'Вывести на печать строку "Таблицы времени"
                Printer.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); Tab(45); strTableTime(2)
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Таблицы времени"
                    If intRowNum < frmTableTime.grdTableTime.Rows - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(45); "Expander"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Таблицы времени"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
    
            'Печать "Таблицы терминалов"
        ElseIf strTableName = "TableTerminal*" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(45); "Description"; _
            Tab(70); "Expander"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
            For intRowNum = intRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
                frmTableTerminal.grdTableTerminal.Row = intRowNum
            'По всем столбцам "Таблицы терминалов"
                For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
                    frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
                    strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
                Next
            'Вывести на печать строку "Таблицы терминалов"
                Printer.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
                Tab(45); strTableTerminal(2); Tab(70); strTableTerminal(3)
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все строки "Таблицы терминалов"
                    If intRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(45); "Description"; _
                        Tab(70); "Expander"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все строки "Таблицы терминалов"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
        
            'Печать "Базы Бухгалтерии"
        ElseIf strTableName = "BookKeepingBase" Then
            'Печать заголовков столбцов
            Printer.Print Tab(3); "Person"; Tab(25); "PersonCode"; Tab(45); "Status"; _
            Tab(55); "Time"; Tab(70); "Date"
            'Вывести пустую строку
            Printer.Print
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 4
            'Определить действительный "путь" к каталогу выполняемой программы
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            'Установка свойств элемента "Data" доступа к "Базе Бухгалтерии"
            datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
            datBase.RecordSource = "BookKeeping"
            
            'Определить количество записей в "Базе Бухгалтерии"
            datBase.Refresh
            datBase.Recordset.MoveLast
            intBookKeepingBaseCount = datBase.Recordset.RecordCount
            'Читать "Базу Бухгалтерии"
            datBase.Recordset.MoveFirst
            'Цикл по всем записям "Базы Бухгалтерии"
            For intRowNum = 1 To intBookKeepingBaseCount Step 1
            'Вывести текущую запись "Базы Бухгалтерии"
                Printer.Print Tab(3); datBase.Recordset.Fields("Person").Value; _
                Tab(25); datBase.Recordset.Fields("PersonCode").Value; _
                Tab(45); datBase.Recordset.Fields("Status").Value; _
                Tab(55); datBase.Recordset.Fields("Time").Value; _
                Tab(70); datBase.Recordset.Fields("Date").Value
                datBase.Recordset.MoveNext
            'Текущий номер строки на странице печати
                intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена
                If intRowPrintNum > intRowPrintQuan Then
            'Напечатаны не все записии "Базы Бухгалтерии"
                    If intRowNum < intBookKeepingBaseCount Then
            'Печать номера новой страницы
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            'Вывести пустую строку
                        Printer.Print
            'Печать заголовков столбцов на новой странице
                        Printer.Print Tab(3); "Person"; Tab(25); "PersonCode"; Tab(45); "Status"; _
                        Tab(55); "Time"; Tab(70); "Date"
            'Вывести пустую строку
                        Printer.Print
            'Текущий номер строки на новой странице печати
                        intRowPrintNum = 5
            'Напечатаны все записи "Базы Бухгалтерии"
                    Else
            'Завершить печать
                        Exit For
                    End If
                End If
            Next
                    
        End If
        
            'Данных для печати больше нет
        Printer.EndDoc
    End If

End Sub

            'Обработка вызова подменю "Print preview..." меню "File"
Private Sub mnuPrintPreview_Click()
            'Загрузить (не показывая) форму "frmSelectRow"
    Load frmSelectRow
            'Инициализировать этикетку "lblColName" формы "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Table type"
    
             'Очистить список объектов
    frmSelectRow.lstSelectRow.Clear
            'Заполнение списка "lstSelectRow"
    frmSelectRow.lstSelectRow.AddItem "TableSystem"
    frmSelectRow.lstSelectRow.AddItem "TablePerson"
    frmSelectRow.lstSelectRow.AddItem "TableCalendar*"
    frmSelectRow.lstSelectRow.AddItem "TableProtocol"
    frmSelectRow.lstSelectRow.AddItem "TableTime*"
    frmSelectRow.lstSelectRow.AddItem "TableTerminal*"
    frmSelectRow.lstSelectRow.AddItem "ProtocolFromArchives"
            'Выбрать элемент списка
    frmSelectRow.lstSelectRow.ListIndex = 0
            'Вывести на экран форму "frmSelectRow" с уровнем модальности 1
    frmSelectRow.Show 1
            'Строка не выбрана
    If frmSelectRow.Tag = "" Then
            'Издать звуковой сигнал
        BeepSound
        MsgBox "The table isn't selected !"
            'Выгрузить форму "frmSelectRow"
        UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
        Set frmSelectRow = Nothing
            'Выбранная таблица для просмотра не "Архив протокола"
    ElseIf frmSelectRow.Tag <> "ProtocolFromArchives" Then
            'Загрузить (не показывая) форму "frmPrintPreview"
        Load frmPrintPreview
            'Имя таблицы, выбранной для предварительной печати
        frmPrintPreview.Tag = frmSelectRow.Tag
            'Выгрузить форму "frmSelectRow"
        UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
        Set frmSelectRow = Nothing
              'Вывести на экран форму "frmPrintPreview" с уровнем модальности 1
        frmPrintPreview.Show 1
              'Выгрузить форму "frmPrintPreview"
        UnLoad frmPrintPreview
            'Освободить память, занимаемую выгруженной формой
        Set frmPrintPreview = Nothing
            'Выбранная таблица для просмотра "Архив протокола"
    ElseIf frmSelectRow.Tag = "ProtocolFromArchives" Then
            'Загрузить (не показывая) форму "frmPrintPreview"
        Load frmPrintPreview
            'Имя таблицы, выбранной для предварительной печати
        frmPrintPreview.Tag = frmSelectRow.Tag
            'Выгрузить форму "frmSelectRow"
        UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
        Set frmSelectRow = Nothing

            'Загрузить (не показывая) форму "frmGetFile"
        Load frmGetFile
            'Заполнить список комбинированного поля "cboFileType
        frmGetFile.cboFileType.AddItem "All files (*.*)"
        frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
        frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            'Выбрать элемент списка "Все файлы"
        frmGetFile.cboFileType.ListIndex = 0
            'Вывести на экран форму "frmGetFile" с уровнем модальности 1
        frmGetFile.Show 1
            'Файл не выбран
        If frmGetFile.Tag = "" Then
            'Издать звуковой сигнал
            BeepSound
            MsgBox "The file isn't selected !"
            'Чтение  "Архива протокола"
        Else
                'Полное имя файла (с указанием "пути" к нему)
            gPathFileName = frmGetFile.Tag
              'Вывести на экран форму "frmPrintPreview" с уровнем модальности 1
            frmPrintPreview.Show 1
              'Выгрузить форму "frmPrintPreview"
            UnLoad frmPrintPreview
            'Освободить память, занимаемую выгруженной формой
            Set frmPrintPreview = Nothing
        End If
            'Выгрузить форму "frmGetFile"
        UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
        Set frmGetFile = Nothing
    End If

End Sub

            'Завершение программы при вызове подменю "Exit" меню "File"
Private Sub mnuExit_Click()
            'Код возврата при сохранении "Таблицы персон"
Dim intSaveTablePerson As Integer
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            'Установлен признак внесенных изменений в "Таблицу персон"
            '  - сохранить таблицу в умалчиваемом файле
        If gChangesTablePerson = True Then _
            intSaveTablePerson = frmTablePerson.SaveTablePerson()
            ' Если имеются Препроцессоры в локальной сети
        If gNetPreprocNum > 0 Then
            'Строка передачи сообщения
            strMessage = "ExitApp"
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call SendMessage(strMessage)
        End If
            ' Если это не "Host Computer"
    ElseIf gPreprocName <> "" Then
            'Строка передачи сообщения
        strMessage = "ExitApp"
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call SendMessage(strMessage)
    End If
            
            'Протоколирование события - "Разгрузка системы"
    gProtocol.strProtocName = "################"
            'Системный пароль
    gProtocol.strProtocPersonCode = txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "Menu EXIT"
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
            
            'Нет признака запрета формирования баз Протокола и Бухгалтерии
    If gMSBase = 1 Then
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
        Call BasesConvert
    End If
            
            'Закрыть файл "Таблицы протокола"
    Close gProtocFileNum
    
            ' Освободить ссылку на объект "FlexGrid" ("Таблица Персон")
    Set gTablePerson = Nothing
            ' Освободить ссылку на объект ActiveX.EXE
    Set objTablePerson = Nothing
            
            ' Освободить ссылку на объект MSMQQueueInfo
    Set qInfoOutput = Nothing
    Set qInfoInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
    Set qQueueInput = Nothing
            ' Освободить ссылку на объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
    Set evQueue = Nothing
            ' Освободить ссылку на объект ПРИНИМАЕМОЕ СООБЩЕНИE
    Set qMsgInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
    Set qQueueOutput = Nothing
            ' Освободить ссылку на объект ПЕРЕДАВАЕМОЕ СООБЩЕНИE
    Set qMsgOutput = Nothing
    
            'Завершить программу
    End
    
End Sub

            'Обработка вызова подменю "Save Protocol" меню "Adjustment"
Private Sub mnuSaveProtocol_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы протокола"
Dim lngRecordLen As Long
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Закрыть файл, ранее открытый для сохранения "Таблицы протокола"
    Close gProtocFileNum
            'Установить признак сохранения протокола событий в умалчиваемом файле
    If mnuSaveProtocol.Checked = True Then mnuSaveProtocolAs.Checked = False
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла "Таблицы протокола"
    gProtocFileNum = FreeFile
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Номер первой свободной строки "Таблицы протокола"
    gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
            'Файл "Таблицы протокола" требует архивирования
    If gProtocRowNum > 32760 Then
            'Издать звуковой сигнал
        BeepSound
        If optEnglish = True Then
            strResponse = MsgBox("The protocol overflow ?", vbYesNo + vbQuestion, "Cancel")
        Else
            strResponse = MsgBox("Protokols ir izpildits ?", vbYesNo + vbQuestion, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Номер последней строки "Таблицы протокола" (потеря записей)
            gProtocRowNum = 32760
            'Нажата кнопка "Да"
        Else
            'Архивирование файла "Таблицы протокола"
            WriteProtocolToArchives
        End If
    End If

End Sub

            'Обработка вызова подменю "SaveProtocolAs..." меню "Adjustment"
Private Sub mnuSaveProtocolAs_Click()
            'Полное имя выбираемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы протокола"
Dim lngRecordLen As Long

             'Загрузить (не показывая) форму "frmGetFile"
    Load frmGetFile
            'Заполнить список комбинированного поля "cboFileType
    frmGetFile.cboFileType.AddItem "All files (*.*)"
    frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
    frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            'Выбрать элемент списка "Все файлы"
    frmGetFile.cboFileType.ListIndex = 0
            'Вывести на экран форму "frmGetFile" с уровнем модальности 1
    frmGetFile.Show 1
            'Файл не выбран
    If frmGetFile.Tag = "" Then
            'Издать звуковой сигнал
        BeepSound
        MsgBox "The file isn't selected !"
            'Выбран файл для сохранения "Таблицы протокола"
    Else
    
            'Закрыть файл, ранее открытый для сохранения "Таблицы протокола"
        Close gProtocFileNum
            'Установить признак сохранения протокола событий в выбираемом файле
        If mnuSaveProtocolAs.Checked = True Then mnuSaveProtocol.Checked = False
    
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
           'Вычислить длину записи (строки) "Таблицы протокола"
        lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла "Таблицы протокола"
        gProtocFileNum = FreeFile
    
            'Открыть выбираемый файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Номер первой свободной строки "Таблицы протокола"
        gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
            'Файл "Таблицы протокола" требует архивирования
        If gProtocRowNum > 32760 Then
            'Издать звуковой сигнал
            BeepSound
            If optEnglish = True Then
                MsgBox "The protocol overflow !", vbExclamation, "Error"
            Else
                MsgBox "Protokols ir izpildits !", vbExclamation, "Error"
            End If
            'Номер последней строки "Таблицы протокола" (потеря записей)
            gProtocRowNum = gProtocRowNum - 1
        End If
    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing

End Sub

            'Обработка вызова подменю "System" меню "Adjustment"
Private Sub mnuSystem_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmTableSystem"
    frmTableSystem.Visible = True
            'Установить фокус на кнопке "Correction"
    frmTableSystem.cmdCorrection.SetFocus

End Sub
            
            'Обработка вызова подменю "Persons" меню "Adjustment"
Private Sub mnuPersons_Click()
            
            'Признак необходимости сжатия "Таблицы персон" не установлен:
            '   Препроцессор использует "Таблицу персон" "Host Computer'а"
            '   - сервисная функция не доступна
    If gCompresTablPers = 0 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The function don't accessable !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmTablePerson"
    frmTablePerson.Visible = True
            'Установить фокус на кнопке "Correction"
    frmTablePerson.cmdCorrection.SetFocus

End Sub
            
            'Обработка вызова подменю "ParkingInData" меню "Parking"
Private Sub mnuParkingInData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingIn"
    Do While frmDataParkingIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataParkingIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgParkingInData"
Private Sub imgParkingInData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
    
            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuParkingInData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            'Установить признак процедуры Регистрации Клиента Автостоянки
    imgParkingInData(intIndex).Tag = 1
            'Сброс признака Активизации формы "frmDataParkingIn"
    frmDataParkingIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingIn"
    Do While frmDataParkingIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingIn.Tag = 0
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataParkingIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Регистрации отказано
    If frmDataParkingIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации
            '  Клиентов Автостоянки
        imgParkingInData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала,
            '  выполнена Регистрация Клиента Автостоянки
            '  и установлена Опция "Физическое/Логическое удаление"
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 And _
    gParkingDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия шлагбаума
        If gParkingPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия шлагбаума
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
            'Не установлена Опция "Физическое/Логическое удаление"
    Else
            'Сбросить признак процедуры Регистрации Клиента
        imgParkingInData(intIndex).Tag = 0
    End If

End Sub

            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоРегистрации (через спец "Controller") Клиента на Автостоянке
Public Function AutoParkReg(ByVal vntPersonCode As Variant, intIndex As Integer)
            
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataParkingIn - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сделать видимой форму "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            'Установить признак процедуры Регистрации Клиента Автостоянки
    imgParkingInData(intIndex).Tag = 1
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingIn.Tag = 0
            'Процедура анализа "PersonCode" при АвтоРегистрации Клиента
            '  Автостоянки через специальный "Controller"
    intAnalysisCode = frmDataParkingIn.Analysis(vntPersonCode)

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    If frmDataParkingIn.Tag = 1 Then frmDataParkingIn.Tag = 0
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataParkingIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Регистрации отказано
    If frmDataParkingIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации
            '   Клиентов Автостоянки
        imgParkingInData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1
            
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
            '  и выполнена Регистрация Клиента Автостоянки
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия шлагбаума
        If gParkingPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия шлагбаума
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
    End If

End Function
            'Обработка вызова подменю "ParkingOutData" меню "Parking"
Private Sub mnuParkingOutData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            'Сброс признака Активизации формы "frmDataParkingOut"
    frmDataParkingOut.Tag = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingOut"
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgParkingOutData"
Private Sub imgParkingOutData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuParkingOutData_Click
        Exit Sub
    End If

    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            'Установить признак процедуры Исключения Клиента Автостоянки
    imgParkingOutData(intIndex).Tag = 1
            'Сброс признака Активизации формы "frmDataParkingOut"
    frmDataParkingOut.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingOut"
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Удалении отказано
    If frmDataParkingOut.Tag = 2 Then
            'Сбросить признак процедуры
            '  Удаления Клиентов Автостоянки
        imgParkingOutData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала,
            '  выполнено Исключение Клиента Автостоянки
            '  и установлена Опция "Физическое/Логическое удаление"
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingOut.Tag = 1 And _
    gParkingDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия шлагбаума
        If gParkingPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия шлагбаума
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
            'Не установлена Опция "Физическое/Логическое удаление"
    Else
            'Сбросить признак процедуры Удаления Клиента
        imgParkingOutData(intIndex).Tag = 0
    End If

End Sub
            
            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоУдалении (через спец "Controller") Клиента на Автостоянке
Public Function AutoParkDel(ByVal vntPersonCode As Variant, intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataParkingOut - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сброс признака (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА (История ???)
    frmDataParkingIn.Tag = 0
            ' Сделать видимой форму "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            'Установить признак процедуры Исключения Клиента
    imgParkingOutData(intIndex).Tag = 1
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingOut.Tag = 0
            'Процедура анализа "PersonCode" при АвтоУдалении Клиента
            '  через специальный "Controller"
    intAnalysisCode = frmDataParkingOut.Analysis(vntPersonCode)
            
            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    If frmDataParkingOut.Tag = 1 Then frmDataParkingOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
             'В (Авто)Удалении отказано
    If frmDataParkingIn.Tag = 2 Then
            'Сбросить признак процедуры
            '  Удаления Клиентов Автостоянки
        imgParkingOutData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
            '  и выполнено Исключение Клиента Автостоянки
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingOut.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия шлагбаума
        If gParkingPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия шлагбаума
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
           gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
    End If
            
End Function
            
            'Обработка вызова подменю "ParkingInfoData" меню "Parking"
Private Sub mnuParkingInfoData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingInfo"
    frmDataParkingInfo.Visible = True
            'Сброс признака Активизации формы "frmDataParkingInfo"
    frmDataParkingInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingInfo"
    Do While frmDataParkingInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака Сжатия данных в Таблице персон
    frmDataParkingInfo.Tag = 0
             'Ждать завершения (или отказа от) Сжатия
    Do While frmDataParkingInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgParkingInfoData"
Private Sub imgParkingInfoData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuParkingInfoData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingInfo"
    frmDataParkingInfo.Visible = True
            'Установить признак процедуры получения информации об Автостоянке
    imgParkingInfoData(intIndex).Tag = 1
            'Сброс признака Активизации формы "frmDataParkingInfo"
    frmDataParkingInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingInfo"
    Do While frmDataParkingInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака Сжатия данных в "Таблице персон"
    frmDataParkingInfo.Tag = 0
             'Ждать завершения (или отказа от) Сжатия данных
    Do While frmDataParkingInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            'Сбросить признак процедуры получения информации об Автостоянке
    imgParkingInfoData(intIndex).Tag = 0

End Sub
            
            'Обработка вызова подменю "ParkingServData" меню "Parking"
Private Sub mnuParkingServData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataParkingServ"
    frmDataParkingServ.Visible = True
            'Сброс признака Активизации формы "frmDataParkingServ"
    frmDataParkingServ.Tag = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataParkingServ"
    Do While frmDataParkingServ.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака (Авто) Коррекции для данного
            '   ПЕРСОНАЛЬНОГО КОДА
    frmDataParkingServ.Tag = 0
             'Ждать завершения (или отказа от) Коррекции
    Do While frmDataParkingServ.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка вызова подменю "AccessInData" меню "Access"
Private Sub mnuAccessInData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataAccessIn"
    Do While frmDataAccessIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataAccessIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgAccessInData"
Private Sub imgAccessInData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuAccessInData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сброс признака Активизации формы "frmDataAccessIn"
    frmDataAccessIn.Tag = 0
            ' Сделать видимой форму "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            'Установить признак процедуры Регистрации Клиента
    imgAccessInData(intIndex).Tag = 1
             'Ждать завершения Активизации формы "frmDataAccessIn"
    Do While frmDataAccessIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            
            'В (Авто)Регистрации отказано
    If frmDataAccessIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации Клиентов
        imgAccessInData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала,
            '  выполнена Регистрация Клиента
            '  и установлена Опция "Физическое/Логическое удаление"
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessIn.Tag = 1 And _
    gAccessDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия турникета
        If gAccessPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия турникета
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
            'Не установлена Опция "Физическое/Логическое удаление"
    Else
            'Сбросить признак процедуры Регистрации Клиента
        imgAccessInData(intIndex).Tag = 0
    End If

End Sub

            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоРегистрации (через спец "Controller") Клиента
Public Function AutoAcceReg(ByVal vntPersonCode As Variant, intIndex As Integer)
            
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataAccessIn - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сделать видимой форму "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            'Установить признак процедуры Регистрации Клиента
    imgAccessInData(intIndex).Tag = 1
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessIn.Tag = 0
            'Процедура анализа "PersonCode" при АвтоРегистрации Клиента
            '  через специальный "Controller"
    intAnalysisCode = frmDataAccessIn.Analysis(vntPersonCode)

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    If frmDataAccessIn.Tag = 1 Then frmDataAccessIn.Tag = 0
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataAccessIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Регистрации отказано
    If frmDataAccessIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации Клиентов
        imgAccessInData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1
            
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
            '  и выполнена Регистрация Клиента
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessIn.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия турникета
        If gAccessPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия турникета
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
    End If

End Function
            
            'Обработка вызова подменю "AccessOutData" меню "Access"
Private Sub mnuAccessOutData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            'Сброс признака Активизации формы "frmDataAccessOut"
    frmDataAccessOut.Tag = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataAccessOut"
    Do While frmDataAccessOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataAccessOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgAccessOutData"
Private Sub imgAccessOutData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuAccessOutData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сброс признака Активизации формы "frmDataAccessOut"
    frmDataAccessOut.Tag = 0
            ' Сделать видимой форму "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            'Установить признак процедуры Исключения Клиента
    imgAccessOutData(intIndex).Tag = 1
             'Ждать завершения Активизации формы "frmDataAccessOut"
    Do While frmDataAccessOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataAccessOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Удалении отказано
    If frmDataAccessOut.Tag = 2 Then
            'Сбросить признак процедуры Удаления Клиентов
        imgAccessOutData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1


            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала,
            '  выполнено Исключение Клиента
            '  и установлена Опция "Физическое/Логическое удаление"
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessOut.Tag = 1 And _
    gAccessDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия турникета
        If gAccessPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия турникета
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
            'Не установлена Опция "Физическое/Логическое удаление"
    Else
            'Сбросить признак процедуры Удаления Клиента
        imgAccessOutData(intIndex).Tag = 0
    End If

End Sub

            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоУдалении (через спец "Controller") Клиента
Public Function AutoAcceDel(ByVal vntPersonCode As Variant, intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataAccessOut - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            'Установить признак процедуры Исключения Клиента
    imgAccessOutData(intIndex).Tag = 1
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessOut.Tag = 0
            'Процедура анализа "PersonCode" при АвтоУдалении Клиента
            '  через специальный "Controller"
    intAnalysisCode = frmDataAccessOut.Analysis(vntPersonCode)

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    If frmDataAccessOut.Tag = 1 Then frmDataAccessOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataAccessOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
             'В (Авто)Удалении отказано
    If frmDataAccessIn.Tag = 2 Then
            'Сбросить признак процедуры Удаления Клиентов
        imgAccessOutData(intIndex).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1


            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
            '  и выполнено Исключение Клиента
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessOut.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Не требуется ручное подтверждение открытия турникета
        If gAccessPresButton = 0 Then
            'Метка "N_?" - (зеленый фон)
            lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
            Call cmdOpen_Click(intIndex)
            'Требуется ручное подтверждение открытия турникета
        Else
            'Сделать электронную "Кнопку" временно доступной
            cmdOpen(intIndex).Enabled = True
            'Установить признак ЗАНЯТОГО "Controller'a", у которого
            '   ожидается Оранжевый индикатор на считывателе
           gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            'Метка "N_?" - (желтый фон)
            lblInform(intIndex).BackColor = vbYellow
            'Включить контроль "TimeOut" электронной "Кнопки"
            tmrButton(intIndex).Enabled = True
        End If
    End If
            
End Function
            
            'Обработка вызова подменю "AccessInfoData" меню "Access"
Private Sub mnuAccessInfoData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessInfo"
    frmDataAccessInfo.Visible = True
            'Сброс признака Активизации формы "frmDataAccessInfo"
    frmDataAccessInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataAccessInfo"
    Do While frmDataAccessInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака Сжатия данных в Таблице персон
    frmDataAccessInfo.Tag = 0
             'Ждать завершения (или отказа от) Сжатия
    Do While frmDataAccessInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgAccessInfoData"
Private Sub imgAccessInfoData_Click(intIndex As Integer)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuAccessInfoData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessInfo"
    frmDataAccessInfo.Visible = True
            'Установить признак процедуры получения информации
    imgAccessInfoData(intIndex).Tag = 1
            'Сброс признака Активизации формы "frmDataAccessInfo"
    frmDataAccessInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataAccessInfo"
    Do While frmDataAccessInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака Сжатия данных в "Таблице персон"
    frmDataAccessInfo.Tag = 0
             'Ждать завершения (или отказа от) Сжатия данных
    Do While frmDataAccessInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            'Сбросить признак процедуры получения информации
    imgAccessInfoData(intIndex).Tag = 0

End Sub
            
            'Обработка вызова подменю "AccessServData" меню "Access"
Private Sub mnuAccessServData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataAccessServ"
    frmDataAccessServ.Visible = True
            'Сброс признака Активизации формы "frmDataAccessServ"
    frmDataAccessServ.Tag = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataAccessServ"
    Do While frmDataAccessServ.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака (Авто) Коррекции для данного
            '   ПЕРСОНАЛЬНОГО КОДА
    frmDataAccessServ.Tag = 0
             'Ждать завершения (или отказа от) Коррекции
    Do While frmDataAccessServ.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка вызова подменю "EmployeInData" меню "Employe"
Private Sub mnuEmployeInData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataEmployeIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgEmployeInData"
Private Sub imgEmployeInData_Click()
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            
            'Сервисная функция доступна только для "Host Computer'a"
    If gPreprocName <> "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuEmployeInData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            'Установить признак процедуры Регистрации Служащего
    imgEmployeInData.Tag = 1
            'Сброс признака Активизации формы "frmDataEmployeIn"
    frmDataEmployeIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака Активизации формы "frmDataEmployeIn"
    If frmDataEmployeIn.Tag = 1 Then frmDataEmployeIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Регистрации отказано
    If frmDataEmployeIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации Служащего
        imgEmployeInData.Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

End Sub

            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоРегистрации (через спец "Controller") Служащего
Public Function AutoEmplReg(ByVal vntPersonCode As Variant)
            
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataEmployeIn - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
          'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            'Сделать видимой форму "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            'Установить признак процедуры Регистрации Служащего
    imgEmployeInData.Tag = 1
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataEmployeIn.Tag = 0
            'Процедура анализа "PersonCode" при АвтоРегистрации Служащего
            '  через специальный "Controller"
    intAnalysisCode = frmDataEmployeIn.Analysis(vntPersonCode)

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    If frmDataEmployeIn.Tag = 1 Then frmDataEmployeIn.Tag = 0
             'Ждать завершения (или отказа от) Регистрации
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Регистрации отказано
    If frmDataEmployeIn.Tag = 2 Then
            'Сбросить признак процедуры Регистрации Служащего
        imgEmployeInData.Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

End Function
            
            'Обработка вызова подменю "EmployeOutData" меню "Employe"
Private Sub mnuEmployeOutData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            'Сброс признака Активизации формы "frmDataEmployeOut"
    frmDataEmployeOut.Tag = 0
            'Сброс признака (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
    frmDataEmployeIn.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур) и режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно) - выход из процедуры
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataEmployeOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataEmployeOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgEmployeOutData"
Private Sub imgEmployeOutData_Click()
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            
            'Сервисная функция доступна только для "Host Computer'a"
    If gPreprocName <> "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuEmployeOutData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            'Установить признак процедуры Исключения Служащего
    imgEmployeOutData.Tag = 1
            'Сброс признака Активизации формы "frmDataEmployeOut"
    frmDataEmployeOut.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака Активизации формы "frmDataEmployeOut"
    If frmDataEmployeOut.Tag = 1 Then frmDataEmployeOut.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'В (Авто)Удалении отказано
    If frmDataEmployeOut.Tag = 2 Then
            'Сбросить признак процедуры Удаления Служащих
        imgEmployeOutData.Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

End Sub

            'Инициализация ввода ПЕРСОНАЛЬНОГО КОДА и других данных
            '  при АвтоУдалении (через спец "Controller") Служащего
Public Function AutoEmplDel(ByVal vntPersonCode As Variant)
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при анализе  "PersonCode"
Dim intAnalysisCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataEmployeOut - выйти из функции
    If frmDemo.Enabled = False Then Exit Function
    
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            'Установить признак процедуры Исключения Служащего
    imgEmployeOutData.Tag = 1
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    frmDataEmployeOut.Tag = 0
            'Процедура анализа "PersonCode" при АвтоУдалении Служащего
            '  через специальный "Controller"
    intAnalysisCode = frmDataEmployeOut.Analysis(vntPersonCode)

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then Exit Function

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сброс признака (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    If frmDataEmployeOut.Tag = 1 Then frmDataEmployeOut.Tag = 0
             'Ждать завершения (или отказа от) Удаления
    Do While frmDataEmployeOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
             'В (Авто)Удалении отказано
    If frmDataEmployeIn.Tag = 2 Then
            'Сбросить признак процедуры Удаления Служащего
        imgEmployeOutData.Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1
            
End Function
            
            'Обработка вызова подменю "EmployeInfoData" меню "Employe"
Private Sub mnuEmployeInfoData_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeInfo"
    frmDataEmployeInfo.Visible = True
            'Сброс признака Активизации формы "frmDataEmployeInfo"
    frmDataEmployeInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop

End Sub
            
            'Обработка "щелчка" мыши на элементе "imgEmployeInfoData"
Private Sub imgEmployeInfoData_Click()
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer

            'Установлена Опция разделения времени
            '   (параллельное выполнение процедур)
    If gTimeShare = 1 Then
        Call mnuEmployeInfoData_Click
        Exit Sub
    End If

            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmDataEmployeInfo"
    frmDataEmployeInfo.Visible = True
            'Установить признак процедуры получения информации
    imgEmployeInfoData.Tag = 1
            'Сброс признака Активизации формы "frmDataEmployeInfo"
    frmDataEmployeInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Сброс признака Активизации формы "frmDataEmployeInfo"
    If frmDataEmployeInfo.Tag = 1 Then frmDataEmployeInfo.Tag = 0
             'Ждать завершения Активизации формы "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1

            'Сбросить признак процедуры получения информации
    imgEmployeInfoData.Tag = 0

End Sub
            
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
Public Sub BasesConvert()
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer


            'Установить курсор мыши "Песочные часы" над Главной формой
    frmDemo.MousePointer = vbHourglass

            ' Запретить прием/передачу информации для терминалов
    If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
    If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
    If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
    If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            'Запретить опрос "Controller'ов" по таймерам
    gTermContr = 0
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
    

            'Если это "Host Computer" и есть Препроцессоры в локальной сети
    If gPreprocName = "" And gNetPreprocNum > 0 Then
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
            '  для всех Препроцессоров и "Host Computer'a" - общий сбор
        Call frmPreprocessors.BasesConvert
            
            'Если это Препроцессор или нет Препроцессоров в локальной сети
    ElseIf gPreprocName <> "" Or gNetPreprocNum = 0 Then
            'Вызов подменю "mnuFormProtocolBase" меню "File"
        mnuFormProtocolBase_Click
            'Вызов подменю "mnuFormBookKeeperBase" меню "File"
        mnuFormBookKeeperBase_Click
            
    End If
    
            
            ' Разрешить прием/передачу информации для отмеченных терминалов
            ' Логические порты "свободны" - могут обрабатывать данные
            ' отмеченных терминалов
    If chkTerm(0).Value = 1 Then
        If prtPortC(0).PortOpen = False Then prtPortC(0).PortOpen = True
        prtPortC(0).Tag = 0
    End If
    If chkTerm(1).Value = 1 Then
        If prtPortC(1).PortOpen = False Then prtPortC(1).PortOpen = True
        prtPortC(1).Tag = 0
    End If
    If chkTerm(2).Value = 1 Then
        If prtPortC(2).PortOpen = False Then prtPortC(2).PortOpen = True
        prtPortC(2).Tag = 0
    End If
    If chkTerm(3).Value = 1 Then
        If prtPortC(3).PortOpen = False Then prtPortC(3).PortOpen = True
        prtPortC(3).Tag = 0
    End If
            'Разрешить опрос "Controller'ов" по таймерам
    gTermContr = 1
        
            'Восстановить стандартный курсор мыши над Главной формой
    frmDemo.MousePointer = 0
            
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True

End Sub
            
            'Обработка вызова подменю "Calendar" меню "Adjustment"
Private Sub mnuCalendar_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmTableCalendar"
    frmTableCalendar.Visible = True
            'Установить фокус на кнопке "Correction"
    frmTableCalendar.cmdCorrection.SetFocus
    

End Sub
            
            'Обработка вызова подменю "Timer" меню "Adjustment"
Private Sub mnuTime_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmTableTime"
    frmTableTime.Visible = True
            'Установить фокус на кнопке "Correction"
    frmTableTime.cmdCorrection.SetFocus
    

End Sub
            
            'Обработка вызова подменю "Terminal" меню "Adjustment"
Private Sub mnuTerminal_Click()
            'Сделать недоступной форму "frmDemo"
    frmDemo.Enabled = False
            ' Сделать видимой форму "frmTableTerminal"
    frmTableTerminal.Visible = True
            'Установить фокус на кнопке "Correction"
    frmTableTerminal.cmdCorrection.SetFocus
    

End Sub

            ' Выбрана опция "English"
Private Sub optEnglish_Click()
            ' Текущий ндекс языка общения не "English"
    If fraFlag.Tag <> 0 Then
            ' Установить новый индекс языка общения
    fraFlag.Tag = 0
            ' Вызов процедуры изменения языка общения
    UpdateLanguage
    End If
    
End Sub
            ' Выбрана опция "Latvian"
Private Sub optLatvian_Click()
            ' Текущий индекс языка общения не "Latvian"
    If fraFlag.Tag <> 1 Then
            ' Установить новый индекс языка общения
    fraFlag.Tag = 1
            ' Вызов процедуры изменения языка общения
    UpdateLanguage
    End If
    
End Sub

            ' Выбрана опция "Russian"
Private Sub optRussian_Click()
            ' Текущий индекс языка общения не "Russian"
    If fraFlag.Tag <> 2 Then
            ' Установить новый индекс языка общения
    fraFlag.Tag = 2
            ' Вызов процедуры изменения языка общения
    UpdateLanguage
    End If
    
End Sub
            ' Процедура изменения языка общения
Public Sub UpdateLanguage()
            'Новый индекс языка общения
    intLang = fraFlag.Tag
            'Отобразить флаг
    If intLang = 0 Then
    imgEnglish.Visible = True
    imgLatvian.Visible = False
    imgRussian.Visible = False
    End If
    
    If intLang = 1 Then
    imgEnglish.Visible = False
    imgLatvian.Visible = True
    imgRussian.Visible = False
    End If
    
    If intLang = 2 Then
    imgEnglish.Visible = False
    imgLatvian.Visible = False
    imgRussian.Visible = True
    End If
    
                ' Инициализировать массив надписей
    chkSetup.Caption = aCaption(intLang, 0)
    fraFlag.Caption = aCaption(intLang, 1)
    optEnglish.Caption = aCaption(intLang, 2)
    optLatvian.Caption = aCaption(intLang, 3)
    optRussian.Caption = aCaption(intLang, 4)
    lblTerminals.Caption = aCaption(intLang, 5)
    chkTerm(0).Caption = aCaption(intLang, 6)
    chkTerm(1).Caption = aCaption(intLang, 7)
    chkTerm(2).Caption = aCaption(intLang, 8)
    chkTerm(3).Caption = aCaption(intLang, 9)
    lblPhoto.Caption = aCaption(intLang, 10)
    cmdExit.Caption = aCaption(intLang, 11)
    optAutomatic.Caption = aCaption(intLang, 12)
    lblOpen.Caption = aCaption(intLang, 13)
    cmdOpen(0).Caption = aCaption(intLang, 14)
    cmdOpen(1).Caption = aCaption(intLang, 15)
    cmdOpen(2).Caption = aCaption(intLang, 16)
    cmdOpen(3).Caption = aCaption(intLang, 17)
    optManual.Caption = aCaption(intLang, 18)
    lblPassword.Caption = aCaption(intLang, 19)
    mnuFile.Caption = aCaption(intLang, 20)
    mnuAdjustment.Caption = aCaption(intLang, 21)
    fraControl.Caption = aCaption(intLang, 22)
            ' Инициализировать массив "всплывающих" подсказок
    txtPassword.ToolTipText = aComment(intLang, 0)
    chkSetup.ToolTipText = aComment(intLang, 1)
    optEnglish.ToolTipText = aComment(intLang, 2)
    optLatvian.ToolTipText = aComment(intLang, 3)
    optRussian.ToolTipText = aComment(intLang, 4)
    chkTerm(0).ToolTipText = aComment(intLang, 5)
    chkTerm(1).ToolTipText = aComment(intLang, 6)
    chkTerm(2).ToolTipText = aComment(intLang, 7)
    chkTerm(3).ToolTipText = aComment(intLang, 8)
    chkPhoto(0).ToolTipText = aComment(intLang, 9)
    chkPhoto(1).ToolTipText = aComment(intLang, 10)
    chkPhoto(2).ToolTipText = aComment(intLang, 11)
    chkPhoto(3).ToolTipText = aComment(intLang, 12)
    cmdExit.ToolTipText = aComment(intLang, 13)
    optAutomatic.ToolTipText = aComment(intLang, 14)
    optManual.ToolTipText = aComment(intLang, 15)
'    chkManual(1).ToolTipText = aComment(intLang, 16)
'    chkManual(2).ToolTipText = aComment(intLang, 17)
'    chkManual(3).ToolTipText = aComment(intLang, 18)
    cmdOpen(0).ToolTipText = aComment(intLang, 19)
    cmdOpen(1).ToolTipText = aComment(intLang, 20)
    cmdOpen(2).ToolTipText = aComment(intLang, 21)
    cmdOpen(3).ToolTipText = aComment(intLang, 22)
            
 End Sub
            
            'Процедура подачи длительного звукового сигнала
Public Sub BeepSound()
            'Рабочий счетчик
Dim intCount As Integer
            'Включить подачу звукового сигнала
    For intCount = 1 To gBeepSound
        Beep
    Next
    
End Sub

            'Обработка события - поступило собщение в
            '  ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
Private Sub qEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Код возврата при вызове функции "Shell"
Dim vntShell As Variant
            'Длина строки "Системной таблицы", "Таблицы протокола" или "Таблицы календаря"
Dim lngRecordLen As Long
            'Cвободный номер файла
Dim intFileNum As Integer
            'Cистемное время
Dim strTime As String
Dim vntTime As Variant
            'Cистемная дата
Dim strDate As String
Dim vntDate As Variant

            'Переменная-объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
    Set evQueue = qQueueInput
            'Принять СООБЩЕНИЕ
    Set qMsgInput = evQueue.Receive(, , , 0)
            
            'Если "Host Computer" принял ЗАПРОС о количестве
            '  свободных мест на Автостоянке
    If Mid(qMsgInput.Body, 4, 15) = "ParkFreePlaces " And _
    gPreprocName = "" Then
            ' Если имеется дисплей-указатель количества свободных мест
        If gParkingPlaceNum <> 0 Then
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            qMsgOutput.Body = "ParkFreePlaces=" + CStr(gParkFreePlaces)
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            '   конкретного Препроцессора, приславшего СООБЩЕНИЕ
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            qMsgInput.Label + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
            qQueueOutput.Close
        End If
            
            'Если "Host Computer" принял ЗАПРОС о количестве
            '  свободных мест на Предприятии
    ElseIf Left(qMsgInput.Body, 15) = "AcceFreePlaces " And _
    gPreprocName = "" Then
            ' Если имеется дисплей-указатель количества свободных мест
        If gAccessPlaceNum <> 0 Then
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
            qMsgOutput.Body = "AcceFreePlaces=" + CStr(gParkFreePlaces)
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            '   конкретного Препроцессора, приславшего СООБЩЕНИЕ
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            qMsgInput.Label + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
            qQueueOutput.Close
        End If
            
            'Если принято СООБЩЕНИЕ о необходимости синхронизации количества
            '  свободных мест и это не был ЗАПРОС от Препроцессора
    ElseIf Mid(qMsgInput.Body, 5, 10) = "FreePlaces" And _
    Mid(qMsgInput.Body, 15, 1) <> " " Then
            'Вызов функции отображения информации на дисплее
        Call Display(qMsgInput.Body)
            
            'Если "Host Computer" принял СООБЩЕНИЕ об
            '  архивировании протокола одним из Препроцессоров
    ElseIf Left(qMsgInput.Body, 7) = "Archive" And gPreprocName = "" Then
            'Копирование архива из Препроцессора в "Host Computer"
'        Call ArchiveCopy(qMsgInput.Body)
            
            'Если "Host Computer" принял СООБЩЕНИЕ о необходимости
            '  синхронизации времени от одного из Препроцессоров
    ElseIf Left(qMsgInput.Body, 4) = "Time" And gPreprocName = "" Then
        vntTime = Time
        strTime = CStr(vntTime)
        vntDate = Date
        strDate = CStr(vntDate)
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = "Time " + Trim(strTime) + "||" + Trim(strDate)
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            '   конкретного Препроцессора, приславшего СООБЩЕНИЕ
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        qMsgInput.Label + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
        qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
        qQueueOutput.Close
        
            'Поступила команда от "Host Computer'a" на
            '  синхронизацию времени
    ElseIf Left(qMsgInput.Body, 4) = "Time" And gPreprocName <> "" Then
        strTime = Mid(qMsgInput.Body, 6, 8)
        vntTime = strTime
        Time = Format(vntTime, "hh:mm:ss")
        strDate = Mid(qMsgInput.Body, 16)
        vntDate = strDate
        Date = Format(vntDate, "dd.MM.yyyy.")
            'Добавление спецсимвола "~" и удаление "лишних" символов перед записью
            '  строки в файл "Таблицы протокола"
        qMsgInput.Body = "~" + Mid(qMsgInput.Body, 6)
        
            'Поступила команда от "Host Computer'a"
            '  на рестарт приложения
    ElseIf Left(qMsgInput.Body, 8) = "StartApp" Then
            
            'Строка передачи сообщения
        strMessage = "ExitApp"
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call SendMessage(strMessage)
            
        gProtocol.strProtocName = "FROM:"
            'Системный пароль
        gProtocol.strProtocPersonCode = qMsgInput.Label
            'Статус
        gProtocol.strProtocStatus = "?? - MSMQ"
            'Время
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = qMsgInput.Body
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            'Закрыть файл "Таблицы протокола"
        Close gProtocFileNum
            
            ' Освободить ссылку на объект "FlexGrid" ("Таблица Персон")
        Set gTablePerson = Nothing
            ' Освободить ссылку на объект ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' Освободить ссылку на объект MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set qQueueInput = Nothing
            ' Освободить ссылку на объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set evQueue = Nothing
            ' Освободить ссылку на объект ПРИНИМАЕМОЕ СООБЩЕНИE
        Set qMsgInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        Set qQueueOutput = Nothing
            ' Освободить ссылку на объект ПЕРЕДАВАЕМОЕ СООБЩЕНИE
        Set qMsgOutput = Nothing
        
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            'Определить "путь" к Препроцессору EXE-модуля
        strPathFileName = strPathFileName + gModuleStartUp
            'Очистить "Clipboard"
        Clipboard.Clear
            'Занести в "Clipboard" данные для Программы-Стартера
        Clipboard.SetText strPathFileName
            
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            'Запустить Программу-Стартер с невидимой формой
        strPathFileName = strPathFileName + "StartModule.exe"
        
        vntShell = Shell(strPathFileName, 0)
            'Запуск Программы-Стартера выполнен -
            '  завершить текущее приложение
        End
    
            'Поступила команда от "Host Computer'a"
            '  на завершение выполнения приложения
    ElseIf Left(qMsgInput.Body, 7) = "StopApp" Then
            
            'Строка передачи сообщения
        strMessage = "ExitApp"
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call SendMessage(strMessage)
            
        gProtocol.strProtocName = "FROM:"
            'Системный пароль
        gProtocol.strProtocPersonCode = qMsgInput.Label
            'Статус
        gProtocol.strProtocStatus = "?? - MSMQ"
            'Время
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = qMsgInput.Body
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            'Закрыть файл "Таблицы протокола"
        Close gProtocFileNum
            
            ' Освободить ссылку на объект "FlexGrid" ("Таблица Персон")
        Set gTablePerson = Nothing
            ' Освободить ссылку на объект ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' Освободить ссылку на объект MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set qQueueInput = Nothing
            ' Освободить ссылку на объект ОЧЕРЕДЬ-СОБЫТИЕ
            ' ПРИНИМАЕМЫХ СООБЩЕНИЙ
        Set evQueue = Nothing
            ' Освободить ссылку на объект ПРИНИМАЕМОЕ СООБЩЕНИE
        Set qMsgInput = Nothing
            ' Освободить ссылку на объект
            '  ОЧЕРЕДЬ ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        Set qQueueOutput = Nothing
            ' Освободить ссылку на объект ПЕРЕДАВАЕМОЕ СООБЩЕНИE
        Set qMsgOutput = Nothing
        
            'Завершить текущее приложение
        End
    
            'Поступило СООБЩЕНИЕ о необходимости добавления в "Таблицу персон"
            '  строки с заданным полями
    ElseIf Left(qMsgInput.Body, 4) = "Reg " Then
            'Добавление строки с заданным полями в "Таблицу персон"
            '  по сообщению MSMQ, полученому по сети
        Call frmTablePerson.MSMQReg(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
               
            'Поступило СООБЩЕНИЕ о необходимости удаления из "Таблицы персон"
            '  строки с заданным персональным кодом
    ElseIf Left(qMsgInput.Body, 4) = "Del " Then
            'Удаление (ЛОГИЧЕСКОЕ) строки с заданным персональным кодом
            '  из "Таблицы персон" по сообщению MSMQ, полученому по сети
        Call frmTablePerson.MSMQDel(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
               
            'Поступило СООБЩЕНИЕ о необходимости коррекции "Таблицы персон"
            '  строки с заданным полями
    ElseIf Left(qMsgInput.Body, 4) = "Cor " Then
            'Добавление строки с заданным полями в "Таблицу персон"
            '  по сообщению MSMQ, полученому по сети
        Call frmTablePerson.MSMQCor(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
            
            'Поступило СООБЩЕНИЕ о необходимости добавления в "Таблицу информации"
            '  строки с заданным полями
    ElseIf Left(qMsgInput.Body, 8) = "RegInfo " Then
            'Добавление строки с заданным полями в "Таблицу информации"
            '  по сообщению MSMQ, полученому по сети
'        Call frmTableInfo.MSMQReg(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
               
            'Поступило СООБЩЕНИЕ о необходимости удаления из "Таблицы информации"
            '  строки с заданным персональным кодом
    ElseIf Left(qMsgInput.Body, 8) = "DelInfo " Then
            'Удаление (ЛОГИЧЕСКОЕ) строки с заданными персональным кодом и именем
            '  из "Таблицы информации" по сообщению MSMQ, полученому по сети
'        Call frmTableInfo.MSMQDel(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
               
            'Поступило СООБЩЕНИЕ о необходимости коррекции в "Таблицы информации"
            '  строки с заданным полями
    ElseIf Left(qMsgInput.Body, 8) = "CorInfo " Then
            'Добавление строки с заданным полями в "Таблицу информации"
            '  по сообщению MSMQ, полученому по сети
'        Call frmTableInfo.MSMQCor(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
            
            'Поступило СООБЩЕНИЕ о двойном Входе/Выходе(Въезде/Выезде)
    ElseIf Left(qMsgInput.Body, 9) = "Error Inp" Or _
    Left(qMsgInput.Body, 9) = "Error Out" Then
            'Игнорировать - не обрабатывать и не отображать СООБЩЕНИЕ
        GoTo IgnoreMSMQ
            
            'Иначе
    Else
            'Строка ПРИНИМАЕМЫХ метки и текста СООБЩЕНИЯ
        strMsgInput = qMsgInput.Label + " || " + qMsgInput.Body
            'Издать звуковой сигнал
        BeepSound
            'Отображение ПРИНЯТЫХ метки и текста СООБЩЕНИЯ
            '  в метке ПРИНЯТЫХ СООБЩЕНИЙ
        lblMessageInput.Caption = "FROM: " + strMsgInput
            'Сделать видимой метку сообщения
        lblMessageInput.Visible = True
    End If
            
            'Протоколирование события - обработка сообщения MSMQ
    gProtocol.strProtocName = "FROM:"
            'Системный пароль
    gProtocol.strProtocPersonCode = qMsgInput.Label
            'Статус
    gProtocol.strProtocStatus = "?? - MSMQ"
            'Время
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = qMsgInput.Body
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
            
IgnoreMSMQ:
            'Восстановить ?? связь с объектом MSMQEvent
    qQueueInput.EnableNotification qEvent

End Sub
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
Public Function SendMessage(strMessage As String)
            'Рабочий счетчик
Dim intCount As Integer
            
            'Если передается СООБЩЕНИЕ об изменении количества
            '  свободных мест  и это не ЗАПРОС от Препроцессора
    If Mid(strMessage, 5, 10) = "FreePlaces" And _
    Mid(strMessage, 15, 1) <> " " Then
            'Вызов функции отображения информации на дисплее
        Call Display(strMessage)
    End If
            
            'Если есть Препроцессоры в локальной сети
    If gNetPreprocNum > 0 And Not (qMsgOutput Is Nothing) Then
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = strMessage
            'По всем элементам массива Имен Процессоров локальной сети
        For intCount = 1 To gNetPreprocNum
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            gSocketNet(intCount) + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
            qQueueOutput.Close
        Next
        
    End If
    
End Function
            
            'Процедура вывода информации на дисплей-указатель
Public Function Display(strMessage As String)
            'Переменная-строка информации для дисплея
Dim strDisplay As String
            'Текущее время
Dim strTime As String
            'Часы
Dim intHour As Integer
            'Минуты
Dim intMinute As Integer
            
            ' Если имеется дисплей-указатель количества свободных
            '  мест на Автостоянке или на Предприятии
    If gParkingPlaceNum <> 0 Or gAccessPlaceNum <> 0 Then
            'Инициализация переменной (количества свободных мест)
        If gParkingPlaceNum > 0 Then
            strDisplay = CStr(gParkFreePlaces)
        ElseIf gAccessPlaceNum > 0 Then
            strDisplay = CStr(gAcceFreePlaces)
        End If
            'Если принято СООБЩЕНИЕ об увеличении количества
            '  свободных мест на Автостоянке
        If strMessage = "ParkFreePlaces+1" And _
        gParkFreePlaces < gParkingPlaceNum Then
            'Есть технологический перерыв в работе Автостоянки
            If Not (Left(gDefaultParkTime, 2) = "00" And _
            Mid(gDefaultParkTime, 4, 2) = "00" And _
            Mid(gDefaultParkTime, 7, 2) = "24" And _
            Mid(gDefaultParkTime, 10, 2) = "00") Then
            'Текущее время
                strTime = Format(Now, "h:mm:ss")
            'Часы
                intHour = Hour(strTime)
            'Минуты
                intMinute = Minute(strTime)
            'Рабочее время (не технологический перерыв)
                If ((Left(gDefaultParkTime, 2) < intHour Or _
                Left(gDefaultParkTime, 2) = intHour And _
                Mid(gDefaultParkTime, 4, 2) <= intMinute) And _
                Mid(gDefaultParkTime, 7, 2) > intHour) Then
            'Увеличение количества свободных мест
                    gParkFreePlaces = gParkFreePlaces + 1
                    strDisplay = CStr(gParkFreePlaces)
                End If
            'Нет технологического перерыва в работе Автостоянки
            Else
            'Увеличение количества свободных мест
                gParkFreePlaces = gParkFreePlaces + 1
                strDisplay = CStr(gParkFreePlaces)
            End If
            'Если принято СООБЩЕНИЕ об уменьшении количества
            '  свободных мест на Автостоянке
        ElseIf strMessage = "ParkFreePlaces-1" And _
        gParkFreePlaces > 0 Then
            'Уменьшение количества свободных мест
            gParkFreePlaces = gParkFreePlaces - 1
            strDisplay = CStr(gParkFreePlaces)
            'Если принято СООБЩЕНИЕ о количестве
            '  свободных мест на Автостоянке
        ElseIf Left(strMessage, 15) = "ParkFreePlaces=" Then
            'Установить количество свободных мест
            gParkFreePlaces = Mid(strMessage, 16)
            'Если принято СООБЩЕНИЕ об увеличении количества
            '  свободных мест на Предприятии
        ElseIf strMessage = "AcceFreePlaces+1" And _
        gAcceFreePlaces < gAccessPlaceNum Then
            'Есть технологический перерыв в работе Предприятия
            If Not (Left(gDefaultAcceTime, 2) = "00" And _
            Mid(gDefaultAcceTime, 4, 2) = "00" And _
            Mid(gDefaultAcceTime, 7, 2) = "24" And _
            Mid(gDefaultAcceTime, 10, 2) = "00") Then
            'Текущее время
                strTime = Format(Now, "h:mm:ss")
            'Часы
                intHour = Hour(strTime)
            'Минуты
                intMinute = Minute(strTime)
            'Рабочее время (не технологический перерыв)
                If ((Left(gDefaultAcceTime, 2) < intHour Or _
                Left(gDefaultAcceTime, 2) = intHour And _
                Mid(gDefaultAcceTime, 4, 2) <= intMinute) And _
                Mid(gDefaultAcceTime, 7, 2) > intHour) Then
            'Увеличение количества свободных мест
                    gAcceFreePlaces = gAcceFreePlaces + 1
                    strDisplay = CStr(gAcceFreePlaces)
                End If
            'Нет технологического перерыва в работе Предприятия
            Else
            'Увеличение количества свободных мест
                gAcceFreePlaces = gAcceFreePlaces + 1
                strDisplay = CStr(gAcceFreePlaces)
            End If
            'Если принято СООБЩЕНИЕ об уменьшении количества
            '  свободных мест на Предприятии
        ElseIf strMessage = "AcceFreePlaces-1" And _
        gAcceFreePlaces > 0 Then
            'Уменьшение количества свободных мест
            gAcceFreePlaces = gAcceFreePlaces - 1
            strDisplay = CStr(gAcceFreePlaces)
            'Если принято СООБЩЕНИЕ о количестве
            '  свободных мест на Предприятии
        ElseIf Left(strMessage, 15) = "AcceFreePlaces=" Then
            'Установить количество свободных мест
            gAcceFreePlaces = Mid(strMessage, 16)
            'Если принято непонятное СООБЩЕНИЕ об изменении
            '  количествa свободных мест или нарушены допустимые границы
        Else
            'Восстановить количество свободных мест в допустимых границах
            If gParkFreePlaces > gParkingPlaceNum Then
                gParkFreePlaces = gParkingPlaceNum
                strDisplay = CStr(gParkFreePlaces)
            ElseIf gParkFreePlaces < 0 Then
                gParkFreePlaces = 0
                strDisplay = CStr(gParkFreePlaces)
            ElseIf gAcceFreePlaces > gAccessPlaceNum Then
                gAcceFreePlaces = gAccessPlaceNum
                strDisplay = CStr(gAcceFreePlaces)
            ElseIf gAcceFreePlaces < 0 Then
                gAcceFreePlaces = 0
                strDisplay = CStr(gAcceFreePlaces)
            'Восстановление невозможно или непонятное СООБЩЕНИЕ
            Else
            'Инициализация Дисплея - ВЕСЬ ЧЕРНЫЙ
                strDisplay = Chr(CLng(CByte(161))) + _
                Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
                Chr(CLng(CByte(0)))
            'Послать данные на Дисплей
                prtPortDisplay.Output = strDisplay
            'Ждать завершения передачи данных на Дисплей
                Do
                Loop Until prtPortDisplay.OutBufferCount = 0
                Exit Function
            End If
        End If
        
            'Уменьшение отображаемого на дисплее количества свободных мест
        If strDisplay - gDisplayDiscount >= 0 Then _
        strDisplay = CStr(strDisplay - gDisplayDiscount)
        
            'Подготовка вывода информации на дисплей
        If strDisplay = "0" Then
            strDisplay = "000"
        ElseIf strDisplay < 10 Then
            strDisplay = "00" + strDisplay
        ElseIf strDisplay < 100 Then
            strDisplay = "0" + strDisplay
        End If
        strDisplay = Chr(CLng(CByte(1))) + _
            Chr(CLng(CByte(Mid(strDisplay, 1, 1)))) + _
                Chr(CLng(CByte(Mid(strDisplay, 2, 1))) * 16 + _
                CLng(CByte(Mid(strDisplay, 3, 1)))) + _
                    Chr(CLng(CByte(Mid(strDisplay, 1, 1))) + _
                    CLng(CByte(Mid(strDisplay, 2, 1))) * 16 + _
                    CLng(CByte(Mid(strDisplay, 3, 1))))
        
            'Послать данные на Дисплей
        prtPortDisplay.Output = strDisplay
            'Ждать завершения передачи данных на Дисплей
        Do
        Loop Until prtPortDisplay.OutBufferCount = 0
        
    End If
    
End Function
            
            'Обработка события TimeOut для "Relay"
Private Sub tmrRelay_Timer()
            
            'Выключить контроль времени
    tmrRelay.Enabled = False
            'Установить признак события TimeOut
    tmrRelay.Tag = 1

End Sub

            'Обработка события TimeOut для "Controller'a"
Private Sub tmrTimeOut_Timer(intIndex As Integer)
            
            'Протоколирование события - "TimeOut"
    gProtocol.strProtocName = "$$$"
            'Системный пароль
    gProtocol.strProtocPersonCode = ""
            'Статус
    gProtocol.strProtocStatus = ""
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "COMMAND TimeOut"
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
            
            'Выключить контроль времени
    tmrTimeOut(intIndex).Enabled = False
            'Установить признак события TimeOut
    tmrTimeOut(intIndex).Tag = 1
    
End Sub
            
            'Формирование команд ЧТЕНИЕ ПЕРСОНАЛЬНОГО КОДА
Private Sub tmrTermContr_Timer()
            'Индекс элемента в массиве элементов управления форм
Static intControlIndex As Integer
            'Номер элемента в массиве "Таблицы терминалов"
Dim intRequest As Integer
            'Рабочий счетчик
Static intCount As Integer
            
            'Запрет опроса терминалов
    If gTermContr = 0 Then Exit Sub
            
            'Порт с текущим номером закрыт - перейдти к следующему порту
    Do While prtPortC(intControlIndex).PortOpen = False
            'Относительный номер текущего элемента
            ' (относительно "базового" элемента)
            '  в массиве "Таблицы терминалов"
        intCount = 0
           'Порядковый номер элемента в массивах управляющих элементов форм
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
    Loop
    
            'Относительный номер текущего элемента
            ' (относительно "базового" элемента)
            '  в массиве "Таблицы терминалов"
    If intCount < 15 Then
        intCount = intCount + 1
    Else
        intCount = 0
           'Порядковый номер элемента в массивах управляющих элементов форм
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
            'Выйти из процедуры
        Exit Sub
    End If
            
            'Есть запросы на обслуживание терминалов - пропустить цикл опроса
            '   терминалов, предоставив его для работы Планировщика (Main)
    If prtPortC(intControlIndex).Tag > 0 Then
            'Восстановить предыдущий Относительный номер
            '  текущего элемента  (относительно "базового" элемента)
            '  - Подготовка корректного следующего входа в данную процедуру
        intCount = intCount - 1
        Exit Sub
    End If
            
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
    intRequest = (prtPortC(intControlIndex).CommPort - 2) * 15 + intCount
        
            ' "Controller" ЛОГИЧЕСКИ ВКЛЮЧЕН и НЕ ЗАНЯТ обслуживанием
            '   терминала либо ожидает Оранжевого индикатора на считывателе
    If (Mid(gAddrPort(0, intRequest), 4) = "0" Or _
    Mid(gAddrPort(0, intRequest), 4) = "#") And _
    Mid(gAddrPort(0, intRequest), 1, 2) <> "00" Then
            'Очистить приемный буфер порта
        prtPortC(intControlIndex).InBufferCount = 0
            'Послать "Controller'y" команду - ЧТЕНИЕ ПЕРСОНАЛЬНОГО КОДА
            '  с собственным адресом
        prtPortC(intControlIndex).Output = Chr(CLng(CByte(176) Or CByte(intCount)))
             'Ждать завершения передачи команды ЧТЕНИЕ ПЕРСОНАЛЬНОГО КОДА
        Do
        Loop Until prtPortC(intControlIndex).OutBufferCount = 0
            ' "Controller" ЛОГИЧЕСКИ ВЫКЛЮЧЕН или ЗАНЯТ  обслуживанием
            '   терминала и не ожидает Оранжевого индикатора на считывателе
    Else
            'Искать следующий "Controller", который ЛОГИЧЕСКИ ВКЛЮЧЕН
            '  и НЕ ЗАНЯТ обслуживанием терминала либо ожидает
            '  Оранжевого индикатора на считывателе
        Do Until (Mid(gAddrPort(0, intRequest), 4) = "0" Or _
        Mid(gAddrPort(0, intRequest), 4) = "#") And _
        Mid(gAddrPort(0, intRequest), 1, 2) <> "00"
        
            'Относительный номер текущего элемента
            ' (относительно "базового" элемента)
            '  в массиве "Таблицы терминалов"
            If intCount < 15 Then
                intCount = intCount + 1
            Else
                intCount = 0
           'Порядковый номер элемента в массивах управляющих элементов форм
                If intControlIndex < 3 Then
                    intControlIndex = intControlIndex + 1
                Else
                    intControlIndex = 0
                End If
            'Выйти из процедуры
                Exit Sub
            End If
            
            'Номер текущего элемента
            ' в массиве "Таблицы терминалов",
            intRequest = (prtPortC(intControlIndex).CommPort - 2) * 15 + intCount
        
        Loop
            'Восстановить предыдущий Относительный номер
            '  текущего элемента  (относительно "базового" элемента)
            '  - Подготовка корректного следующего входа в данную процедуру
        If intCount <> 0 Then intCount = intCount - 1
    
    End If
            
End Sub
            
            'Вызов процедуры обработки нажатия электронной "Кнопки"
Public Function OpenBarrier(intIndex As Integer)
            
            'Имитировать нажатие электронной "Кнопки"
    Call cmdOpen_Click(intIndex)

End Function

            'Процедура обработки нажатия электронной "Кнопки"
Private Sub cmdOpen_Click(intIndex As Integer)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Код возврата функций коррекции ячейки "Reserve" в "Таблице персон"
            '  после Регистрации и Исключения Клиента
Dim intCode As Integer
            'Сделать электронную "Кнопку" недоступной
    cmdOpen(intIndex).Enabled = False
            ' "Кнопка" хранит адрес "Controller'a", требующего
            '  ручного подтверждения открытия терминала или
            '  включено "Ручное" управление терминалами
    If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Or _
    cmdOpen(intIndex).Caption <> chkTerm(intIndex).Caption _
    And optManual = True Then
            'Анализ состояния соответствующего "Controller'a"
        vntAddr = cmdOpen(intIndex).Tag
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
            ' "Controller" ЗАНЯТ обслуживанием терминала - выход из процедуры
        If Mid(gAddrPort(0, intRequest), 4) <> "0" Then Exit Sub
            ' "Controller" ЛОГИЧЕСКИ выключен из системы - выход из процедуры
        If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then Exit Sub
            'Запросить "Последовательность открытия терминала"
            '  от электронной "Кнопки"
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "1"
            'Увеличить счетчик "Controller'ов", ЗАНЯТых обслуживанием терминалов
        prtPortC(intIndex).Tag = prtPortC(intIndex).Tag + 1
            'Очистить электронную "Кнопку" при включенном
            '  "Автоматическом" управлении терминалами
        If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Then
            cmdOpen(intIndex).Caption = chkTerm(intIndex).Caption
            cmdOpen(intIndex).Tag = 0
            'Выключить таймер электронной "Кнопки"
            tmrButton(intIndex).Enabled = False
            'Установлен признак процедуры Регистрации Клиента Автостоянки
            If imgParkingInData(intIndex).Tag = 1 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
            '  (Автомобиль въехал) после Регистрации Клиента Автостоянки
                intCode = frmTablePerson.InputParking(intIndex)
            'Была некорректная ситуация при коррекции ячейки
            '  "Reserve" в "Таблице персон" (отсутствует признак Регистрации)
                If intCode <> 0 Then
            'Окно собщения с запросом изменения "Таблицы персон" - на экран
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                    BeepSound
                    MsgBox "Error Parking Registration  !!!", intButtonsAndIcons, "Error"
                End If
            
            'Установлен признак процедуры Регистрации Клиента Предприятия
            ElseIf imgAccessInData(intIndex).Tag = 1 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
            '  (Посетитель вошел) после Регистрации Клиента Предприятия
                intCode = frmTablePerson.InputAccess(intIndex)
            'Была некорректная ситуация при коррекции ячейки
            '  "Reserve" в "Таблице персон" (отсутствует признак Регистрации)
                If intCode <> 0 Then
            'Окно собщения с запросом изменения "Таблицы персон" - на экран
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                    BeepSound
                    MsgBox "Error Access Registration  !!!", intButtonsAndIcons, "Error"
                End If
            
            'Установлен признак процедуры Исключения Клиента Автостоянки
            ElseIf imgParkingOutData(intIndex).Tag = 1 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
            '  (Автомобиль выехал) после Исключения Клиента Автостоянки
                intCode = frmTablePerson.OutputParking(intIndex, 6)
            
            'Установлен признак процедуры Исключения Клиента Предприятия
            ElseIf imgAccessOutData(intIndex).Tag = 1 Then
            'Коррекция ячейки "Reserve" в "Таблице персон"
            '  (Посетитель вышел) после Исключения Клиента
                intCode = frmTablePerson.OutputAccess(intIndex, 9)
            End If
            
            'Сбросить признаки процедур Регистрации
            '   и Удаления Клиентов
            imgParkingInData(intIndex).Tag = 0
            imgParkingOutData(intIndex).Tag = 0
            
            imgAccessInData(intIndex).Tag = 0
            imgAccessOutData(intIndex).Tag = 0
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора
            imgParkingInData(intIndex).Enabled = True
            imgParkingOutData(intIndex).Enabled = True
            imgParkingInfoData(intIndex).Enabled = True
            
            imgAccessInData(intIndex).Enabled = True
            imgAccessOutData(intIndex).Enabled = True
            imgAccessInfoData(intIndex).Enabled = True
        End If
        
            'Протоколирование события - нажатие электронной "Кнопки"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            'Системный пароль
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "BUTTON PRESSING"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
        
    End If

End Sub

            'Обработка события "TimeOut" электронной "Кнопки"
Private Sub tmrButton_Timer(intIndex As Integer)
            'Выключить таймер электронной "Кнопки"
    tmrButton(intIndex).Enabled = False
            'Установить признак "TimeOut" для электронной "Кнопки"
    tmrButton(intIndex).Tag = 1
            'Очистить электронную "Кнопку" при включенном
            '  "Автоматическом" управлении терминалами
    If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Then
        cmdOpen(intIndex).Caption = chkTerm(intIndex).Caption
        cmdOpen(intIndex).Tag = 0
            'Сделать электронную "Кнопку" недоступной
        cmdOpen(intIndex).Enabled = False
    End If
            'Сбросить признаки процедур Регистрации
            '   и Удаления Клиентов
    imgParkingInData(intIndex).Tag = 0
    imgParkingOutData(intIndex).Tag = 0
    
    imgAccessInData(intIndex).Tag = 0
    imgAccessOutData(intIndex).Tag = 0
            'Сделать доступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора
    imgParkingInData(intIndex).Enabled = True
    imgParkingOutData(intIndex).Enabled = True
    imgParkingInfoData(intIndex).Enabled = True
    
    imgAccessInData(intIndex).Enabled = True
    imgAccessOutData(intIndex).Enabled = True
    imgAccessInfoData(intIndex).Enabled = True

End Sub

            'Процедура контроля времени при вводе пароля - событие "TimeOut"
Private Sub tmrPasswTimeOut_Timer()
            'Издать звуковой сигнал
    BeepSound
    
                'Протоколирование события - "TimeOut" при вводе пароля
    gProtocol.strProtocName = "????????????????"
            'Системный пароль
    gProtocol.strProtocPersonCode = ""
            'Статус
    gProtocol.strProtocStatus = ""
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "PASSWORD TimeOut"
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol

            ' "Очистка" поля пароля пробелами
    txtPassword.Text = ""
            'Сделать недоступным поле пароля
    txtPassword.Enabled = False
            ' "Погасить" этикетку "Пароль"
    lblPassword.Enabled = False
            'Сбросить контроль времени ввода пароля
    tmrPasswTimeOut.Enabled = False
            'Сделать доступным переключатель "Выполнение/Установки"
    chkSetup.Enabled = True
            'Сделать доступной кнопку "Exit"
    cmdExit.Enabled = True
            'Установить фокус на опции "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    
End Sub

            'Процедура обработки "Щелчка мыши" на поле пароля
Private Sub txtPassword_Click()
            'Ранее запрошены "Exit" или "SetUp" - выход из процедуры
    If tmrPasswTimeOut.Enabled = True Then Exit Sub
            ' "Очистка" нового пароля пробелами
    strPassword = ""
            'Установить контроль времени ввода пароля
    tmrPasswTimeOut.Enabled = True
            'Установить фокус на поле пароля
    txtPassword.SetFocus
           'Удержание фокуса клавиатуры на поле пароля до его ввода
           '   или истечения времени ввода пароля
    Do While strPassword = "" And tmrPasswTimeOut.Enabled = True
        DoEvents
    Loop
    
            'Протоколирование события - "Ввод нового пароля"
    gProtocol.strProtocName = "????????????????"
            'Системный пароль
    gProtocol.strProtocPersonCode = strPassword
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "New PASSWORD"
            'Записать строку в файл "Таблицы протокола"
    WriteProtocol
            'Установить новый пароль в качестве текущего
    If strPassword <> "" Then txtPassword.Tag = strPassword
            ' "Очистка" поля пароля пробелами
    txtPassword.Text = ""
            'Сделать недоступным поле пароля
    txtPassword.Enabled = False
            ' "Погасить" этикетку "Пароль"
    lblPassword.Enabled = False
            'Сделать доступным переключатель "Выполнение/Установки"
    chkSetup.Enabled = True
    

End Sub

            'Процедура обработки получения фокуса полем пароля
Private Sub txtPassword_GotFocus()

            'Сделать недоступной кнопку "Exit"
    cmdExit.Enabled = False
            'Сделать недоступным переключатель "Выполнение/Установки"
    chkSetup.Enabled = False

            ' "Проявить" этикетку "Пароль"
    lblPassword.Enabled = True

End Sub

            'Процедура ввода и анализа пароля
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
            'Пароль ввведен
    If KeyAscii = vbKeyReturn Then
    
    
            'Пароль ?
        strPassword = txtPassword.Text
        
            'Протоколирование события - "Ввод пароля"
        gProtocol.strProtocName = "????????????????"
            'Системный пароль
        gProtocol.strProtocPersonCode = txtPassword.Text
            'Статус
        gProtocol.strProtocStatus = "04 - Manager"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "PASSWORD Input"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
        
            'Анализ правильности пароля
        If txtPassword.Text = txtPassword.Tag Then
            ' "Очистка" поля пароля пробелами
            txtPassword.Text = ""
            'Сделать недоступным поле пароля
            txtPassword.Enabled = False
            ' "Погасить" этикетку "Пароль"
            lblPassword.Enabled = False
        End If
    End If

End Sub

            'Обработка "щелчка" мыши на элементе "imgViewClose"
Private Sub imgViewClose_click(intIndex As Integer)

'''          'Форма "frmDemo" доступна и режим "Выполнение"
'''    If frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
           
'''            'Установлена Опция разделения времени
'''            '   (параллельное выполнение процедур)
'''        If gTimeShare = 1 Then
'''            'Видимый элемент "imgAccessOutData"
'''            If imgAccessOutData(intIndex).Visible = True Then
'''                Call mnuAccessServData_Click
'''            'Видимый элемент "imgParkingOutData"
'''            ElseIf imgParkingOutData(intIndex).Visible = True Then
'''                Call mnuParkingServData_Click
'''            End If
'''            Exit Sub
'''        End If
    
'''    End If

End Sub

            'Инициализация формирования ПЕРСОНАЛЬНОГО КОДА
            '  и других данных при АвтоРегистрации (через спец "Controller"
            '  с кнопкой "DALLAS") Временных Клиентов Автостоянке или
            '  Предприятия
Public Function AutoRegDallasButton(ByVal vntPersonCode As Variant, _
intIndex As Integer, ByVal strAddrPortType As String)
            
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при формировании  "PersonCode"
Dim intDallasCode  As Integer

            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataParkingIn" - выйти из функции
    If frmDataParkingIn.Enabled = False And _
    Mid(Trim(strAddrPortType), 4) = "ParkI" Then Exit Function
            'Попытка повторного вызова функции до нормального
            '  выхода из модуля формы "frmDataAccessIn" - выйти из функции
    If frmDataAccessIn.Enabled = False And _
    Mid(Trim(strAddrPortType), 4) = "AcceI" Then Exit Function
    
            'Сделать недоступной форму "frmDataParkingIn"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    frmDataParkingIn.Enabled = False
            'Сделать недоступной форму "frmDataAccessIn"
    If Mid(Trim(strAddrPortType), 4) = "AcceI" Then _
    frmDataAccessIn.Enabled = False
            
            'Если это Временный Клиент Автостоянки, то выполняеть
            '  процедуру формирования "PersonCode", "Info" и Печать
            '  талона со штрих-кодом (+ чека) при АвтоРегистрации Клиента
            '  Автостоянки через специальный "Controller" с кнопкой "DALLAS"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    intDallasCode = frmDataParkingIn.DallasButton(strAddrPortType, intIndex)
            'Если это Временный Клиент Предприятия, то выполнить
            '  процедуру формирования "PersonCode", "Info" и Печать
            '  талона со штрих-кодом (+ чека) при АвтоРегистрации Клиента
            '  Предприятия через специальный "Controller" с кнопкой "DALLAS"
    If Mid(Trim(strAddrPortType), 4, 2) = "AcceI" Then _
    intDallasCode = frmDataAccessIn.DallasButton(strAddrPortType, intIndex)

            'Сделать доступной форму "frmDataParkingIn"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    frmDataParkingIn.Enabled = True
            'Сделать доступной форму "frmDataAccessIn"
    If Mid(Trim(strAddrPortType), 4) = "AcceI" Then _
    frmDataAccessIn.Enabled = True

End Function

            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
Public Sub PrintDocument(ByVal strProtocName As String, _
            ByVal strProtocPersonCode As String, _
            ByVal strProtocStatus As String, _
            ByVal strProtocTime As String, _
            ByVal strProtocDate As String, _
            ByVal strProtocReserve As String, _
            ByRef intError As Integer)
        
            'Номер текстового файла
Dim intFileNum As Integer
            'Полное имя текстового файла (с указанием "пути")
Dim strPathFileName As String
            'Массив строк "исправленного" текстового файла Кассового аппарата
Dim strCashPrinter(2) As String
            'Номер текущей строки текстового файла Кассового аппарата
Dim intRowNum As Integer
            'Номер позиции заданного символа в строке
Dim intPosNum As Integer
            'Код возврата при вызове функции "Shell"
Dim vntShell As Variant
            'Переменная-строка "Печать документа"
Dim strDocument As String
            'Время регистрации Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Клиента
Dim strHour As String
Dim strMinute As String
            'Рабочий счетчик
Dim intCount As Integer
            'Счетчик времени ожидания кода состояния
            '  принтера штрих-кода
Dim lngTimeCount As Long
            'Рабочая переменная
Dim intWork As Integer
            'Рабочая строка
Dim strWork As String

            'Буфер приема данных состояния от Принтера штрих-кода
Dim Buffer() As Byte
           'Пременнные для преобразования битовых данных состояния
            ' от Принтера-штрих-кода в шестнадцатиричное представление
Dim strBuffer As String
Dim intBuffer1 As Integer
Dim intBuffer2 As Integer

            'Сброс признака ошибки
    intError = 0

            'Типы устройств для "Печати Документа" (1 - ПРОСТОЙ ЧЕКОВЫЙ
            '  ПРИНТЕР, 2 - Принтер ШтрихКода, 4 - Кассовый Чековый принтер;
            '  Возможны комбинации: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7 Then
        strDocument = ""
            'Инициализация Чекового Принтера
        strDocument = strDocument + Chr(CLng(CByte(27))) + Chr(CLng(CByte(64)))
            'Установка шрифтов на Чековом Принтере
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(82))) + Chr(CLng(CByte(0)))
            'Подача бумаги на Чековом Принтере на 5-ь линий
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(5)))
            'Установка размера шрифта на Чековом Принтере 7х7
        strDocument = strDocument + Chr(CLng(CByte(27))) + Chr(CLng(CByte(77)))
            'Спецсимволы управления "sp" Чековым Принтером - Отступы
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            'Разделительные нули "00"H - Отступы
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            'Формирование Названия Компании
        intWork = Len(Trim(gPrintSIAName))
        For intCount = 1 To intWork Step 1
            'Текущий символ из Названия Компании и Возврат каретки
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(gPrintSIAName), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            'Подача бумаги на Чековом Принтере на 1-у линию
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            'Спецсимволы управления "sp" Чековым Принтером - Отступы
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            'Разделительные нули "00"H - Отступы
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            'Формирование Подчеркивания Названия Компании
        For intCount = 1 To 16 Step 1
            'Текущий символ из Подчеркивания и Возврат каретки
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid("================", intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            'Подача бумаги на Чековом Принтере на 1-у линию
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            'Спецсимволы управления "sp" Чековым Принтером - Отступы
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            'Разделительные нули "00"H - Отступы
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            'Формирование Персонального кода
        strWork = "#### = " + Trim(strProtocPersonCode)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            'Текущий символ из Персонального кода и Возврат каретки
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strWork), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            'Подача бумаги на Чековом Принтере на 1-у линию
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            'Спецсимволы управления "sp" Чековым Принтером - Отступы
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            'Разделительные нули "00"H - Отступы
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            'Формирование Времени/Даты
        strWork = Trim(strProtocTime) + "||" + Trim(strProtocDate)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            'Текущий символ из Времени/Даты и Возврат каретки
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strWork), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            'Подача бумаги на Чековом Принтере на 1-у линию
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            'Спецсимволы управления "sp" Чековым Принтером - Отступы
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            'Разделительные нули "00"H - Отступы
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            'Формирование Примечания/Суммы
        intWork = Len(Trim(strProtocReserve))
        For intCount = 1 To intWork Step 1
            'Текущий символ из Примечания/Суммы и Возврат каретки
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strProtocReserve), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            'Возврат каретки
        strDocument = strDocument + Chr(CLng(CByte(13)))
            'Подача бумаги на Чековом Принтере на 10-ь линий
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(10)))
            'Послать данные на Чековый Принтер
        prtPortDocument.Output = strDocument
             'Ждать завершения передачи данных на Чековый Принтер
        Do
        Loop Until prtPortDocument.OutBufferCount = 0
    End If
    
            'Типы устройств для "Печати Документа" (1 - Простой Чековый принтер,
            '  2 - ПРИНТЕР ШТРИХ КОДА, 4 - Кассовый Чековый принтер;
            '  Возможны комбинации: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            'Не признак Регистрация Клиента или Клиента Автостоянки -
            '  Печать штрих-кода не производится
        If frmDataAccessIn.Tag <> 1 And frmDataParkingIn.Tag <> 1 _
        Then GoTo BarCodeOK
            'Обработка ошибок при подготовке Печати штрих-кода на Пропуске
        On Error GoTo BarCodeError
            
            
            'Инициализация Принтера штрих-кода
        strDocument = ""
        strDocument = Chr(CLng(CByte(27))) + Chr(CLng(CByte(64)))
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            'Очистить приемный буфер порта
        prtPortBarCode.InBufferCount = 0
            'Получение информации о статусе Принтера штрих-кода
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(97))) + _
        Chr(CLng(CByte(15)))
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтера штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
             
            'Ждать от Принтера штрих-кода кода состояния
        For lngTimeCount = 1 To 99000 Step 1
            If prtPortBarCode.InBufferCount = 4 Then
                Buffer = ""
            'Полученные данные в приемный буфер для дальнейшей обработки
                Buffer = prtPortBarCode.Input
            'Очистить приемный буфер порта
                prtPortBarCode.InBufferCount = 0
                Exit For
            End If
            'Обработать возможные события
            DoEvents
        Next
            'Неготовность принтера штрих-кода
        If lngTimeCount > 99000 Then GoTo BarCodeError
            
            'Преобразовать данные из буфера терминала в шестнадцатиричный вид
        strBuffer = ""
        intCount = 0
        Do While intCount <= 3
            intBuffer1 = (CByte(Buffer(intCount)) And CByte(240)) / 16
            intBuffer2 = CByte(Buffer(intCount)) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCount = intCount + 1
        Loop
            'Некорректный код состояния Принтера штрих-кода
        If Trim(strBuffer) <> "00400014" Then GoTo BarCodeError
            
            'Корректный код состояния Принтера штрих-кода
            
            'Формирование Даты/Времени
        strDocument = ""
        strWork = Trim(Format(Now, "h:mm:ss")) + " || " + _
        Trim(Format(Now, "dd/mm/yyyy"))
        strWork = Trim(Format(Now, "h:mm:ss"))
            'Часы
        intHour = Hour(strWork)
        If intHour < 10 Then
            strHour = "0" + Trim(Str(intHour))
        Else
            strHour = Trim(Str(intHour))
        End If
            'Минуты
        intMinute = Minute(strWork)
        If intMinute < 10 Then
            strMinute = "0" + Trim(Str(intMinute))
        Else
            strMinute = Trim(Str(intMinute))
        End If
        strWork = "IZDOTS " + strHour + ":" + strMinute + _
        " || " + Trim(Format(Now, "dd/mm/yyyy"))
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            'Текущий символ из Даты/Времени
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            'Пропуск строки
        strDocument = strDocument + Chr(CLng(CByte(10)))
             'Формирование Названия Компании или другой информации
        intWork = Len("BEZMAKSAS LAIKS 2 st.")
        For intCount = 1 To intWork Step 1
            'Текущий символ из Названия Компании
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid("BEZMAKSAS LAIKS 2 st.", intCount, 1))))
        Next
            'Пропуск строки
        strDocument = strDocument + Chr(CLng(CByte(10)))
            'Формирование Подчеркивания Названия Компании
        For intCount = 1 To intWork Step 1
            'Текущий символ из Подчеркивания
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid("================================", intCount, 1))))
        Next
            'Пропуск строки
        strDocument = strDocument + Chr(CLng(CByte(10)))
            'Формирование Персонального кода
        strWork = Right(Trim(strProtocPersonCode), 10)
        strWork = "#### = " + Trim(strWork)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            'Текущий символ из Персонального кода
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            'Подача бумаги на Чековом Принтере на 1-у линию
        strDocument = strDocument + Chr(CLng(CByte(10)))
            'Послать данные на Чековый Принтер
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
    
        strDocument = ""
            'Установить ширину штрих-кода
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(119))) + _
        Chr(CLng(CByte(4)))
            'Установить фонт типа "А" для печати символов над штрих-кодом
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(102))) + _
        Chr(CLng(CByte(0)))
            'Сбросить режим печати символов над штрих-кодом
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(72))) + _
        Chr(CLng(CByte(0)))
            'Послать данные на Чековый Принтер
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
    
        strDocument = ""
        strWork = "99" + Right(Trim(strProtocPersonCode), 10)
            'Установка кода "EAN13"
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(107))) + _
        Chr(CLng(CByte(2)))
            'Формирование основной части штрих-кода EAN13
        For intCount = 1 To 12 Step 1
            'Текущий символ из Персонального кода
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            'Формирование контрольной суммы штрих-кода EAN13
        intWork = 0
        For intCount = 1 To 12 Step 1
            If (intCount Mod 2) = 0 Then
                intWork = intWork + CInt(Mid(strWork, intCount, 1)) * 3
            Else
                intWork = intWork + CInt(Mid(strWork, intCount, 1))
            End If
        Next
        If (intWork Mod 10) = 0 Then
            intWork = 0
        Else
            intWork = Int(intWork / 10) * 10 + 10 - intWork
        End If
            'Контрольная сумма
        strDocument = strDocument + Trim(Str(intWork))
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0

            'Подача бумаги на Принтере штрих-кода на N линий
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(100))) + Chr(CLng(CByte(9)))
            
            
'            'Частичный надрез талона
'        strDocument = strDocument + Chr(CLng(CByte(27))) + _
'        Chr(CLng(CByte(109)))
            'Полное отрезание талона
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(105)))
            
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            'Подача бумаги на Принтере штрих-кода на N линий
            ' (Выталкивание талона)
        strDocument = ""
        For intCount = 1 To gTalonLength Step 1
            'Подача бумаги на Чековом Принтере на 1-у линию
            strDocument = strDocument + Chr(CLng(CByte(10)))
        Next
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтера штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            'Очистить приемный буфер порта
        prtPortBarCode.InBufferCount = 0
            'Получение информации о статусе Принтера штрих-кода
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(97))) + _
        Chr(CLng(CByte(15)))
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтера штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
             
            'Ждать от Принтера штрих-кода кода состояния
        For lngTimeCount = 1 To 99000 Step 1
            If prtPortBarCode.InBufferCount = 4 Then
                Buffer = ""
            'Полученные данные в приемный буфер для дальнейшей обработки
                Buffer = prtPortBarCode.Input
            'Очистить приемный буфер порта
                prtPortBarCode.InBufferCount = 0
                Exit For
            End If
            'Обработать возможные события
            DoEvents
        Next
            'Неготовность принтера штрих-кода
        If lngTimeCount > 99000 Then GoTo BarCodeError
            
            'Преобразовать данные из буфера терминала в шестнадцатиричный вид
        strBuffer = ""
        intCount = 0
        Do While intCount <= 3
            intBuffer1 = (CByte(Buffer(intCount)) And CByte(240)) / 16
            intBuffer2 = CByte(Buffer(intCount)) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCount = intCount + 1
        Loop
            'Некорректный код состояния Принтера штрих-кода
        If Trim(strBuffer) <> "00400014" Then GoTo BarCodeError
            
            'Корректный код состояния Принтера штрих-кода -
            '  протоколирование события "Печать Штрих-кода на Пропуске"
        gProtocol.strProtocName = "Print BarCode"
            'Персональный код
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Примечания
        gProtocol.strProtocReserve = "BAR_CODE BOX"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            'Сброс признака ошибки
        intError = 0
        
        GoTo BarCodeOK
BarCodeError:
            'Издать звуковой сигнал
        BeepSound
            'Сделать видимой метку сообщения
        lblErrorBarCodePrinter.Visible = True
        
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        strMessage = "BarCode Printer Error !!!"
            'Отослать СООБЩЕНИЕ
        Call SendMessage(strMessage)
        
            'Протоколирование события - "Ошибка Принтера штрих-кода"
        gProtocol.strProtocName = "Print BarCode"
            'Персональный код
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "BAR_CODE ERROR"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
            ' Если нет Препроцессоров в локальной сети
        If gNetPreprocNum = 0 Then
            'Удаление персонального кода зарегистрированного Временного
            '  Клиента из "Таблицы персон"
            Call frmTablePerson.AutoDelParking(strProtocPersonCode, _
            strProtocStatus)
        End If
            
            'Установка признака ошибки
        intError = 1

BarCodeOK:
        On Error GoTo 0
    End If
            
            'Типы устройств для "Печати Документа" (1 - Простой Чековый принтер,
            '  2 - Принтер ШтрихКода, 4 - КАССОВЫЙ ЧЕКОВЫЙ ПРИНТЕР;
            '  Возможны комбинации: 1+2, 1+4, 2+4, 1+2+4)
            '  и Признак ошибки сброшен
    If (gDocument = 4 Or gDocument = 5 Or gDocument = 6 Or gDocument = 7) And _
    intError = 0 Then
            'Нулевая сумма оплаты - Печать чека не производится
        If Right(Trim(strProtocReserve), 9) = "000,00 Ls" Or _
        Right(Trim(strProtocReserve), 3) <> " Ls" Then GoTo CashOK
            'Обработка ошибок при подготовке Печати чека на Кассовом аппарате
        On Error GoTo CashError
            
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить полный "путь" к текстовому файлу Кассового аппарата
        strPathFileName = "C:\BarCashPrinter\Rs2810s.txt"
    
            'Открыть текстовый файл Кассового аппарата для Ввода
        Open strPathFileName For Input As intFileNum
            'Цикл по всем строкам текстового файла Кассового аппарата
        For intRowNum = 1 To 3
            'Читать строку текстового файла Кассового аппарата в буфер
            strDocument = Input(34, intFileNum)
            'Корректировать 1-ую строку текстового файла Кассового аппарата
            If intRowNum = 1 Then
            'Позиция подстроки "text=" в буфере
                intPosNum = InStr(1, strDocument, "text=")
            'Неверное содержимое буфера
                If intPosNum = 0 Then
                    Close intFileNum
                    GoTo CashError
            'Записать в Чек "Персональный код" Клиента
                Else
                    strDocument = Left(strDocument, intPosNum + 5) + _
                    Trim(strProtocPersonCode) + """" + Right(strDocument, 2)
                    strCashPrinter(intRowNum - 1) = strDocument
                End If
            'Корректировать 2-ую строку текстового файла Кассового аппарата
            ElseIf intRowNum = 2 Then
            'Позиция подстроки "deptnr=" в буфере
                intPosNum = InStr(1, strDocument, "deptnr=")
            'Неверное содержимое буфера
                If intPosNum = 0 Then
                    Close intFileNum
                    GoTo CashError
            'Запись в Чек "Кода товара"
                Else
            'Признак (Авто)Регистрация Клиента
                    If frmDataAccessIn.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "1" + _
                        Mid(strDocument, intPosNum + 8)
            'Признак (Авто)Удаления Клиента
                    ElseIf frmDataAccessOut.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "2" + _
                        Mid(strDocument, intPosNum + 8)
            'Признак (Авто)Коррекции Клиента
                    ElseIf frmDataAccessServ.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "3" + _
                        Mid(strDocument, intPosNum + 8)
            'Признак (Авто)Регистрации Автомобиля
                    ElseIf frmDataParkingIn.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "4" + _
                        Mid(strDocument, intPosNum + 8)
            'Признак (Авто)Удаления Автомобиля
                    ElseIf frmDataParkingOut.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "5" + _
                        Mid(strDocument, intPosNum + 8)
            'Признак (Авто)Коррекции Автомобиля
                    ElseIf frmDataParkingServ.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "6" + _
                        Mid(strDocument, intPosNum + 8)
                    End If
            'Позиция подстроки "amount=" в буфере
                    intPosNum = InStr(intPosNum + 8, strDocument, "amount=")
            'Неверное содержимое буфера
                    If intPosNum = 0 Then
                        Close intFileNum
                        GoTo CashError
            'Запись в Чек "Стоимости товара"
                    Else
                        strDocument = Left(strDocument, intPosNum + 6) + _
                        Right(strDocument, 2)
            'Позиция символа "," в "Примечании" протокола
                        intPosNum = InStr(1, strProtocReserve, ",")
            'Неверное содержимое буфера
                        If intPosNum = 0 Then
                            Close intFileNum
                            GoTo CashError
            'Стандартизация формата записи "Стоимость товара"
                        Else
                            strDocument = Left(strDocument, Len(strDocument) - 2) + _
                            Mid(strProtocReserve, intPosNum - 3, 3) + "." + _
                            Mid(strProtocReserve, intPosNum + 1, 2) + _
                            Right(strDocument, 2)
                            If Len(strDocument) < 34 Then _
                            strDocument = Left(strDocument, Len(strDocument) - 2) + _
                            Left("          ", 34 - Len(strDocument)) + Right(strDocument, 2)
                            strCashPrinter(intRowNum - 1) = strDocument
                        End If
                    End If
                    
                End If
            'Не корректировать 3-ью строку текстового файла Кассового аппарата
            ElseIf intRowNum = 3 Then
                strCashPrinter(intRowNum - 1) = strDocument
            End If
            
        Next
            'Закрыть текстовый файл Кассового аппарата
        Close intFileNum
            'Открыть текстовый файл Кассового аппарата для Вывода
            '  (в режиме двоичного доступа - для избавления от кавычек)
        Open strPathFileName For Binary As intFileNum
            'Цикл по всем строкам текстового файла Кассового аппарата
        For intRowNum = 1 To 3
            'Записать массив байтов в текстовый файл Кассового аппарата
            Put #intFileNum, , strCashPrinter(intRowNum - 1)
        Next
            'Закрыть текстовый файл Кассового аппарата
        Close intFileNum
    
        
        vntShell = Shell("C:\BarCashPrinter\Rs2810s.bat", 0)
        If vntShell <> 0 Then
            'Протоколирование события - "Печать Чека на Кассовом аппарате"
            gProtocol.strProtocName = "Print Check"
            'Персональный код
            gProtocol.strProtocPersonCode = ""
            'Статус
            gProtocol.strProtocStatus = ""
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
            gProtocol.strProtocReserve = "CASH BOX"
            'Записать строку в файл "Таблицы протокола"
            WriteProtocol
            GoTo CashOK
        End If
CashError:
            'Издать звуковой сигнал
        BeepSound
            'Вывод сообщения
        If optEnglish = True Then
            MsgBox ("The CashPrinter Error")
        Else
            MsgBox ("Nepareizs 'CashPrinter' ")
        End If
        
            'Протоколирование события - "Печать Чека на Кассовом аппарате"
        gProtocol.strProtocName = "Print Check"
            'Персональный код
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "CASH BOX ERROR"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
CashOK:
        On Error GoTo 0
    
    End If

End Sub

            'Печать Документа (Z_Отчета на принтере Штрих-кода)
Public Sub PrintZReport(ByVal strProtocName As String, _
            ByVal strProtocPersonCode As String, _
            ByVal strProtocStatus As String, _
            ByVal strProtocTime As String, _
            ByVal strProtocDate As String, _
            ByVal strProtocReserve As String, _
            ByVal strMoney_Report As String, _
            ByRef strZ_Report As String)
            'Номер текстового файла
Dim intFileNum As Integer
            'Полное имя текстового файла (с указанием "пути")
Dim strPathFileName As String
            'Номер позиции заданного символа в строке
Dim intPosNum As Integer
            'Код возврата при печати этикетки на Принтере штрих-кода
Dim vntBuffer As Variant
            'Переменная-строка "Печать документа"
Dim strDocument As String
            'Время печати Z_Отчета
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время печати Z_Отчета
Dim strHour As String
Dim strMinute As String

            'Типы устройств для "Печати Документа" (1 - Простой Чековый принтер,
            '  2 - ПРИНТЕР ШТРИХ КОДА, 4 - Кассовый Чековый принтер;
            '  Возможны комбинации: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            'Обработка ошибок при подготовке Печати
        On Error GoTo BarCodeError
            
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить полный "путь" к текстовому файлу Принтера штрих-кода
        strPathFileName = "C:\BarCashPrinter\ZReport.txt"
    
            'Открыть текстовый файл Принтера штрих-кода для Ввода
        Open strPathFileName For Input As intFileNum
            'Читать файл Принтера штрих-кода в буфер
        strDocument = Input(FileLen(strPathFileName), intFileNum)
            'Закрыть текстовый файл Принтера штрих-кода
        Close intFileNum
    
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
       
            'Часы
        intHour = Hour(gProtocol.strProtocTime)
        If intHour < 10 Then
            strHour = "0" + Trim(Str(intHour))
        Else
            strHour = Trim(Str(intHour))
        End If
            'Минуты
        intMinute = Minute(gProtocol.strProtocTime)
        If intMinute < 10 Then
            strMinute = "0" + Trim(Str(intMinute))
        Else
            strMinute = Trim(Str(intMinute))
        End If

            'Корректировать строку "Время-Дата" текстового файла в буфере
        intPosNum = InStr(1, strDocument, """""")
            'Неверное содержимое буфера
        If intPosNum = 0 Then
            GoTo BarCodeError
            'Записать в Z_Отчет "Время-Дату" его выдачи
        Else
            strDocument = Left(strDocument, intPosNum) + _
            strHour + ":" + strMinute + " / " + _
            CStr(Trim(gProtocol.strProtocDate)) + Mid(strDocument, intPosNum + 1)
        End If
            'Корректировать строку "Z_Отчет" текстового файла в буфере
        intPosNum = InStr(intPosNum + 7, strDocument, "Z_Report = ")
            'Неверное содержимое буфера
        If intPosNum = 0 Then
            GoTo BarCodeError
            'Записать Сумму в Z_Отчет
        Else
            strDocument = Left(strDocument, intPosNum + 10) + _
            strMoney_Report + Mid(strDocument, intPosNum + 11)
        End If
            
            
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
        tmrTimeOut(1).Tag = 0
        tmrTimeOut(1).Enabled = True
            ' Цикл опроса состояния Принтера штрих-кода
        Do While DoEvents()
            'Если данные поступили
            If prtPortBarCode.InBufferCount > 1 Then
            'Выключить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
                tmrTimeOut(1).Enabled = False
           'Полученные данные в приемный буфер для дальнейшей обработки
                vntBuffer = prtPortBarCode.Input
            'Корректный ответ от Принтера штрих-кода
                If Len(vntBuffer) <= 3 Then
            'На протоколирование события
                    GoTo Protocol
            'Ошибка
                Else
            'На обработку ошибки
                    GoTo BarCodeError
                End If
            'Произошло событие TimeOut при ожидании
            '   КОДА СОСТОЯНИЯ от Принтера штрих-кода
            ElseIf tmrTimeOut(1).Tag <> 0 Then
            'Выключить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
                tmrTimeOut(1).Enabled = False
            'На обработку ошибки
                GoTo BarCodeError
            End If
        Loop
            
Protocol:
            'Протоколирование события - "Печать Z_Отчета"
        gProtocol.strProtocName = "Print Z_Report"
            'Персональный код
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "BAR_CODE BOX"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
            
        GoTo BarCodeOK
BarCodeError:
            'Издать звуковой сигнал
        BeepSound
            'Вывод сообщения
        If optEnglish = True Then
            MsgBox ("The BarCodePrinter Error")
        Else
            MsgBox ("Nepareizs 'BarCodePrinter' ")
        End If
            'Очистить указатель текущей точки "Z_Отчета"
        strZ_Report = ""
        On Error GoTo 0
        
            'Послать данные на Принтер штрих-кода (СБРОС ПРИНТЕРА)
        prtPortBarCode.Output = Chr(94) + Chr(64) + Chr(13)
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
        frmDataAccessIn.tmrParoleTimeOut.Enabled = True
            ' Цикл опроса состояния Принтера штрих-кода
        Do While DoEvents()
            'Если данные поступили
            If prtPortBarCode.InBufferCount > 1 Then
            'Выключить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
                frmDataAccessIn.tmrParoleTimeOut.Enabled = False
            'На печать ПУСТОЙ ЭТИКЕТКИ
                Exit Do
            'Произошло событие TimeOut при ожидании
            '   КОДА СОСТОЯНИЯ от Принтера штрих-кода
            ElseIf frmDataAccessIn.tmrParoleTimeOut.Enabled = False Then
            'На печать ПУСТОЙ ЭТИКЕТКИ
                Exit Do
            End If
        Loop
            'Сделать недоступными кнопки "OK" и "Cancel" формы "frmDataAccessIn"
        frmDataAccessIn.cmdOK.Enabled = False
        frmDataAccessIn.cmdCancel.Enabled = False
            'Закрыть последовательный порт для ПРИНТЕРА ШТРИХ-КОДА - очистка порта
        prtPortBarCode.PortOpen = False
            'Oткрыть последовательный порт для ПРИНТЕРА ШТРИХ-КОДА
        prtPortBarCode.PortOpen = True
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить полный "путь" к текстовому файлу Принтера штрих-кода
        strPathFileName = "C:\BarCashPrinter\ZReport.txt"
            'Открыть текстовый файл Принтера штрих-кода для Ввода
        Open strPathFileName For Input As intFileNum
            'Читать файл Принтера штрих-кода в буфер
        strDocument = Input(FileLen(strPathFileName), intFileNum)
            'Закрыть текстовый файл Принтера штрих-кода
        Close intFileNum
            'Послать данные на Принтер штрих-кода
        prtPortBarCode.Output = strDocument
             'Ждать завершения передачи данных на Принтер штрих-кода
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            'Включить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
        frmDataAccessIn.tmrParoleTimeOut.Enabled = True
            ' Цикл опроса состояния Принтера штрих-кода
        Do While DoEvents()
            'Если данные поступили
            If prtPortBarCode.InBufferCount > 1 Then
            'Выключить контроль события TimeOut при ожидании
            '  КОДА СОСТОЯНИЯ от Принтера штрих-кода
                frmDataAccessIn.tmrParoleTimeOut.Enabled = False
            'Выход из печатати ПУСТОЙ ЭТИКЕТКИ
                Exit Do
            'Произошло событие TimeOut при ожидании
            '   КОДА СОСТОЯНИЯ от Принтера штрих-кода
            ElseIf frmDataAccessIn.tmrParoleTimeOut.Enabled = False Then
            'Выход из печатати ПУСТОЙ ЭТИКЕТКИ
                Exit Do
            End If
        Loop
            
            'Протоколирование события - "Печать Z_Отчета"
        gProtocol.strProtocName = "Print Z_Report"
            'Персональный код
        gProtocol.strProtocPersonCode = ""
            'Статус
        gProtocol.strProtocStatus = ""
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "BAR_CODE ERROR"
            'Записать строку в файл "Таблицы протокола"
        WriteProtocol
        
BarCodeOK:
        On Error GoTo 0
    End If

End Sub
            
            'Процедура записи строки в файл "Таблицы протокола"
Public Sub WriteProtocol()
            
            'Записать строку в файл "Таблицы протокола"
    Put gProtocFileNum, gProtocRowNum, gProtocol
            'Номер следующей свободной строки (записи) в файле "Таблицы протокола"
    gProtocRowNum = gProtocRowNum + 1
            'Файл "Таблицы протокола" требует архивирования
    If gProtocRowNum > 32760 Then
            'Архивирование файла "Таблицы протокола"
        WriteProtocolToArchives
    End If
    
End Sub

            'Процедура автоматической записи "Protocol to Archives..."
Private Sub WriteProtocolToArchives()
            'Полное имя файла архива (с указанием "пути" к нему)
Dim strPathFileName As String
            'Номер архивного файла
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола"
Dim lngRecordLen As Long
            'Текущий номер строки "Таблицы протокола"
Dim intRowNum As Integer
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer

            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            ' Если это "Host Computer"
    If gPreprocName = "" Then
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = strPathFileName + gHost + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
            ' Если это Препроцессор
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = strPathFileName + gPreprocName + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
    End If
            
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
            'Начальная позиция в полном имени файла (за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить "старый" архивный файл, если он существует
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            'Обработка ошибок
    On Error GoTo UnDefError
            'Открыть выбранный архивный файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
    
            'Цикл по всем строкам "Таблицы протокола"
    For intRowNum = 1 To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
        Get gProtocFileNum, intRowNum, gProtocol
            'Вывести строку "Таблицы протокола" в архивный файл
        Put intFileNum, intRowNum, gProtocol
    Next
            'Закрыть выбранный  архивный файл
    Close intFileNum
            
            ' Если это Препроцессор
    If gPreprocName <> "" Then
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        strMessage = "Archive" + " " + Mid(strPathFileName, intSymbPos)
            'Отослать СООБЩЕНИЕ
        Call SendMessage(strMessage)
    End If
        
             'Закрыть "текущий" файл "Таблицы протокола"
    Close gProtocFileNum
           'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
        
            'Начальная позиция в полном имени файла (за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить  "текущий" файл "Таблицы протокола"
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
            'Получить свободный номер файла для новой "Таблицы протокола"
    gProtocFileNum = FreeFile
            'Открыть новый файл "Таблицы протокола" для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Номер следующей "свободной" строки в новом файле "Таблицы протокола"
    gProtocRowNum = 1
            'Установить признак сохранения протокола событий в умалчиваемом файле
    mnuSaveProtocol.Checked = True
    mnuSaveProtocolAs.Checked = False
    
    Exit Sub
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

End Sub
