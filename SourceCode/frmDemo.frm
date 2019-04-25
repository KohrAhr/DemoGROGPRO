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
            '����������-������ ������� ���
            ' ����������� ���������
Dim WithEvents qEvent As MSMQEvent
Attribute qEvent.VB_VarHelpID = -1
            '������ "������� ������"
Dim gPerson As PersonInfo
            '������ "��������� �������"
Dim gSystem As SystemInfo
            '������ "������� ����������"
Dim gTerminal As TerminalInfo
            '�������� � ����� �������� (��� ����������� ����������)
Dim lngResource As Long
           ' �������� ������ "�����������" ���������
Dim aComment(3, 23) As String
            ' �������� ������ ��������
Dim aCaption(3, 23) As String
             '����� ������
Dim strPassword As String
            '����� ������ ����� �������
Dim intLang As Integer
            '������ ����������� ���������
Dim strMessage As String


            '�������� ������� ���������� ������ "Alt"+ {"i", "<-" , "->", "+" � "-"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            '������ ������ �� �����
Dim intIndex As Integer
            '����� "����������"
    If chkSetup.Value = 1 And frmDemo.Enabled = True Then
            '������������ "������" ���� �� �������� "imgXXXXXXInData"
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
            '������������ "������" ���� �� �������� "imgXXXXXXOutData"
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
            '������������ "������" ���� �� �������� "imgXXXXXXInfoData"
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
            '������������ "������" ���� �� �������� "imgEmployeInData"
        ElseIf KeyCode = 107 And Shift = 4 Then
            If imgEmployeInData.Visible = True Then
                Call imgEmployeInData_Click
                Exit Sub
            End If
            '������������ "������" ���� �� �������� "imgEmployeOutData"
        ElseIf KeyCode = 109 And Shift = 4 Then
            If imgEmployeOutData.Visible = True Then
                Call imgEmployeOutData_Click
                Exit Sub
            End If
        End If
    End If
    
End Sub
            
            '�������� ���������� ������� ������� "������"
Private Sub chkDummy_Click()
    chkDummy.Value = 0
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus

End Sub

            '��������� ������� ���� �����
Private Sub Form_Resize()
            '������������ ������� ���� �����
    frmDemo.WindowState = 0
    
End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '��������� ������ ������� "Form BookKeeper Base" ���� "File"
Private Sub mnuFormBookKeeperBase_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '���������� ����� � "���� ���������"
Dim lngProtocolBaseCount As Long
            '����� ����� ������
Dim intFileNum As Integer
            '����� ������ "������� ���������" � DUMMY �����
Dim lngRecordLen As Long
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
Dim strDummyFileName As String
            '������� ����� ������ ������� DUMMY �����
Dim lngRowDummy As Long
            '������ ��� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ��� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � DUMMY ����
Dim intDayArchive As Integer
            '���������� ����� � ���������� ����� (������ ��� "TableProtocol")
Dim intRowQuan As Integer
            '������� ����� ������ ����������� ������
            '   ��� ������� "TableProtocol"
Dim intRowNumArchive As Integer
            '���������� ����� � "���� �����������"
Dim lngBookKeepingBaseCount As Long
            '������� ����� ����������������� ������ "���� �����������"
Dim lngBookKeepingRowNum As Long
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmDemo.MousePointer = vbHourglass
            '������� ������������ �������� ���������� �����
    frmDemo.Enabled = False
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            '������ ��� ����� "������� ��������� "(� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
    gFileDummy = FreeFile
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            '��������� ������� � ������ ����� DUMMY �����(�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '������� "������" DUMMY ����, ���� �� ����������
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            '��������� ������
    On Error GoTo UnDefError
            '������� DUMMY ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            '������� �����  ��������� ������ DUMMY �����
    gDummyRowNum = 1
            
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
            frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmDemo.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    � DUMMY ����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmDemo.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������"
            Close intFileNum
                        
            '���������������� ������� - "���������� ��������"
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "TableProtocol.dat"

            '�������� ������ � ���� "������� ���������"
            frmDemo.WriteProtocol
                    
        End If
            
            ' ���� ��� ������������
    Else
            
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
            frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmDemo.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    � DUMMY ����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmDemo.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������"
            Close intFileNum
                        
            '���������������� ������� - "���������� ��������"
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "TableProtocol.dat"

            '�������� ������ � ���� "������� ���������"
            frmDemo.WriteProtocol
                    
        End If
    
    End If
            
            '���������� �������������� "����" � ��������
            '  ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            '��������� ������� �������� "Data" ������� � "���� �����������"
    frmDemo.datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
    frmDemo.datBase.RecordSource = "BookKeeping"
            
            '���������� ���������� ������� � "���� �����������"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngBookKeepingBaseCount = frmDemo.datBase.Recordset.RecordCount
            '�������� "���� �����������"
    frmDemo.datBase.Recordset.MoveFirst
            '������� ����� ����������������� ������ "���� �����������"
    lngBookKeepingRowNum = 0
    For lngRowDummy = 0 To gDummyRowNum - 1 Step 1
            '��������� ���������� ��� ��������� ��������� �������
        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
        frmDemo.MousePointer = vbHourglass
            '������� ������ ������������ "���������" ������
        If lngRowDummy = 0 Then
            '��������������� ������� ������ "���� �����������"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = "Fiktive Record"
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = "0000000000000000"
            frmDemo.datBase.Recordset.Fields("Status").Value = "00"
            frmDemo.datBase.Recordset.Fields("Time").Value = "00:00:00AM"
            frmDemo.datBase.Recordset.Fields("Date").Value = "01.01.2000"
            '���������� ������ � "���� �����������"
            frmDemo.datBase.Recordset.Update
        Else
            '������ ������ DUMMY ����� � �����
            Get gFileDummy, lngRowDummy, gProtocol
            '��������������� ������� ������ "���� �����������"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = gProtocol.strProtocName
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = gProtocol.strProtocPersonCode
            frmDemo.datBase.Recordset.Fields("Status").Value = Left(Trim(gProtocol.strProtocStatus), 2)
            frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            '������� ���������:
            '                                  - ����/����� ("18"/"19") ��������� ���
            '                                  - ��������������� ("16) ��������� ���
            '                                  - ������������ ("17") ��������� ���
            '                                  - ����������� ("12") �������� ������� ����������� ���
            '                                  - ���������� ("13") �������� ������� �����������
            '                                  - ����������� ("14") �������� ���������� ����������� ���
            '                                  - ���������� ("15") �������� ���������� �����������
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
            '�������  - ��������������� ��������� (���������� �������
            '  � "���� �����������")
                If Trim(gProtocol.strProtocReserve) = "AutoRegistration" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "16"
            '�������  - ������������ ��������� (���������� �������
            '  � "���� �����������")
                ElseIf Trim(gProtocol.strProtocReserve) = "AutoDelete" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "17"
            '�������  - ���� ��������� �� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "18"
            '�������  - ����� ��������� � ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 6) = "Output" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "19"
            '�������  - ����������� ������� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "12"
            '�������  - ���������� ������� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "13"
            '�������  - ����������� ���������� ����������� (����������
            '  ������� � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "14"
            '�������  - ���������� ���������� ����������� (����������
            '  ������� � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "15"
                End If
            '���������� ������ � "���� �����������"
                frmDemo.datBase.Recordset.Update
            '������� ����� ����������������� ������ "���� �����������"
                lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            '�� ��������� ������ ������ "���� �����������"
                If lngBookKeepingRowNum < lngBookKeepingBaseCount Then
                    frmDemo.datBase.Recordset.MoveNext
            '��������� ������ ������ "���� �����������"
                Else
                    frmDemo.datBase.Recordset.AddNew
                    frmDemo.datBase.Recordset.Update
                    frmDemo.datBase.Recordset.MoveNext
                End If
            End If
        End If
    Next
            '������� ����� ������ "���� �����������"
    lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            '�������� ����� ������ ������ ��  "���� �����������"
    If lngBookKeepingRowNum > lngBookKeepingBaseCount Then
        frmDemo.datBase.Recordset.Delete
            '�������� ������ ������� ��  "���� �����������",
            '  ����� ������������
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum = 1 Then
        frmDemo.datBase.Recordset.MoveFirst
        frmDemo.datBase.Recordset.MoveNext
        For lngBookKeepingRowNum = 2 To lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
        Next
            '�������� ������ ������� ��  "���� �����������"
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum <> 1 Then
        For lngBookKeepingRowNum = lngBookKeepingRowNum To _
        lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
        Next
    End If
            
            '���������������� ������� - "������������ ���� ���������"
    gProtocol.strProtocName = "BookKeeperBase"
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Creation"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol
            
    GoTo EndProcedure
            '�������������� ������
UnDefError:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            '������� DUMMY ����
    Close gFileDummy
    On Error GoTo 0
            
            '������������ ����������� ������ ����
    frmDemo.MousePointer = 0
            '������� ���������� �������� ���������� �����
    frmDemo.Enabled = True
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    

End Sub
            
            '��������� ������ ������� "Form Protocol Base" ���� "File"
Private Sub mnuFormProtocolBase_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '���������� ����� � "���� ���������"
Dim lngProtocolBaseCount As Long
            '����� ����� ������
Dim intFileNum As Integer
            '����� ������ "������� ���������" � DUMMY �����
Dim lngRecordLen As Long
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
Dim strDummyFileName As String
            '������� ����� ������ ������� DUMMY �����
Dim lngRowDummy As Long
            '������ ��� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ��� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � DUMMY ����
Dim intDayArchive As Integer
            '���������� ����� � ���������� ����� (������ ��� "TableProtocol")
Dim intRowQuan As Integer
            '������� ����� ������ ����������� ������
            '   ��� ������� "TableProtocol"
Dim intRowNumArchive As Integer
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmDemo.MousePointer = vbHourglass
            '������� ������������ �������� ���������� �����
    frmDemo.Enabled = False
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            '������ ��� ����� "������� ��������� "(� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
    gFileDummy = FreeFile
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            '��������� ������� � ������ ����� DUMMY �����(�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '������� "������" DUMMY ����, ���� �� ����������
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            '��������� ������
    On Error GoTo UnDefError
            '������� DUMMY ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            '������� �����  ��������� ������ DUMMY �����
    gDummyRowNum = 1
            
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
            frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmDemo.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    � DUMMY ����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmDemo.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������"
            Close intFileNum
                        
            '���������������� ������� - "���������� ��������"
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "TableProtocol.dat"

            '�������� ������ � ���� "������� ���������"
            frmDemo.WriteProtocol
                    
        End If
            
            ' ���� ��� ������������
    Else
            
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
            frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum _
                Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmDemo.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
                    
        Next
            
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    � DUMMY ����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                frmPreprocessors.WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmDemo.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������"
            Close intFileNum
                        
            '���������������� ������� - "���������� ��������"
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "TableProtocol.dat"

            '�������� ������ � ���� "������� ���������"
            frmDemo.WriteProtocol
                    
        End If
    
    End If
            
            '���������� �������������� "����" � ��������
            '  ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            '��������� ������� �������� "Data" ������� � "���� ���������"
    frmDemo.datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
    frmDemo.datBase.RecordSource = "Protocol"
            '���������� ���������� ������� � "���� ���������"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngProtocolBaseCount = frmDemo.datBase.Recordset.RecordCount
            '�������� "���� ���������"
    frmDemo.datBase.Recordset.MoveFirst
            '���� �� ���� ������� DUMMY �����
    For lngRowDummy = 1 To gDummyRowNum - 1 Step 1
            '��������� ���������� ��� ��������� ��������� �������
        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
        frmDemo.MousePointer = vbHourglass
            '������ ������ DUMMY ����� � �����
        Get gFileDummy, lngRowDummy, gProtocol
            '�������� ������� ������ "���� ���������"
        frmDemo.datBase.Recordset.Edit
        frmDemo.datBase.Recordset.Fields("Name").Value = gProtocol.strProtocName
        frmDemo.datBase.Recordset.Fields("CodeOrPassword").Value = _
        gProtocol.strProtocPersonCode
        frmDemo.datBase.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
        frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
        frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
        frmDemo.datBase.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
        frmDemo.datBase.Recordset.Update
            '�� ��������� ������ ������ "���� ���������"
        If lngRowDummy < lngProtocolBaseCount Then
            frmDemo.datBase.Recordset.MoveNext
            '��������� ������ ������ "���� ���������"
        Else
            frmDemo.datBase.Recordset.AddNew
            frmDemo.datBase.Recordset.Update
            frmDemo.datBase.Recordset.MoveNext
        End If
    Next
            '�������� ����� ������ ������ ��  "���� ���������"
    If lngRowDummy > lngProtocolBaseCount Then
        frmDemo.datBase.Recordset.Delete
            '�������� ������ ������� ��  "���� ���������"
    Else
        For lngRowDummy = lngRowDummy To lngProtocolBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmDemo.MousePointer = vbHourglass
        Next
    End If
            
            '���������������� ������� - "������������ ���� ���������"
    gProtocol.strProtocName = "ProtocolBase"
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Creation"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol
    
    GoTo EndProcedure
            '�������������� ������
UnDefError:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            '������� DUMMY ����
    Close gFileDummy
    On Error GoTo 0
            
            '������������ ����������� ������ ����
    frmDemo.MousePointer = 0
            '������� ���������� �������� ���������� �����
    frmDemo.Enabled = True
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    
End Sub

            ' ��������� ��������� ������������� "����������/���������"
Private Sub chkSetup_Click()
            ' ���������� ��������� ������������
    If chkSetup.Value = 0 Then
            ' ����� "���������" - ������� ���������� �������� ����������
        
            ' ��������� �����/�������� ���������� ��� ����������
        If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
        If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
        If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
        If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            '��������� ����� "Controller'��" �� ��������
        gTermContr = 0
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = False
            imgEmployeOutData.Enabled = False
            imgEmployeInfoData.Enabled = False
        End If
            
            ' ������� ������������ ������ ������� ���������� �����������
        cmdOpen(0).Enabled = False
        cmdOpen(1).Enabled = False
        cmdOpen(2).Enabled = False
        cmdOpen(3).Enabled = False
            '���������������� ������� - ������ ����� �� ������������� "Execute/Setup"
        gProtocol.strProtocName = "????????????????"
            '��������� ������
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "SETUP option"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            '���������� ����� �� ���� ������
        txtPassword.Enabled = True
        txtPassword.SetFocus
            '���������� �������� ������� ����� ������
        tmrPasswTimeOut.Enabled = True
            '��������� ������ ���������� �� ���� ������ �� ��� �����
        Do While txtPassword.Enabled = True
            DoEvents
        Loop
        
            '������� "TimeOut" ��� ����� ������ - ���������� ���������� ���������
        If tmrPasswTimeOut.Enabled = False Then
            ' ����� "����������" - ������� ������������ �������� ����������
            GoTo Execution
        End If
        
            '�������� �������� ������� ����� ������
        tmrPasswTimeOut.Enabled = False
        
            '������� ��������� ������������� "����������/���������"
        chkSetup.Enabled = True
        
            
            '������� ��������� ���� ������ - ��� ���������� ����� ������ ������
        txtPassword.Enabled = True
        
            '������� ������� ����
        mnuFile.Visible = True
        mnuAdjustment.Visible = True
        mnuParking.Visible = True
        mnuAccess.Visible = True
        mnuEmploye.Visible = True
            '������� ������� ������ ������������
        picTools.Visible = True
            ' ������� ��������� ����� ����� �������
        fraFlag.Enabled = True
        optEnglish.Enabled = True
        optLatvian.Enabled = True
        optRussian.Enabled = True
            ' ������� ���������� ������������� ���������� ����������
        chkTerm(0).Enabled = True
        chkTerm(1).Enabled = True
        chkTerm(2).Enabled = True
        chkTerm(3).Enabled = True
            ' "��������" �������� ���������� ����������
        lblTerminals.Enabled = True
           ' ������� ���������� ������������� ���������������
        chkPhoto(0).Enabled = True
        chkPhoto(1).Enabled = True
        chkPhoto(2).Enabled = True
        chkPhoto(3).Enabled = True
            ' "��������" �������� ������������� ���������������
        lblPhoto.Enabled = True
            ' ������� ���������� ����� "��������������/������" ���������� �����������
        fraControl.Enabled = True
        optAutomatic.Enabled = True
        optManual.Enabled = True
               ' ������� ������������ ������ ������� ���������� �����������
        cmdOpen(0).Enabled = False
        cmdOpen(1).Enabled = False
        cmdOpen(2).Enabled = False
        cmdOpen(3).Enabled = False
                ' "��������" �������� ������ ������� ���������� �����������
        lblOpen.Enabled = False


    Else
            ' ����� "����������" - ������� ������������ �������� ����������
Execution:
            '������� ����������� ���� ������
        txtPassword.Enabled = False
        
        '���������������� ������� - ������ ����� �� ������������� "Execute/Setup"
        gProtocol.strProtocName = "????????????????"
            '��������� ������
        gProtocol.strProtocPersonCode = txtPassword.Tag
            '������
        gProtocol.strProtocStatus = "04 - Manager"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "EXECUTE option"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
        
            '������� ��������� ����
        mnuFile.Visible = False
        mnuAdjustment.Visible = False
        mnuParking.Visible = False
        mnuAccess.Visible = False
        mnuEmploye.Visible = False
            '������� ��������� ������ ������������
        picTools.Visible = False
    ' ������� ����������� ����� ����� �������
        fraFlag.Enabled = False
        optEnglish.Enabled = False
        optLatvian.Enabled = False
        optRussian.Enabled = False
            ' ������� ������������ ������������� ���������� ����������
        chkTerm(0).Enabled = False
        chkTerm(1).Enabled = False
        chkTerm(2).Enabled = False
        chkTerm(3).Enabled = False
            ' "��������" �������� ���������� ����������
        lblTerminals.Enabled = False
           ' ������� ������������ ������������� ���������������
        chkPhoto(0).Enabled = False
        chkPhoto(1).Enabled = False
        chkPhoto(2).Enabled = False
        chkPhoto(3).Enabled = False
            ' "��������" ���� ���������������
        imgPhoto(0).Picture = LoadPicture("")
        imgPhoto(1).Picture = LoadPicture("")
        imgPhoto(2).Picture = LoadPicture("")
        imgPhoto(3).Picture = LoadPicture("")
            ' "��������" �������� ������������� ���������������
        lblPhoto.Enabled = False
            ' ������� ����������� ����� "��������������/������" ���������� �����������
        fraControl.Enabled = False
        optAutomatic.Enabled = False
        optManual.Enabled = False
    
            '���� ������� ����� "��������������" ���������� �����������
        If optAutomatic.Value = True Then
            '������������ �������� ��������� � ������� ����������� "������"
            cmdOpen(0).Tag = 0
            cmdOpen(0).Caption = chkTerm(0).Caption
            cmdOpen(1).Tag = 0
            cmdOpen(1).Caption = chkTerm(1).Caption
            cmdOpen(2).Tag = 0
            cmdOpen(2).Caption = chkTerm(2).Caption
            cmdOpen(3).Tag = 0
            cmdOpen(3).Caption = chkTerm(3).Caption
               ' ������� ������������ ������ ������� ���������� �����������
            cmdOpen(0).Enabled = False
            cmdOpen(1).Enabled = False
            cmdOpen(2).Enabled = False
            cmdOpen(3).Enabled = False
                ' "��������" �������� ������ ������� ���������� �����������
            lblOpen.Enabled = False
                '���� ������� ����� "������" ���������� �����������
        Else
            '������ "Controller'��", ����������� �� ����������� "������"
            cmdOpen(0).Tag = CByte(CInt(Trim(gAddrManual(0))))
            cmdOpen(0).Caption = Trim(gAddrManual(0))
            cmdOpen(1).Tag = CByte(CInt(Trim(gAddrManual(1))))
            cmdOpen(1).Caption = Trim(gAddrManual(1))
            cmdOpen(2).Tag = CByte(CInt(Trim(gAddrManual(2))))
            cmdOpen(2).Caption = Trim(gAddrManual(2))
            cmdOpen(3).Tag = CByte(CInt(Trim(gAddrManual(3))))
            cmdOpen(3).Caption = Trim(gAddrManual(3))
               ' ������� ���������� ������ ������� ���������� ���������� ����������
            If chkTerm(0).Value = 1 Then cmdOpen(0).Enabled = True
            If chkTerm(1).Value = 1 Then cmdOpen(1).Enabled = True
            If chkTerm(2).Value = 1 Then cmdOpen(2).Enabled = True
            If chkTerm(3).Value = 1 Then cmdOpen(3).Enabled = True
                ' "��������" �������� ������ ������� ���������� �����������
            lblOpen.Enabled = True
            
        End If
        
            ' ��������� �����/�������� ���������� ��� ���������� ����������
            ' ���������� ����� "��������" - ����� ������������ ������
            ' ���������� ����������
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
            '��������� ����� "Controller'��" �� ��������
        gTermContr = 1
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
    
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = True
            imgEmployeOutData.Enabled = True
            imgEmployeInfoData.Enabled = True
        End If
            
            '������� ��������� ������������� "����������/���������"
        chkSetup.Enabled = True
            '���������� ����� "����������"
        chkSetup.Value = 1
            ' ��������� ����� �� ��������� ����������
        cmdExit.Enabled = True
            '���������� ����� �� ����� "Dummy"
        If frmDemo.Visible = True Then chkDummy.SetFocus

    End If

End Sub

            ' ������ ������ "Exit"
Private Sub cmdExit_Click()
            '��� �������� ��� ���������� "������� ������"
Dim intSaveTablePerson As Integer
    
            ' ��������� �����/�������� ���������� ��� ����������
    If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
    If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
    If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
    If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = False
            imgEmployeOutData.Enabled = False
            imgEmployeInfoData.Enabled = False
        End If
            
            ' ������� ������������ ������ ������� ���������� �����������
    cmdOpen(0).Enabled = False
    cmdOpen(1).Enabled = False
    cmdOpen(2).Enabled = False
    cmdOpen(3).Enabled = False
            '���������������� ������� - ������ ������ "Exit"
    gProtocol.strProtocName = "????????????????"
            '��������� ������
    gProtocol.strProtocPersonCode = ""
            '������
    gProtocol.strProtocStatus = ""
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "EXIT button"
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
    
            '���������� ����� �� ���� ������
    txtPassword.Enabled = True
    txtPassword.SetFocus
            '���������� �������� ������� ����� ������
    tmrPasswTimeOut.Enabled = True
            '��������� ������ ���������� �� ���� ������ �� ��� �����
    Do While txtPassword.Enabled = True
        DoEvents
    Loop
    
            '������� "TimeOut" ��� ����� ������ - ���������� ���������� ���������
    If tmrPasswTimeOut.Enabled = False Then
            ' ����� "����������" - ������� ������������ �������� ����������
            
            ' "��������" ���� ���������������
        imgPhoto(0).Picture = LoadPicture("")
        imgPhoto(1).Picture = LoadPicture("")
        imgPhoto(2).Picture = LoadPicture("")
        imgPhoto(3).Picture = LoadPicture("")
    
            '���� ������� ����� "��������������" ���������� �����������
        If optAutomatic.Value = True Then
               ' ������� ������������ ������ ������� ���������� �����������
            cmdOpen(0).Enabled = False
            cmdOpen(1).Enabled = False
            cmdOpen(2).Enabled = False
            cmdOpen(3).Enabled = False
                ' "��������" �������� ������ ������� ���������� �����������
            lblOpen.Enabled = False
                    '���� ������� ����� "������" ���������� �����������
        Else
               ' ������� ���������� ������ ������� ���������� ���������� ����������
            If chkTerm(0).Value = 1 Then cmdOpen(0).Enabled = True
            If chkTerm(1).Value = 1 Then cmdOpen(1).Enabled = True
            If chkTerm(2).Value = 1 Then cmdOpen(2).Enabled = True
            If chkTerm(3).Value = 1 Then cmdOpen(3).Enabled = True
                ' "��������" �������� ������ ������� ���������� �����������
            lblOpen.Enabled = True
        End If
        
            ' ��������� �����/�������� ���������� ��� ���������� ����������
            ' ���������� ����� "��������" - ����� ������������ ������ ���������� ����������
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
        
            '��������� ����� "Controller'��" �� ��������
        gTermContr = 1
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
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
            
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������)
        If imgEmployeInfoData.Visible = True Then
            imgEmployeInData.Enabled = True
            imgEmployeOutData.Enabled = True
            imgEmployeInfoData.Enabled = True
        End If
            
            '������� ��������� ������������� "����������/���������"
        chkSetup.Enabled = True
            ' ��������� ����� �� ��������� ����������
        cmdExit.Enabled = True
            '���������� ����� �� ����� "Dummy"
        If frmDemo.Visible = True Then chkDummy.SetFocus
        
            '������ ������ �� ���������� �����
    Else
            '���������� ������� ������������� ������ "������� ������":
            '   ��������������� ������ � "Host Computer'e" � � ��� �������
            '   � "�������������", ����� ��������� ���������� ����
            '   ����������� "������� ������" - "���������� ������a ������"
        If gCompresTablPers = 1 Then
            '���������� ������� ��������� ��������� � "������� ������"
            '  - ��������� ������� � ������������ �����
            If gChangesTablePerson = True Then _
                Call frmTablePerson.SaveTablePerson
            ' ���� ������� ������������� � ��������� ����
            If gNetPreprocNum > 0 Then
            '������ �������� ���������
                strMessage = "ExitApp"
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                Call SendMessage(strMessage)
            End If
        End If
            
            '���������������� ������� - "��������� �������"
        gProtocol.strProtocName = "################"
            '��������� ������
        gProtocol.strProtocPersonCode = txtPassword.Tag
            '������
        gProtocol.strProtocStatus = "04 - Manager"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Unload the Acc. Syst."
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            '��� �������� ������� ������������ ��� ��������� � �����������
        If gMSBase = 1 Then
            '������������ ��� ��������� � ����������� � ������� ACCESS"
            Call BasesConvert
        End If
            
            '������� ���� "������� ���������"
        Close gProtocFileNum
    
            ' ���������� ������ �� ������ "FlexGrid" ("������� ������")
        Set gTablePerson = Nothing
            ' ���������� ������ �� ������ ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' ���������� ������ �� ������ MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ����������� ���������
        Set qQueueInput = Nothing
            ' ���������� ������ �� ������ �������-�������
            ' ����������� ���������
        Set evQueue = Nothing
            ' ���������� ������ �� ������ ����������� ��������E
        Set qMsgInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ������������ ���������
        Set qQueueOutput = Nothing
            ' ���������� ������ �� ������ ������������ ��������E
        Set qMsgOutput = Nothing
            
            '��������� ���������
        End
    End If
    
End Sub

            ' ��������� �����
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '������������� �������� � ����� ��������
    lngResource = 101
            ' ���������������� ������ ��������
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
            '������� ��������� ����
    mnuFile.Visible = False
    mnuAdjustment.Visible = False
    mnuParking.Visible = False
    mnuAccess.Visible = False
    mnuEmploye.Visible = False
            '������� ��������� ������ ������������
    picTools.Visible = False
    
            ' ���������������� ������ "�����������" ���������
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
            '���������� ����� ����� ������� � ���������� ����
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
            '����� ������ ����� �������
    intLang = fraFlag.Tag
            ' ���������������� ������ ��������
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
            ' ���������������� ������ "�����������" ���������
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
        
            
            ' �������� ������� MSMQQueueInfo ��� ����������
            '  �������� ����������� ���������
    Set qInfoInput = New MSMQQueueInfo
            ' ���������� ���� � ������� ����������� ���������
    qInfoInput.PathName = ".\Private$\GeneralQueue"
            ' ��������� ��� ������� ����������� ���������
    qInfoInput.Label = "Input Message Queue"
    On Error Resume Next
            ' ������� ������� ����������� ���������
            '   �����������
    qInfoInput.Create
    On Error GoTo 0
            ' ������� ������� ��������� � ����������� (��� ������
            '   ���������, ������ � ������� �������� ����)
    Set qQueueInput = qInfoInput.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
            ' ������� ��������� ������� MSMQEvent � ��� �����������
    Set qEvent = New MSMQEvent
    qQueueInput.EnableNotification qEvent
            
            
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� ����� "������� ���������"
    gProtocFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '���� "������� ���������" ������� �������������
    If FileLen(strPathFileName) / lngRecordLen + 1 > 32760 Then
            '������ �������� ������
        BeepSound
        If optEnglish = True Then
            strResponse = MsgBox("The protocol overflow ?", vbYesNo + vbQuestion, "Cancel")
        Else
            strResponse = MsgBox("Protokols ir izpildits ?", vbYesNo + vbQuestion, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '����� ��������� ������ "������� ���������" (������ �������)
            gProtocRowNum = 32760
            '������ ������ "��"
        Else
            '������������� ����� "������� ���������"
            WriteProtocolToArchives
        End If
            '���� "������� ���������" �� ������� �������������
    Else
            '����� ������ ��������� ������ "������� ���������"
        gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
    End If
            '���������������� ������� - "����� �������"
    gProtocol.strProtocName = "****************"
            '��������� ������
    gProtocol.strProtocPersonCode = txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Restart the Acc. Syst."
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
     
End Sub

            '����� ������� "mnuCalendar" ���� "Adjustment"
Private Sub imgCalendar_Click()
    mnuCalendar_Click
End Sub

            '����� ������� "mnuPersons" ���� "Adjustment"
Private Sub imgPersons_Click()
    mnuPersons_Click
End Sub

            '����� ������� "mnuParkingInData" ���� "Parking"
Private Sub imgParkingIn_Click()
    mnuParkingInData_Click
End Sub

            '����� ������� "mnuParkingOutData" ���� "Parking"
Private Sub imgParkingOut_Click()
    mnuParkingOutData_Click
End Sub

            '����� ������� "mnuParkingInfoData" ���� "Parking"
Private Sub imgParkingInfo_Click()
    mnuParkingInfoData_Click
End Sub

            '����� ������� "mnuParkingServData" ���� "Parking"
Private Sub imgParkingServ_Click()
    mnuParkingServData_Click
End Sub

            '����� ������� "mnuAccessInData" ���� "Access"
Private Sub imgAccessIn_Click()
    mnuAccessInData_Click
End Sub

            '����� ������� "mnuAccessOutData" ���� "Access"
Private Sub imgAccessOut_Click()
    mnuAccessOutData_Click
End Sub

            '����� ������� "mnuAccessInfoData" ���� "Access"
Private Sub imgAccessInfo_Click()
    mnuAccessInfoData_Click
End Sub

            '����� ������� "mnuAccessServData" ���� "Access"
Private Sub imgAccessServ_Click()
    mnuAccessServData_Click
End Sub

            '����� ������� "mnuPrint..." ���� "File"
Private Sub imgPrint_Click()
    mnuPrint_Click
End Sub

            '����� ������� "mnuFormProtocolBase" ���� "File"
Private Sub imgProtocolBase_Click()
    mnuFormProtocolBase_Click
End Sub

            '����� ������� "mnuFormBookKeeperBase" ���� "File"
Private Sub imgBookKeeperBase_Click()
    mnuFormBookKeeperBase_Click
End Sub

            '����� ������� "mnuProtocolToArchives..." ���� "File"
Private Sub imgProtocArchives_Click()
    mnuProtocolToArchives_Click
End Sub

            '����� ������� "mnuSaveProtocol" ���� "Adjustment"
Private Sub imgSaveProtocol_Click()
    mnuSaveProtocol_Click
End Sub

            '����� ������� "mnuSystem" ���� "Adjustment"
Private Sub imgSystem_Click()
    mnuSystem_Click
End Sub

            '����� ������� "mnuTerminal" ���� "Adjustment"
Private Sub imgTerminal_Click()
    mnuTerminal_Click
End Sub

            '����� ������� "mnuTime" ���� "Adjustment"
Private Sub imgTime_Click()
    mnuTime_Click

End Sub

            '����� ������� "mnuPreprocessors" ���� "Adjustment"
Private Sub imgPreprocessors_Click()
    mnuPreprocessors_Click
    
End Sub

            '��������� ������ ������� "Preprocessors" ���� "Adjustment"
Private Sub mnuPreprocessors_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer

            '��������� ������� �������� ������ ��� "Host Computer'a"
    If gPreprocName <> "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If

             '��������� (�� ���������) ����� "frmPreprocessors"
    Load frmPreprocessors
    
            '������������� ���� � ������� (� "��������� �������")
    If frmPreprocessors.cboFileName.ListCount <> 0 Then
            '������� �� ����� ����� "frmPreprocessors"
            '   � ������� ����������� 1
        frmPreprocessors.Show 1
    End If
    
            '��������� ����� "frmPreprocessors"
    UnLoad frmPreprocessors
            '���������� ������, ���������� ����������� ������
    Set frmPreprocessors = Nothing
            '������������ ����������� ������ ����
    frmDemo.MousePointer = 0
            '������� ���������� �������� ���������� �����
    frmDemo.Enabled = True
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then frmDemo.chkDummy.SetFocus

End Sub

            '��������� ������ ������� "Protocol to Archives..." ���� "File"
Private Sub mnuProtocolToArchives_Click()
            '������ ��� ����� ������ (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ��������� �����
Dim intFileNum As Integer
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer

            '��������� (�� ���������) ����� "frmGetFile"
    Load frmGetFile
            '��������� ������ ���������������� ���� "cboFileType
    frmGetFile.cboFileType.AddItem "All files (*.*)"
    frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
    frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            '������� ������� ������ "��� �����"
    frmGetFile.cboFileType.ListIndex = 0
    
            '������������ ������������ ��� ������������� �����
            
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            '������ ��� ����� (� ��������� "����" � ����)
        frmGetFile.txtFileName = gHost + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
            ' ���� ��� ������������
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        frmGetFile.txtFileName = gPreprocName + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
    End If
    
    
            '������� �� ����� ����� "frmGetFile" � ������� ����������� 1
    frmGetFile.Show 1
            '���� �� ������
    If frmGetFile.Tag = "" Then
            '������ �������� ������
        BeepSound
        MsgBox "The file isn't selected !"
            '������ "������� ���������" � �������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� ���������"
        lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '��������� ������� � ������ ����� ����� (�� ��������� "C:\")
        intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
        Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
        Loop
            '������� "������" �������� ����, ���� �� ����������
        If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
            Kill strPathFileName
        End If
        
            '��������� ������
                On Error GoTo UnDefError
            '������� ��������� �������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
    
            '���� �� ���� ������� "������� ���������"
        For intRowNum = 1 To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, intRowNum, gProtocol
            '������� ������ "������� ���������" � �������� ����
            Put intFileNum, intRowNum, gProtocol
        Next
            '������� ���������  �������� ����
        Close intFileNum
            
            ' ���� ��� ������������
        If gPreprocName <> "" Then
            '������������ ����������� ���������
            strMessage = "Archive" + " " + Mid(strPathFileName, intSymbPos)
            '�������� ���������
            Call SendMessage(strMessage)
        End If
        
             '������� "�������" ���� "������� ���������"
        Close gProtocFileNum
           '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableProtocol.dat"
        
            '��������� ������� � ������ ����� ����� (�� ��������� "C:\")
        intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
        Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
        Loop
            '�������  "�������" ���� "������� ���������"
        If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
            Kill strPathFileName
        End If
            '�������� ��������� ����� ����� ��� ����� "������� ���������"
        gProtocFileNum = FreeFile
            '������� ����� ���� "������� ���������" ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '����� ��������� "���������" ������ � ����� ����� "������� ���������"
        gProtocRowNum = 1
            '���������� ������� ���������� ��������� ������� � ������������ �����
        mnuSaveProtocol.Checked = True
        mnuSaveProtocolAs.Checked = False
    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
    
    Exit Sub
            '�������������� ������
UnDefError:
            '������ �������� ������
    BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing

End Sub
            
            '��������� ������ ������� "Print..." ���� "File"
Private Sub mnuPrint_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
    Dim strPathFileName As String
            '���������� ����� � "���� �����������"
    Dim intBookKeepingBaseCount As Integer
            '���������� ����� � "���� ���������"
    Dim intProtocolBaseCount As Integer
            '��� �������, ��������� ��� ������
    Dim strTableName As String
            '������� ����� ������ ����� "frmPrintPreview"
    Dim intRowPrintNum As Integer
            '���������� ����� �� ����� �������� ����� "frmPrintPreview"
    Dim intRowPrintQuan As Integer
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    Dim intRowNum As Integer
            '������� ����� ������� ������ ("TablePerson", "TableCalendar", "TableProtocol"
            '  "TableSystem", "TableTime", "TableTerminal")
    Dim intColNum As Integer
            '����� ������ ������ "��������� �������"
    Dim strTableSystem(5) As String
            '����� ������ ������ "������� ������"
    Dim strTablePerson(6) As String
            '����� ������ ������ "������� ���������"
    Dim strTableCalendar(8) As String
            '����� ������ ������ "������� �������"
    Dim strTableTime(3) As String
            '����� ������ ������ "������� ����������"
    Dim strTableTerminal(4) As String
    
            '��������� (�� ���������) ����� "frmSelectRow"
    Load frmSelectRow
            '���������������� �������� "lblColName" ����� "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Table type"
    
             '�������� ������ ��������
    frmSelectRow.lstSelectRow.Clear
            '���������� ������ "lstSelectRow"
    frmSelectRow.lstSelectRow.AddItem "TableSystem"
    frmSelectRow.lstSelectRow.AddItem "TablePerson"
    frmSelectRow.lstSelectRow.AddItem "TableCalendar*"
    frmSelectRow.lstSelectRow.AddItem "TableProtocol"
    frmSelectRow.lstSelectRow.AddItem "TableTime*"
    frmSelectRow.lstSelectRow.AddItem "TableTerminal*"
    frmSelectRow.lstSelectRow.AddItem "BookKeepingBase"
    frmSelectRow.lstSelectRow.AddItem "ProtocolBase"
            '������� ������� ������
    frmSelectRow.lstSelectRow.ListIndex = 0
            '������� �� ����� ����� "frmSelectRow" � ������� ����������� 1
    frmSelectRow.Show 1
            '������ �� �������
    If frmSelectRow.Tag = "" Then
            '������ �������� ������
        BeepSound
        MsgBox "The table isn't selected !"
            '��������� ����� "frmSelectRow"
        UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
        Set frmSelectRow = Nothing
            '������� ������� ��� ���������
    Else
            '��� �������, ��������� ��� ������
        strTableName = frmSelectRow.Tag
            '��������� ����� "frmSelectRow"
        UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
        Set frmSelectRow = Nothing
              
            '������� ����� ������ �� �������� ������
        intRowPrintNum = 1
            '���������� ����� �� ����� �������� ������
        intRowPrintQuan = gRowPrintQuan
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal") ��� ������� ���("ProtocolBase",
            '  "BookKeepingBase")
        intRowNum = 1
            
            '������� ������� "�� ��������� - �������"
        Set Printer = Printers(0)
            '�������� ������� �� "��������" ���������� ������
        Printer.EndDoc
            '������� 3-� ������ ������
        Printer.CurrentY = 4
            '������ ������ ��������
        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
        Printer.Print
    
            '������ "��������� �������"
        If strTableName = "TableSystem" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(45); "Type"; _
            Tab(70); "Index"; Tab(95); "Appendix"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ��������������� ������� "��������� �������"
            For intRowNum = intRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
                frmTableSystem.grdTableSystem.Row = intRowNum
            '�� ���� �������� "��������� �������"
                For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
                    frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
                    strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
                Next
            '������� �� ������ ������ "��������� �������"
                Printer.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
                Tab(45); strTableSystem(2); Tab(70); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "��������� �������"
                    If intRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(45); "Type"; _
                        Tab(70); "Index"; Tab(95); "Appendix"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "��������� �������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
    
            '������ "������� ������"
        ElseIf strTableName = "TablePerson" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ��������������� ������� "������� ������"
            For intRowNum = intRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
                gTablePerson.Row = intRowNum
            '�� ���� �������� "������� ������"
                For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
                    gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
                    strTablePerson(intColNum) = gTablePerson.Text
                Next
            '������ - ������ �����������
                If Left(Trim(strTablePerson(2)), 2) = "07" Or _
                Left(Trim(strTablePerson(2)), 2) = "05" Or _
                Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
                    strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
                ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
                Left(Trim(strTablePerson(2)), 2) = "08" Or _
                Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
                    strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
                End If
            '������� �� ������ ������ "������� ������"
                Printer.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
                Tab(45); strTablePerson(2); Tab(70); strTablePerson(3); Tab(95); strTablePerson(4); _
                Tab(115); strTablePerson(5)
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "������� ������"
                    If intRowNum < gTablePerson.Rows - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "������� ������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
    
            '������ "������� ���������"
        ElseIf strTableName = "TableCalendar*" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
            Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
            Tab(115); "Sunday"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ��������������� ������� "������� ���������"
            For intRowNum = intRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
                frmTableCalendar.grdTableCalendar.Row = intRowNum
            '�� ���� �������� "������� ���������"
                For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
                    frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
                    strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
                Next
            '������� �� ������ ������ "������� ���������"
                Printer.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
                Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
                Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "������� ���������"
                    If intRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
                        Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
                        Tab(115); "Sunday"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "������� ���������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
        
            '������ "������� ���������"
        ElseIf strTableName = "TableProtocol" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ������� "������� ���������"
            For intRowNum = intRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
                Get gProtocFileNum, intRowNum, gProtocol
            '������� ������ "������� ���������"
                Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(45); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(85); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "������� ���������"
                    If intRowNum < gProtocRowNum - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "������� ���������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
                    
            '������ "���� ���������"
        ElseIf strTableName = "ProtocolBase" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
            Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            '���������� �������������� "����" � �������� ����������� ���������
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            '��������� ������� �������� "Data" ������� � "���� ���������"
            datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
            datBase.RecordSource = "Protocol"
            
            '���������� ���������� ������� � "���� ���������"
            datBase.Refresh
            datBase.Recordset.MoveLast
            intProtocolBaseCount = datBase.Recordset.RecordCount
            '������ "���� ���������"
            datBase.Recordset.MoveFirst
            '���� �� ���� ������� "���� ���������"
            For intRowNum = 1 To intProtocolBaseCount Step 1
            '������� ������� ������ "���� ���������"
                Printer.Print Tab(3); datBase.Recordset.Fields("Name").Value; _
                Tab(25); datBase.Recordset.Fields("CodeOrPassword").Value; _
                Tab(45); datBase.Recordset.Fields("Status").Value; _
                Tab(70); datBase.Recordset.Fields("Time").Value; _
                Tab(85); datBase.Recordset.Fields("Date").Value; _
                Tab(100); datBase.Recordset.Fields("ReservOrNote").Value
                datBase.Recordset.MoveNext
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������� "���� ���������"
                    If intRowNum < intProtocolBaseCount Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(45); "Status"; _
                        Tab(70); "Time"; Tab(85); "Date"; Tab(100); "Reserv. or Note"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "���� ���������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
                    
                    '������ "������� �������"
        ElseIf strTableName = "TableTime*" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(45); "Expander"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ��������������� ������� "������� �������"
            For intRowNum = intRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
                frmTableTime.grdTableTime.Row = intRowNum
            '�� ���� �������� "������� �������"
                For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
                    frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
                    strTableTime(intColNum) = frmTableTime.grdTableTime.Text
                Next
            '������� �� ������ ������ "������� �������"
                Printer.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); Tab(45); strTableTime(2)
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "������� �������"
                    If intRowNum < frmTableTime.grdTableTime.Rows - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(45); "Expander"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "������� �������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
    
            '������ "������� ����������"
        ElseIf strTableName = "TableTerminal*" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(45); "Description"; _
            Tab(70); "Expander"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            
            '���� �� ���� ��������������� ������� "������� ����������"
            For intRowNum = intRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
                frmTableTerminal.grdTableTerminal.Row = intRowNum
            '�� ���� �������� "������� ����������"
                For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
                    frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
                    strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
                Next
            '������� �� ������ ������ "������� ����������"
                Printer.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
                Tab(45); strTableTerminal(2); Tab(70); strTableTerminal(3)
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������ "������� ����������"
                    If intRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(45); "Description"; _
                        Tab(70); "Expander"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "������� ����������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
        
            '������ "���� �����������"
        ElseIf strTableName = "BookKeepingBase" Then
            '������ ���������� ��������
            Printer.Print Tab(3); "Person"; Tab(25); "PersonCode"; Tab(45); "Status"; _
            Tab(55); "Time"; Tab(70); "Date"
            '������� ������ ������
            Printer.Print
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 4
            '���������� �������������� "����" � �������� ����������� ���������
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            '��������� ������� �������� "Data" ������� � "���� �����������"
            datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
            datBase.RecordSource = "BookKeeping"
            
            '���������� ���������� ������� � "���� �����������"
            datBase.Refresh
            datBase.Recordset.MoveLast
            intBookKeepingBaseCount = datBase.Recordset.RecordCount
            '������ "���� �����������"
            datBase.Recordset.MoveFirst
            '���� �� ���� ������� "���� �����������"
            For intRowNum = 1 To intBookKeepingBaseCount Step 1
            '������� ������� ������ "���� �����������"
                Printer.Print Tab(3); datBase.Recordset.Fields("Person").Value; _
                Tab(25); datBase.Recordset.Fields("PersonCode").Value; _
                Tab(45); datBase.Recordset.Fields("Status").Value; _
                Tab(55); datBase.Recordset.Fields("Time").Value; _
                Tab(70); datBase.Recordset.Fields("Date").Value
                datBase.Recordset.MoveNext
            '������� ����� ������ �� �������� ������
                intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ���������
                If intRowPrintNum > intRowPrintQuan Then
            '���������� �� ��� ������� "���� �����������"
                    If intRowNum < intBookKeepingBaseCount Then
            '������ ������ ����� ��������
                        Printer.NewPage
                        Printer.Print Tab(3); "Page " + CStr(Printer.Page)
            '������� ������ ������
                        Printer.Print
            '������ ���������� �������� �� ����� ��������
                        Printer.Print Tab(3); "Person"; Tab(25); "PersonCode"; Tab(45); "Status"; _
                        Tab(55); "Time"; Tab(70); "Date"
            '������� ������ ������
                        Printer.Print
            '������� ����� ������ �� ����� �������� ������
                        intRowPrintNum = 5
            '���������� ��� ������ "���� �����������"
                    Else
            '��������� ������
                        Exit For
                    End If
                End If
            Next
                    
        End If
        
            '������ ��� ������ ������ ���
        Printer.EndDoc
    End If

End Sub

            '��������� ������ ������� "Print preview..." ���� "File"
Private Sub mnuPrintPreview_Click()
            '��������� (�� ���������) ����� "frmSelectRow"
    Load frmSelectRow
            '���������������� �������� "lblColName" ����� "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Table type"
    
             '�������� ������ ��������
    frmSelectRow.lstSelectRow.Clear
            '���������� ������ "lstSelectRow"
    frmSelectRow.lstSelectRow.AddItem "TableSystem"
    frmSelectRow.lstSelectRow.AddItem "TablePerson"
    frmSelectRow.lstSelectRow.AddItem "TableCalendar*"
    frmSelectRow.lstSelectRow.AddItem "TableProtocol"
    frmSelectRow.lstSelectRow.AddItem "TableTime*"
    frmSelectRow.lstSelectRow.AddItem "TableTerminal*"
    frmSelectRow.lstSelectRow.AddItem "ProtocolFromArchives"
            '������� ������� ������
    frmSelectRow.lstSelectRow.ListIndex = 0
            '������� �� ����� ����� "frmSelectRow" � ������� ����������� 1
    frmSelectRow.Show 1
            '������ �� �������
    If frmSelectRow.Tag = "" Then
            '������ �������� ������
        BeepSound
        MsgBox "The table isn't selected !"
            '��������� ����� "frmSelectRow"
        UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
        Set frmSelectRow = Nothing
            '��������� ������� ��� ��������� �� "����� ���������"
    ElseIf frmSelectRow.Tag <> "ProtocolFromArchives" Then
            '��������� (�� ���������) ����� "frmPrintPreview"
        Load frmPrintPreview
            '��� �������, ��������� ��� ��������������� ������
        frmPrintPreview.Tag = frmSelectRow.Tag
            '��������� ����� "frmSelectRow"
        UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
        Set frmSelectRow = Nothing
              '������� �� ����� ����� "frmPrintPreview" � ������� ����������� 1
        frmPrintPreview.Show 1
              '��������� ����� "frmPrintPreview"
        UnLoad frmPrintPreview
            '���������� ������, ���������� ����������� ������
        Set frmPrintPreview = Nothing
            '��������� ������� ��� ��������� "����� ���������"
    ElseIf frmSelectRow.Tag = "ProtocolFromArchives" Then
            '��������� (�� ���������) ����� "frmPrintPreview"
        Load frmPrintPreview
            '��� �������, ��������� ��� ��������������� ������
        frmPrintPreview.Tag = frmSelectRow.Tag
            '��������� ����� "frmSelectRow"
        UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
        Set frmSelectRow = Nothing

            '��������� (�� ���������) ����� "frmGetFile"
        Load frmGetFile
            '��������� ������ ���������������� ���� "cboFileType
        frmGetFile.cboFileType.AddItem "All files (*.*)"
        frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
        frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            '������� ������� ������ "��� �����"
        frmGetFile.cboFileType.ListIndex = 0
            '������� �� ����� ����� "frmGetFile" � ������� ����������� 1
        frmGetFile.Show 1
            '���� �� ������
        If frmGetFile.Tag = "" Then
            '������ �������� ������
            BeepSound
            MsgBox "The file isn't selected !"
            '������  "������ ���������"
        Else
                '������ ��� ����� (� ��������� "����" � ����)
            gPathFileName = frmGetFile.Tag
              '������� �� ����� ����� "frmPrintPreview" � ������� ����������� 1
            frmPrintPreview.Show 1
              '��������� ����� "frmPrintPreview"
            UnLoad frmPrintPreview
            '���������� ������, ���������� ����������� ������
            Set frmPrintPreview = Nothing
        End If
            '��������� ����� "frmGetFile"
        UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
        Set frmGetFile = Nothing
    End If

End Sub

            '���������� ��������� ��� ������ ������� "Exit" ���� "File"
Private Sub mnuExit_Click()
            '��� �������� ��� ���������� "������� ������"
Dim intSaveTablePerson As Integer
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            '���������� ������� ��������� ��������� � "������� ������"
            '  - ��������� ������� � ������������ �����
        If gChangesTablePerson = True Then _
            intSaveTablePerson = frmTablePerson.SaveTablePerson()
            ' ���� ������� ������������� � ��������� ����
        If gNetPreprocNum > 0 Then
            '������ �������� ���������
            strMessage = "ExitApp"
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call SendMessage(strMessage)
        End If
            ' ���� ��� �� "Host Computer"
    ElseIf gPreprocName <> "" Then
            '������ �������� ���������
        strMessage = "ExitApp"
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call SendMessage(strMessage)
    End If
            
            '���������������� ������� - "��������� �������"
    gProtocol.strProtocName = "################"
            '��������� ������
    gProtocol.strProtocPersonCode = txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Menu EXIT"
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
            
            '��� �������� ������� ������������ ��� ��������� � �����������
    If gMSBase = 1 Then
            '������������ ��� ��������� � ����������� � ������� ACCESS"
        Call BasesConvert
    End If
            
            '������� ���� "������� ���������"
    Close gProtocFileNum
    
            ' ���������� ������ �� ������ "FlexGrid" ("������� ������")
    Set gTablePerson = Nothing
            ' ���������� ������ �� ������ ActiveX.EXE
    Set objTablePerson = Nothing
            
            ' ���������� ������ �� ������ MSMQQueueInfo
    Set qInfoOutput = Nothing
    Set qInfoInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ����������� ���������
    Set qQueueInput = Nothing
            ' ���������� ������ �� ������ �������-�������
            ' ����������� ���������
    Set evQueue = Nothing
            ' ���������� ������ �� ������ ����������� ��������E
    Set qMsgInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ������������ ���������
    Set qQueueOutput = Nothing
            ' ���������� ������ �� ������ ������������ ��������E
    Set qMsgOutput = Nothing
    
            '��������� ���������
    End
    
End Sub

            '��������� ������ ������� "Save Protocol" ���� "Adjustment"
Private Sub mnuSaveProtocol_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '������� ����, ����� �������� ��� ���������� "������� ���������"
    Close gProtocFileNum
            '���������� ������� ���������� ��������� ������� � ������������ �����
    If mnuSaveProtocol.Checked = True Then mnuSaveProtocolAs.Checked = False
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� ����� "������� ���������"
    gProtocFileNum = FreeFile
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '����� ������ ��������� ������ "������� ���������"
    gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
            '���� "������� ���������" ������� �������������
    If gProtocRowNum > 32760 Then
            '������ �������� ������
        BeepSound
        If optEnglish = True Then
            strResponse = MsgBox("The protocol overflow ?", vbYesNo + vbQuestion, "Cancel")
        Else
            strResponse = MsgBox("Protokols ir izpildits ?", vbYesNo + vbQuestion, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '����� ��������� ������ "������� ���������" (������ �������)
            gProtocRowNum = 32760
            '������ ������ "��"
        Else
            '������������� ����� "������� ���������"
            WriteProtocolToArchives
        End If
    End If

End Sub

            '��������� ������ ������� "SaveProtocolAs..." ���� "Adjustment"
Private Sub mnuSaveProtocolAs_Click()
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long

             '��������� (�� ���������) ����� "frmGetFile"
    Load frmGetFile
            '��������� ������ ���������������� ���� "cboFileType
    frmGetFile.cboFileType.AddItem "All files (*.*)"
    frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
    frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            '������� ������� ������ "��� �����"
    frmGetFile.cboFileType.ListIndex = 0
            '������� �� ����� ����� "frmGetFile" � ������� ����������� 1
    frmGetFile.Show 1
            '���� �� ������
    If frmGetFile.Tag = "" Then
            '������ �������� ������
        BeepSound
        MsgBox "The file isn't selected !"
            '������ ���� ��� ���������� "������� ���������"
    Else
    
            '������� ����, ����� �������� ��� ���������� "������� ���������"
        Close gProtocFileNum
            '���������� ������� ���������� ��������� ������� � ���������� �����
        If mnuSaveProtocolAs.Checked = True Then mnuSaveProtocol.Checked = False
    
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
           '��������� ����� ������ (������) "������� ���������"
        lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� ����� "������� ���������"
        gProtocFileNum = FreeFile
    
            '������� ���������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '����� ������ ��������� ������ "������� ���������"
        gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
            '���� "������� ���������" ������� �������������
        If gProtocRowNum > 32760 Then
            '������ �������� ������
            BeepSound
            If optEnglish = True Then
                MsgBox "The protocol overflow !", vbExclamation, "Error"
            Else
                MsgBox "Protokols ir izpildits !", vbExclamation, "Error"
            End If
            '����� ��������� ������ "������� ���������" (������ �������)
            gProtocRowNum = gProtocRowNum - 1
        End If
    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing

End Sub

            '��������� ������ ������� "System" ���� "Adjustment"
Private Sub mnuSystem_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmTableSystem"
    frmTableSystem.Visible = True
            '���������� ����� �� ������ "Correction"
    frmTableSystem.cmdCorrection.SetFocus

End Sub
            
            '��������� ������ ������� "Persons" ���� "Adjustment"
Private Sub mnuPersons_Click()
            
            '������� ������������� ������ "������� ������" �� ����������:
            '   ������������ ���������� "������� ������" "Host Computer'�"
            '   - ��������� ������� �� ��������
    If gCompresTablPers = 0 Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The function don't accessable !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmTablePerson"
    frmTablePerson.Visible = True
            '���������� ����� �� ������ "Correction"
    frmTablePerson.cmdCorrection.SetFocus

End Sub
            
            '��������� ������ ������� "ParkingInData" ���� "Parking"
Private Sub mnuParkingInData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            '����� �������� (����)����������� ������������� ����
    frmDataParkingIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingIn"
    Do While frmDataParkingIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataParkingIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgParkingInData"
Private Sub imgParkingInData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
    
            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuParkingInData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            '���������� ������� ��������� ����������� ������� �����������
    imgParkingInData(intIndex).Tag = 1
            '����� �������� ����������� ����� "frmDataParkingIn"
    frmDataParkingIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingIn"
    Do While frmDataParkingIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� (����)����������� ������������� ����
    frmDataParkingIn.Tag = 0
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataParkingIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)����������� ��������
    If frmDataParkingIn.Tag = 2 Then
            '�������� ������� ��������� �����������
            '  �������� �����������
        imgParkingInData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������,
            '  ��������� ����������� ������� �����������
            '  � ����������� ����� "����������/���������� ��������"
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 And _
    gParkingDeletion = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gParkingPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
            '�� ����������� ����� "����������/���������� ��������"
    Else
            '�������� ������� ��������� ����������� �������
        imgParkingInData(intIndex).Tag = 0
    End If

End Sub

            '������������� ����� ������������� ���� � ������ ������
            '  ��� ��������������� (����� ���� "Controller") ������� �� �����������
Public Function AutoParkReg(ByVal vntPersonCode As Variant, intIndex As Integer)
            
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataParkingIn - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '������� ������� ����� "frmDataParkingIn"
    frmDataParkingIn.Visible = True
            '���������� ������� ��������� ����������� ������� �����������
    imgParkingInData(intIndex).Tag = 1
            '����� �������� (����)����������� ������������� ����
    frmDataParkingIn.Tag = 0
            '��������� ������� "PersonCode" ��� ��������������� �������
            '  ����������� ����� ����������� "Controller"
    intAnalysisCode = frmDataParkingIn.Analysis(vntPersonCode)

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)����������� ������������� ����
    If frmDataParkingIn.Tag = 1 Then frmDataParkingIn.Tag = 0
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataParkingIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)����������� ��������
    If frmDataParkingIn.Tag = 2 Then
            '�������� ������� ��������� �����������
            '   �������� �����������
        imgParkingInData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1
            
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
            '  � ��������� ����������� ������� �����������
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gParkingPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
    End If

End Function
            '��������� ������ ������� "ParkingOutData" ���� "Parking"
Private Sub mnuParkingOutData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            '����� �������� ����������� ����� "frmDataParkingOut"
    frmDataParkingOut.Tag = 0
            '����� �������� (����)����������� ������������� ����
    frmDataParkingIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingOut"
    Do While frmDataParkingOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� (����)�������� ������������� ����
    frmDataParkingOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataParkingOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgParkingOutData"
Private Sub imgParkingOutData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuParkingOutData_Click
        Exit Sub
    End If

    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            '���������� ������� ��������� ���������� ������� �����������
    imgParkingOutData(intIndex).Tag = 1
            '����� �������� ����������� ����� "frmDataParkingOut"
    frmDataParkingOut.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingOut"
    Do While frmDataParkingOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� (����)�������� ������������� ����
    frmDataParkingOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataParkingOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)�������� ��������
    If frmDataParkingOut.Tag = 2 Then
            '�������� ������� ���������
            '  �������� �������� �����������
        imgParkingOutData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������,
            '  ��������� ���������� ������� �����������
            '  � ����������� ����� "����������/���������� ��������"
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingOut.Tag = 1 And _
    gParkingDeletion = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gParkingPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
            '�� ����������� ����� "����������/���������� ��������"
    Else
            '�������� ������� ��������� �������� �������
        imgParkingOutData(intIndex).Tag = 0
    End If

End Sub
            
            '������������� ����� ������������� ���� � ������ ������
            '  ��� ������������ (����� ���� "Controller") ������� �� �����������
Public Function AutoParkDel(ByVal vntPersonCode As Variant, intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataParkingOut - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '����� �������� (����)����������� ������������� ���� (������� ???)
    frmDataParkingIn.Tag = 0
            ' ������� ������� ����� "frmDataParkingOut"
    frmDataParkingOut.Visible = True
            '���������� ������� ��������� ���������� �������
    imgParkingOutData(intIndex).Tag = 1
            '����� �������� (����)�������� ������������� ����
    frmDataParkingOut.Tag = 0
            '��������� ������� "PersonCode" ��� ������������ �������
            '  ����� ����������� "Controller"
    intAnalysisCode = frmDataParkingOut.Analysis(vntPersonCode)
            
            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)�������� ������������� ����
    If frmDataParkingOut.Tag = 1 Then frmDataParkingOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataParkingOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
             '� (����)�������� ��������
    If frmDataParkingIn.Tag = 2 Then
            '�������� ������� ���������
            '  �������� �������� �����������
        imgParkingOutData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
            '  � ��������� ���������� ������� �����������
    If cmdOpen(intIndex).Tag = 0 And frmDataParkingOut.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
        imgParkingInData(intIndex).Enabled = False
        imgParkingOutData(intIndex).Enabled = False
        imgParkingInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gParkingPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
           gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
    End If
            
End Function
            
            '��������� ������ ������� "ParkingInfoData" ���� "Parking"
Private Sub mnuParkingInfoData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingInfo"
    frmDataParkingInfo.Visible = True
            '����� �������� ����������� ����� "frmDataParkingInfo"
    frmDataParkingInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingInfo"
    Do While frmDataParkingInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� ������ ������ � ������� ������
    frmDataParkingInfo.Tag = 0
             '����� ���������� (��� ������ ��) ������
    Do While frmDataParkingInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgParkingInfoData"
Private Sub imgParkingInfoData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuParkingInfoData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingInfo"
    frmDataParkingInfo.Visible = True
            '���������� ������� ��������� ��������� ���������� �� �����������
    imgParkingInfoData(intIndex).Tag = 1
            '����� �������� ����������� ����� "frmDataParkingInfo"
    frmDataParkingInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingInfo"
    Do While frmDataParkingInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� ������ ������ � "������� ������"
    frmDataParkingInfo.Tag = 0
             '����� ���������� (��� ������ ��) ������ ������
    Do While frmDataParkingInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            '�������� ������� ��������� ��������� ���������� �� �����������
    imgParkingInfoData(intIndex).Tag = 0

End Sub
            
            '��������� ������ ������� "ParkingServData" ���� "Parking"
Private Sub mnuParkingServData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataParkingServ"
    frmDataParkingServ.Visible = True
            '����� �������� ����������� ����� "frmDataParkingServ"
    frmDataParkingServ.Tag = 0
            '����� �������� (����)����������� ������������� ����
    frmDataParkingIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataParkingServ"
    Do While frmDataParkingServ.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� (����) ��������� ��� �������
            '   ������������� ����
    frmDataParkingServ.Tag = 0
             '����� ���������� (��� ������ ��) ���������
    Do While frmDataParkingServ.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� ������ ������� "AccessInData" ���� "Access"
Private Sub mnuAccessInData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            '����� �������� (����)����������� ������������� ����
    frmDataAccessIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataAccessIn"
    Do While frmDataAccessIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataAccessIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgAccessInData"
Private Sub imgAccessInData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuAccessInData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '����� �������� ����������� ����� "frmDataAccessIn"
    frmDataAccessIn.Tag = 0
            ' ������� ������� ����� "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            '���������� ������� ��������� ����������� �������
    imgAccessInData(intIndex).Tag = 1
             '����� ���������� ����������� ����� "frmDataAccessIn"
    Do While frmDataAccessIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            
            '� (����)����������� ��������
    If frmDataAccessIn.Tag = 2 Then
            '�������� ������� ��������� ����������� ��������
        imgAccessInData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������,
            '  ��������� ����������� �������
            '  � ����������� ����� "����������/���������� ��������"
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessIn.Tag = 1 And _
    gAccessDeletion = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gAccessPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
            '�� ����������� ����� "����������/���������� ��������"
    Else
            '�������� ������� ��������� ����������� �������
        imgAccessInData(intIndex).Tag = 0
    End If

End Sub

            '������������� ����� ������������� ���� � ������ ������
            '  ��� ��������������� (����� ���� "Controller") �������
Public Function AutoAcceReg(ByVal vntPersonCode As Variant, intIndex As Integer)
            
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataAccessIn - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '������� ������� ����� "frmDataAccessIn"
    frmDataAccessIn.Visible = True
            '���������� ������� ��������� ����������� �������
    imgAccessInData(intIndex).Tag = 1
            '����� �������� (����)����������� ������������� ����
    frmDataAccessIn.Tag = 0
            '��������� ������� "PersonCode" ��� ��������������� �������
            '  ����� ����������� "Controller"
    intAnalysisCode = frmDataAccessIn.Analysis(vntPersonCode)

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)����������� ������������� ����
    If frmDataAccessIn.Tag = 1 Then frmDataAccessIn.Tag = 0
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataAccessIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)����������� ��������
    If frmDataAccessIn.Tag = 2 Then
            '�������� ������� ��������� ����������� ��������
        imgAccessInData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1
            
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
            '  � ��������� ����������� �������
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessIn.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ���������
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gAccessPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
    End If

End Function
            
            '��������� ������ ������� "AccessOutData" ���� "Access"
Private Sub mnuAccessOutData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            '����� �������� ����������� ����� "frmDataAccessOut"
    frmDataAccessOut.Tag = 0
            '����� �������� (����)����������� ������������� ����
    frmDataAccessIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataAccessOut"
    Do While frmDataAccessOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� (����)�������� ������������� ����
    frmDataAccessOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataAccessOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgAccessOutData"
Private Sub imgAccessOutData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuAccessOutData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '����� �������� ����������� ����� "frmDataAccessOut"
    frmDataAccessOut.Tag = 0
            ' ������� ������� ����� "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            '���������� ������� ��������� ���������� �������
    imgAccessOutData(intIndex).Tag = 1
             '����� ���������� ����������� ����� "frmDataAccessOut"
    Do While frmDataAccessOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� (����)�������� ������������� ����
    frmDataAccessOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataAccessOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)�������� ��������
    If frmDataAccessOut.Tag = 2 Then
            '�������� ������� ��������� �������� ��������
        imgAccessOutData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1


            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������,
            '  ��������� ���������� �������
            '  � ����������� ����� "����������/���������� ��������"
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessOut.Tag = 1 And _
    gAccessDeletion = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ���������
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gAccessPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
            '�� ����������� ����� "����������/���������� ��������"
    Else
            '�������� ������� ��������� �������� �������
        imgAccessOutData(intIndex).Tag = 0
    End If

End Sub

            '������������� ����� ������������� ���� � ������ ������
            '  ��� ������������ (����� ���� "Controller") �������
Public Function AutoAcceDel(ByVal vntPersonCode As Variant, intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataAccessOut - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessOut"
    frmDataAccessOut.Visible = True
            '���������� ������� ��������� ���������� �������
    imgAccessOutData(intIndex).Tag = 1
            '����� �������� (����)�������� ������������� ����
    frmDataAccessOut.Tag = 0
            '��������� ������� "PersonCode" ��� ������������ �������
            '  ����� ����������� "Controller"
    intAnalysisCode = frmDataAccessOut.Analysis(vntPersonCode)

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)�������� ������������� ����
    If frmDataAccessOut.Tag = 1 Then frmDataAccessOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataAccessOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
             '� (����)�������� ��������
    If frmDataAccessIn.Tag = 2 Then
            '�������� ������� ��������� �������� ��������
        imgAccessOutData(intIndex).Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1


            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
            '  � ��������� ���������� �������
    If cmdOpen(intIndex).Tag = 0 And frmDataAccessOut.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ���������
        imgAccessInData(intIndex).Enabled = False
        imgAccessOutData(intIndex).Enabled = False
        imgAccessInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
        vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
        cmdOpen(intIndex).Tag = vntAddr
        cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '�� ��������� ������ ������������� �������� ���������
        If gAccessPresButton = 0 Then
            '����� "N_?" - (������� ���)
            lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
            Call cmdOpen_Click(intIndex)
            '��������� ������ ������������� �������� ���������
        Else
            '������� ����������� "������" �������� ���������
            cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
           gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
            lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
            tmrButton(intIndex).Enabled = True
        End If
    End If
            
End Function
            
            '��������� ������ ������� "AccessInfoData" ���� "Access"
Private Sub mnuAccessInfoData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessInfo"
    frmDataAccessInfo.Visible = True
            '����� �������� ����������� ����� "frmDataAccessInfo"
    frmDataAccessInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataAccessInfo"
    Do While frmDataAccessInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� ������ ������ � ������� ������
    frmDataAccessInfo.Tag = 0
             '����� ���������� (��� ������ ��) ������
    Do While frmDataAccessInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgAccessInfoData"
Private Sub imgAccessInfoData_Click(intIndex As Integer)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuAccessInfoData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessInfo"
    frmDataAccessInfo.Visible = True
            '���������� ������� ��������� ��������� ����������
    imgAccessInfoData(intIndex).Tag = 1
            '����� �������� ����������� ����� "frmDataAccessInfo"
    frmDataAccessInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataAccessInfo"
    Do While frmDataAccessInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� ������ ������ � "������� ������"
    frmDataAccessInfo.Tag = 0
             '����� ���������� (��� ������ ��) ������ ������
    Do While frmDataAccessInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            '�������� ������� ��������� ��������� ����������
    imgAccessInfoData(intIndex).Tag = 0

End Sub
            
            '��������� ������ ������� "AccessServData" ���� "Access"
Private Sub mnuAccessServData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataAccessServ"
    frmDataAccessServ.Visible = True
            '����� �������� ����������� ����� "frmDataAccessServ"
    frmDataAccessServ.Tag = 0
            '����� �������� (����)����������� ������������� ����
    frmDataAccessIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataAccessServ"
    Do While frmDataAccessServ.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� (����) ��������� ��� �������
            '   ������������� ����
    frmDataAccessServ.Tag = 0
             '����� ���������� (��� ������ ��) ���������
    Do While frmDataAccessServ.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� ������ ������� "EmployeInData" ���� "Employe"
Private Sub mnuEmployeInData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            '����� �������� (����)����������� ������������� ����
    frmDataEmployeIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgEmployeInData"
Private Sub imgEmployeInData_Click()
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            
            '��������� ������� �������� ������ ��� "Host Computer'a"
    If gPreprocName <> "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuEmployeInData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            '���������� ������� ��������� ����������� ���������
    imgEmployeInData.Tag = 1
            '����� �������� ����������� ����� "frmDataEmployeIn"
    frmDataEmployeIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� ����������� ����� "frmDataEmployeIn"
    If frmDataEmployeIn.Tag = 1 Then frmDataEmployeIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeIn"
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)����������� ��������
    If frmDataEmployeIn.Tag = 2 Then
            '�������� ������� ��������� ����������� ���������
        imgEmployeInData.Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

End Sub

            '������������� ����� ������������� ���� � ������ ������
            '  ��� ��������������� (����� ���� "Controller") ���������
Public Function AutoEmplReg(ByVal vntPersonCode As Variant)
            
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataEmployeIn - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
          '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            '������� ������� ����� "frmDataEmployeIn"
    frmDataEmployeIn.Visible = True
            '���������� ������� ��������� ����������� ���������
    imgEmployeInData.Tag = 1
            '����� �������� (����)����������� ������������� ����
    frmDataEmployeIn.Tag = 0
            '��������� ������� "PersonCode" ��� ��������������� ���������
            '  ����� ����������� "Controller"
    intAnalysisCode = frmDataEmployeIn.Analysis(vntPersonCode)

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)����������� ������������� ����
    If frmDataEmployeIn.Tag = 1 Then frmDataEmployeIn.Tag = 0
             '����� ���������� (��� ������ ��) �����������
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)����������� ��������
    If frmDataEmployeIn.Tag = 2 Then
            '�������� ������� ��������� ����������� ���������
        imgEmployeInData.Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

End Function
            
            '��������� ������ ������� "EmployeOutData" ���� "Employe"
Private Sub mnuEmployeOutData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            '����� �������� ����������� ����� "frmDataEmployeOut"
    frmDataEmployeOut.Tag = 0
            '����� �������� (����)����������� ������������� ����
    frmDataEmployeIn.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������) � ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������) - ����� �� ���������
    If gTimeShare = 1 And chkSetup.Value = 1 Then Exit Sub
            '����� �������� (����)�������� ������������� ����
    frmDataEmployeOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataEmployeOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgEmployeOutData"
Private Sub imgEmployeOutData_Click()
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            
            '��������� ������� �������� ������ ��� "Host Computer'a"
    If gPreprocName <> "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The function accessable only to HostComputer !", _
        vbExclamation, "Error"
        Exit Sub
    End If
            
            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuEmployeOutData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            '���������� ������� ��������� ���������� ���������
    imgEmployeOutData.Tag = 1
            '����� �������� ����������� ����� "frmDataEmployeOut"
    frmDataEmployeOut.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� ����������� ����� "frmDataEmployeOut"
    If frmDataEmployeOut.Tag = 1 Then frmDataEmployeOut.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeOut"
    Do While frmDataEmployeOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '� (����)�������� ��������
    If frmDataEmployeOut.Tag = 2 Then
            '�������� ������� ��������� �������� ��������
        imgEmployeOutData.Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

End Sub

            '������������� ����� ������������� ���� � ������ ������
            '  ��� ������������ (����� ���� "Controller") ���������
Public Function AutoEmplDel(ByVal vntPersonCode As Variant)
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� �������  "PersonCode"
Dim intAnalysisCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataEmployeOut - ����� �� �������
    If frmDemo.Enabled = False Then Exit Function
    
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeOut"
    frmDataEmployeOut.Visible = True
            '���������� ������� ��������� ���������� ���������
    imgEmployeOutData.Tag = 1
            '����� �������� (����)�������� ������������� ����
    frmDataEmployeOut.Tag = 0
            '��������� ������� "PersonCode" ��� ������������ ���������
            '  ����� ����������� "Controller"
    intAnalysisCode = frmDataEmployeOut.Analysis(vntPersonCode)

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then Exit Function

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '����� �������� (����)�������� ������������� ����
    If frmDataEmployeOut.Tag = 1 Then frmDataEmployeOut.Tag = 0
             '����� ���������� (��� ������ ��) ��������
    Do While frmDataEmployeOut.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
             '� (����)�������� ��������
    If frmDataEmployeIn.Tag = 2 Then
            '�������� ������� ��������� �������� ���������
        imgEmployeOutData.Tag = 0
    End If
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1
            
End Function
            
            '��������� ������ ������� "EmployeInfoData" ���� "Employe"
Private Sub mnuEmployeInfoData_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeInfo"
    frmDataEmployeInfo.Visible = True
            '����� �������� ����������� ����� "frmDataEmployeInfo"
    frmDataEmployeInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop

End Sub
            
            '��������� "������" ���� �� �������� "imgEmployeInfoData"
Private Sub imgEmployeInfoData_Click()
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer

            '����������� ����� ���������� �������
            '   (������������ ���������� ��������)
    If gTimeShare = 1 Then
        Call mnuEmployeInfoData_Click
        Exit Sub
    End If

            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmDataEmployeInfo"
    frmDataEmployeInfo.Visible = True
            '���������� ������� ��������� ��������� ����������
    imgEmployeInfoData.Tag = 1
            '����� �������� ����������� ����� "frmDataEmployeInfo"
    frmDataEmployeInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '����� �������� ����������� ����� "frmDataEmployeInfo"
    If frmDataEmployeInfo.Tag = 1 Then frmDataEmployeInfo.Tag = 0
             '����� ���������� ����������� ����� "frmDataEmployeInfo"
    Do While frmDataEmployeInfo.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1

            '�������� ������� ��������� ��������� ����������
    imgEmployeInfoData.Tag = 0

End Sub
            
            '������������ ��� ��������� � ����������� � ������� ACCESS"
Public Sub BasesConvert()
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer


            '���������� ������ ���� "�������� ����" ��� ������� ������
    frmDemo.MousePointer = vbHourglass

            ' ��������� �����/�������� ���������� ��� ����������
    If prtPortC(0).PortOpen = True Then prtPortC(0).PortOpen = False
    If prtPortC(1).PortOpen = True Then prtPortC(1).PortOpen = False
    If prtPortC(2).PortOpen = True Then prtPortC(2).PortOpen = False
    If prtPortC(3).PortOpen = True Then prtPortC(3).PortOpen = False
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 0
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
    

            '���� ��� "Host Computer" � ���� ������������� � ��������� ����
    If gPreprocName = "" And gNetPreprocNum > 0 Then
            '������������ ��� ��������� � ����������� � ������� ACCESS"
            '  ��� ���� �������������� � "Host Computer'a" - ����� ����
        Call frmPreprocessors.BasesConvert
            
            '���� ��� ������������ ��� ��� �������������� � ��������� ����
    ElseIf gPreprocName <> "" Or gNetPreprocNum = 0 Then
            '����� ������� "mnuFormProtocolBase" ���� "File"
        mnuFormProtocolBase_Click
            '����� ������� "mnuFormBookKeeperBase" ���� "File"
        mnuFormBookKeeperBase_Click
            
    End If
    
            
            ' ��������� �����/�������� ���������� ��� ���������� ����������
            ' ���������� ����� "��������" - ����� ������������ ������
            ' ���������� ����������
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
            '��������� ����� "Controller'��" �� ��������
    gTermContr = 1
        
            '������������ ����������� ������ ���� ��� ������� ������
    frmDemo.MousePointer = 0
            
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True

End Sub
            
            '��������� ������ ������� "Calendar" ���� "Adjustment"
Private Sub mnuCalendar_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmTableCalendar"
    frmTableCalendar.Visible = True
            '���������� ����� �� ������ "Correction"
    frmTableCalendar.cmdCorrection.SetFocus
    

End Sub
            
            '��������� ������ ������� "Timer" ���� "Adjustment"
Private Sub mnuTime_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmTableTime"
    frmTableTime.Visible = True
            '���������� ����� �� ������ "Correction"
    frmTableTime.cmdCorrection.SetFocus
    

End Sub
            
            '��������� ������ ������� "Terminal" ���� "Adjustment"
Private Sub mnuTerminal_Click()
            '������� ����������� ����� "frmDemo"
    frmDemo.Enabled = False
            ' ������� ������� ����� "frmTableTerminal"
    frmTableTerminal.Visible = True
            '���������� ����� �� ������ "Correction"
    frmTableTerminal.cmdCorrection.SetFocus
    

End Sub

            ' ������� ����� "English"
Private Sub optEnglish_Click()
            ' ������� ����� ����� ������� �� "English"
    If fraFlag.Tag <> 0 Then
            ' ���������� ����� ������ ����� �������
    fraFlag.Tag = 0
            ' ����� ��������� ��������� ����� �������
    UpdateLanguage
    End If
    
End Sub
            ' ������� ����� "Latvian"
Private Sub optLatvian_Click()
            ' ������� ������ ����� ������� �� "Latvian"
    If fraFlag.Tag <> 1 Then
            ' ���������� ����� ������ ����� �������
    fraFlag.Tag = 1
            ' ����� ��������� ��������� ����� �������
    UpdateLanguage
    End If
    
End Sub

            ' ������� ����� "Russian"
Private Sub optRussian_Click()
            ' ������� ������ ����� ������� �� "Russian"
    If fraFlag.Tag <> 2 Then
            ' ���������� ����� ������ ����� �������
    fraFlag.Tag = 2
            ' ����� ��������� ��������� ����� �������
    UpdateLanguage
    End If
    
End Sub
            ' ��������� ��������� ����� �������
Public Sub UpdateLanguage()
            '����� ������ ����� �������
    intLang = fraFlag.Tag
            '���������� ����
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
    
                ' ���������������� ������ ��������
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
            ' ���������������� ������ "�����������" ���������
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
            
            '��������� ������ ����������� ��������� �������
Public Sub BeepSound()
            '������� �������
Dim intCount As Integer
            '�������� ������ ��������� �������
    For intCount = 1 To gBeepSound
        Beep
    Next
    
End Sub

            '��������� ������� - ��������� �������� �
            '  ������� ����������� ���������
Private Sub qEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '��� �������� ��� ������ ������� "Shell"
Dim vntShell As Variant
            '����� ������ "��������� �������", "������� ���������" ��� "������� ���������"
Dim lngRecordLen As Long
            'C�������� ����� �����
Dim intFileNum As Integer
            'C�������� �����
Dim strTime As String
Dim vntTime As Variant
            'C�������� ����
Dim strDate As String
Dim vntDate As Variant

            '����������-������ �������-�������
            ' ����������� ���������
    Set evQueue = qQueueInput
            '������� ���������
    Set qMsgInput = evQueue.Receive(, , , 0)
            
            '���� "Host Computer" ������ ������ � ����������
            '  ��������� ���� �� �����������
    If Mid(qMsgInput.Body, 4, 15) = "ParkFreePlaces " And _
    gPreprocName = "" Then
            ' ���� ������� �������-��������� ���������� ��������� ����
        If gParkingPlaceNum <> 0 Then
            '������������ ����������� ���������
            qMsgOutput.Body = "ParkFreePlaces=" + CStr(gParkFreePlaces)
            '���������� ���� � ������� ������������ ���������
            '   ����������� �������������, ����������� ���������
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            qMsgInput.Label + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
            qMsgOutput.Send qQueueOutput
            '������� ������� ���������
            qQueueOutput.Close
        End If
            
            '���� "Host Computer" ������ ������ � ����������
            '  ��������� ���� �� �����������
    ElseIf Left(qMsgInput.Body, 15) = "AcceFreePlaces " And _
    gPreprocName = "" Then
            ' ���� ������� �������-��������� ���������� ��������� ����
        If gAccessPlaceNum <> 0 Then
            '������������ ����������� ���������
            qMsgOutput.Body = "AcceFreePlaces=" + CStr(gParkFreePlaces)
            '���������� ���� � ������� ������������ ���������
            '   ����������� �������������, ����������� ���������
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            qMsgInput.Label + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
            qMsgOutput.Send qQueueOutput
            '������� ������� ���������
            qQueueOutput.Close
        End If
            
            '���� ������� ��������� � ������������� ������������� ����������
            '  ��������� ���� � ��� �� ��� ������ �� �������������
    ElseIf Mid(qMsgInput.Body, 5, 10) = "FreePlaces" And _
    Mid(qMsgInput.Body, 15, 1) <> " " Then
            '����� ������� ����������� ���������� �� �������
        Call Display(qMsgInput.Body)
            
            '���� "Host Computer" ������ ��������� ��
            '  ������������� ��������� ����� �� ��������������
    ElseIf Left(qMsgInput.Body, 7) = "Archive" And gPreprocName = "" Then
            '����������� ������ �� ������������� � "Host Computer"
'        Call ArchiveCopy(qMsgInput.Body)
            
            '���� "Host Computer" ������ ��������� � �������������
            '  ������������� ������� �� ������ �� ��������������
    ElseIf Left(qMsgInput.Body, 4) = "Time" And gPreprocName = "" Then
        vntTime = Time
        strTime = CStr(vntTime)
        vntDate = Date
        strDate = CStr(vntDate)
            ' ������������ ����������� ���������
        qMsgOutput.Body = "Time " + Trim(strTime) + "||" + Trim(strDate)
            '���������� ���� � ������� ������������ ���������
            '   ����������� �������������, ����������� ���������
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        qMsgInput.Label + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
        qMsgOutput.Send qQueueOutput
            '������� ������� ���������
        qQueueOutput.Close
        
            '��������� ������� �� "Host Computer'a" ��
            '  ������������� �������
    ElseIf Left(qMsgInput.Body, 4) = "Time" And gPreprocName <> "" Then
        strTime = Mid(qMsgInput.Body, 6, 8)
        vntTime = strTime
        Time = Format(vntTime, "hh:mm:ss")
        strDate = Mid(qMsgInput.Body, 16)
        vntDate = strDate
        Date = Format(vntDate, "dd.MM.yyyy.")
            '���������� ����������� "~" � �������� "������" �������� ����� �������
            '  ������ � ���� "������� ���������"
        qMsgInput.Body = "~" + Mid(qMsgInput.Body, 6)
        
            '��������� ������� �� "Host Computer'a"
            '  �� ������� ����������
    ElseIf Left(qMsgInput.Body, 8) = "StartApp" Then
            
            '������ �������� ���������
        strMessage = "ExitApp"
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call SendMessage(strMessage)
            
        gProtocol.strProtocName = "FROM:"
            '��������� ������
        gProtocol.strProtocPersonCode = qMsgInput.Label
            '������
        gProtocol.strProtocStatus = "?? - MSMQ"
            '�����
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = qMsgInput.Body
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            '������� ���� "������� ���������"
        Close gProtocFileNum
            
            ' ���������� ������ �� ������ "FlexGrid" ("������� ������")
        Set gTablePerson = Nothing
            ' ���������� ������ �� ������ ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' ���������� ������ �� ������ MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ����������� ���������
        Set qQueueInput = Nothing
            ' ���������� ������ �� ������ �������-�������
            ' ����������� ���������
        Set evQueue = Nothing
            ' ���������� ������ �� ������ ����������� ��������E
        Set qMsgInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ������������ ���������
        Set qQueueOutput = Nothing
            ' ���������� ������ �� ������ ������������ ��������E
        Set qMsgOutput = Nothing
        
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            '���������� "����" � ������������� EXE-������
        strPathFileName = strPathFileName + gModuleStartUp
            '�������� "Clipboard"
        Clipboard.Clear
            '������� � "Clipboard" ������ ��� ���������-��������
        Clipboard.SetText strPathFileName
            
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            '��������� ���������-������� � ��������� ������
        strPathFileName = strPathFileName + "StartModule.exe"
        
        vntShell = Shell(strPathFileName, 0)
            '������ ���������-�������� �������� -
            '  ��������� ������� ����������
        End
    
            '��������� ������� �� "Host Computer'a"
            '  �� ���������� ���������� ����������
    ElseIf Left(qMsgInput.Body, 7) = "StopApp" Then
            
            '������ �������� ���������
        strMessage = "ExitApp"
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call SendMessage(strMessage)
            
        gProtocol.strProtocName = "FROM:"
            '��������� ������
        gProtocol.strProtocPersonCode = qMsgInput.Label
            '������
        gProtocol.strProtocStatus = "?? - MSMQ"
            '�����
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = qMsgInput.Body
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            '������� ���� "������� ���������"
        Close gProtocFileNum
            
            ' ���������� ������ �� ������ "FlexGrid" ("������� ������")
        Set gTablePerson = Nothing
            ' ���������� ������ �� ������ ActiveX.EXE
        Set objTablePerson = Nothing
            
            ' ���������� ������ �� ������ MSMQQueueInfo
        Set qInfoOutput = Nothing
        Set qInfoInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ����������� ���������
        Set qQueueInput = Nothing
            ' ���������� ������ �� ������ �������-�������
            ' ����������� ���������
        Set evQueue = Nothing
            ' ���������� ������ �� ������ ����������� ��������E
        Set qMsgInput = Nothing
            ' ���������� ������ �� ������
            '  ������� ������������ ���������
        Set qQueueOutput = Nothing
            ' ���������� ������ �� ������ ������������ ��������E
        Set qMsgOutput = Nothing
        
            '��������� ������� ����������
        End
    
            '��������� ��������� � ������������� ���������� � "������� ������"
            '  ������ � �������� ������
    ElseIf Left(qMsgInput.Body, 4) = "Reg " Then
            '���������� ������ � �������� ������ � "������� ������"
            '  �� ��������� MSMQ, ���������� �� ����
        Call frmTablePerson.MSMQReg(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
               
            '��������� ��������� � ������������� �������� �� "������� ������"
            '  ������ � �������� ������������ �����
    ElseIf Left(qMsgInput.Body, 4) = "Del " Then
            '�������� (����������) ������ � �������� ������������ �����
            '  �� "������� ������" �� ��������� MSMQ, ���������� �� ����
        Call frmTablePerson.MSMQDel(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
               
            '��������� ��������� � ������������� ��������� "������� ������"
            '  ������ � �������� ������
    ElseIf Left(qMsgInput.Body, 4) = "Cor " Then
            '���������� ������ � �������� ������ � "������� ������"
            '  �� ��������� MSMQ, ���������� �� ����
        Call frmTablePerson.MSMQCor(Right(qMsgInput.Body, Len(qMsgInput.Body) - 4))
            
            '��������� ��������� � ������������� ���������� � "������� ����������"
            '  ������ � �������� ������
    ElseIf Left(qMsgInput.Body, 8) = "RegInfo " Then
            '���������� ������ � �������� ������ � "������� ����������"
            '  �� ��������� MSMQ, ���������� �� ����
'        Call frmTableInfo.MSMQReg(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
               
            '��������� ��������� � ������������� �������� �� "������� ����������"
            '  ������ � �������� ������������ �����
    ElseIf Left(qMsgInput.Body, 8) = "DelInfo " Then
            '�������� (����������) ������ � ��������� ������������ ����� � ������
            '  �� "������� ����������" �� ��������� MSMQ, ���������� �� ����
'        Call frmTableInfo.MSMQDel(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
               
            '��������� ��������� � ������������� ��������� � "������� ����������"
            '  ������ � �������� ������
    ElseIf Left(qMsgInput.Body, 8) = "CorInfo " Then
            '���������� ������ � �������� ������ � "������� ����������"
            '  �� ��������� MSMQ, ���������� �� ����
'        Call frmTableInfo.MSMQCor(Right(qMsgInput.Body, Len(qMsgInput.Body) - 8))
            
            '��������� ��������� � ������� �����/������(������/������)
    ElseIf Left(qMsgInput.Body, 9) = "Error Inp" Or _
    Left(qMsgInput.Body, 9) = "Error Out" Then
            '������������ - �� ������������ � �� ���������� ���������
        GoTo IgnoreMSMQ
            
            '�����
    Else
            '������ ����������� ����� � ������ ���������
        strMsgInput = qMsgInput.Label + " || " + qMsgInput.Body
            '������ �������� ������
        BeepSound
            '����������� �������� ����� � ������ ���������
            '  � ����� �������� ���������
        lblMessageInput.Caption = "FROM: " + strMsgInput
            '������� ������� ����� ���������
        lblMessageInput.Visible = True
    End If
            
            '���������������� ������� - ��������� ��������� MSMQ
    gProtocol.strProtocName = "FROM:"
            '��������� ������
    gProtocol.strProtocPersonCode = qMsgInput.Label
            '������
    gProtocol.strProtocStatus = "?? - MSMQ"
            '�����
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = qMsgInput.Body
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
            
IgnoreMSMQ:
            '������������ ?? ����� � �������� MSMQEvent
    qQueueInput.EnableNotification qEvent

End Sub
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
Public Function SendMessage(strMessage As String)
            '������� �������
Dim intCount As Integer
            
            '���� ���������� ��������� �� ��������� ����������
            '  ��������� ����  � ��� �� ������ �� �������������
    If Mid(strMessage, 5, 10) = "FreePlaces" And _
    Mid(strMessage, 15, 1) <> " " Then
            '����� ������� ����������� ���������� �� �������
        Call Display(strMessage)
    End If
            
            '���� ���� ������������� � ��������� ����
    If gNetPreprocNum > 0 And Not (qMsgOutput Is Nothing) Then
            '������������ ����������� ���������
        qMsgOutput.Body = strMessage
            '�� ���� ��������� ������� ���� ����������� ��������� ����
        For intCount = 1 To gNetPreprocNum
            '���������� ���� � ������� ������������ ���������
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            gSocketNet(intCount) + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
            qMsgOutput.Send qQueueOutput
            '������� ������� ���������
            qQueueOutput.Close
        Next
        
    End If
    
End Function
            
            '��������� ������ ���������� �� �������-���������
Public Function Display(strMessage As String)
            '����������-������ ���������� ��� �������
Dim strDisplay As String
            '������� �����
Dim strTime As String
            '����
Dim intHour As Integer
            '������
Dim intMinute As Integer
            
            ' ���� ������� �������-��������� ���������� ���������
            '  ���� �� ����������� ��� �� �����������
    If gParkingPlaceNum <> 0 Or gAccessPlaceNum <> 0 Then
            '������������� ���������� (���������� ��������� ����)
        If gParkingPlaceNum > 0 Then
            strDisplay = CStr(gParkFreePlaces)
        ElseIf gAccessPlaceNum > 0 Then
            strDisplay = CStr(gAcceFreePlaces)
        End If
            '���� ������� ��������� �� ���������� ����������
            '  ��������� ���� �� �����������
        If strMessage = "ParkFreePlaces+1" And _
        gParkFreePlaces < gParkingPlaceNum Then
            '���� ��������������� ������� � ������ �����������
            If Not (Left(gDefaultParkTime, 2) = "00" And _
            Mid(gDefaultParkTime, 4, 2) = "00" And _
            Mid(gDefaultParkTime, 7, 2) = "24" And _
            Mid(gDefaultParkTime, 10, 2) = "00") Then
            '������� �����
                strTime = Format(Now, "h:mm:ss")
            '����
                intHour = Hour(strTime)
            '������
                intMinute = Minute(strTime)
            '������� ����� (�� ��������������� �������)
                If ((Left(gDefaultParkTime, 2) < intHour Or _
                Left(gDefaultParkTime, 2) = intHour And _
                Mid(gDefaultParkTime, 4, 2) <= intMinute) And _
                Mid(gDefaultParkTime, 7, 2) > intHour) Then
            '���������� ���������� ��������� ����
                    gParkFreePlaces = gParkFreePlaces + 1
                    strDisplay = CStr(gParkFreePlaces)
                End If
            '��� ���������������� �������� � ������ �����������
            Else
            '���������� ���������� ��������� ����
                gParkFreePlaces = gParkFreePlaces + 1
                strDisplay = CStr(gParkFreePlaces)
            End If
            '���� ������� ��������� �� ���������� ����������
            '  ��������� ���� �� �����������
        ElseIf strMessage = "ParkFreePlaces-1" And _
        gParkFreePlaces > 0 Then
            '���������� ���������� ��������� ����
            gParkFreePlaces = gParkFreePlaces - 1
            strDisplay = CStr(gParkFreePlaces)
            '���� ������� ��������� � ����������
            '  ��������� ���� �� �����������
        ElseIf Left(strMessage, 15) = "ParkFreePlaces=" Then
            '���������� ���������� ��������� ����
            gParkFreePlaces = Mid(strMessage, 16)
            '���� ������� ��������� �� ���������� ����������
            '  ��������� ���� �� �����������
        ElseIf strMessage = "AcceFreePlaces+1" And _
        gAcceFreePlaces < gAccessPlaceNum Then
            '���� ��������������� ������� � ������ �����������
            If Not (Left(gDefaultAcceTime, 2) = "00" And _
            Mid(gDefaultAcceTime, 4, 2) = "00" And _
            Mid(gDefaultAcceTime, 7, 2) = "24" And _
            Mid(gDefaultAcceTime, 10, 2) = "00") Then
            '������� �����
                strTime = Format(Now, "h:mm:ss")
            '����
                intHour = Hour(strTime)
            '������
                intMinute = Minute(strTime)
            '������� ����� (�� ��������������� �������)
                If ((Left(gDefaultAcceTime, 2) < intHour Or _
                Left(gDefaultAcceTime, 2) = intHour And _
                Mid(gDefaultAcceTime, 4, 2) <= intMinute) And _
                Mid(gDefaultAcceTime, 7, 2) > intHour) Then
            '���������� ���������� ��������� ����
                    gAcceFreePlaces = gAcceFreePlaces + 1
                    strDisplay = CStr(gAcceFreePlaces)
                End If
            '��� ���������������� �������� � ������ �����������
            Else
            '���������� ���������� ��������� ����
                gAcceFreePlaces = gAcceFreePlaces + 1
                strDisplay = CStr(gAcceFreePlaces)
            End If
            '���� ������� ��������� �� ���������� ����������
            '  ��������� ���� �� �����������
        ElseIf strMessage = "AcceFreePlaces-1" And _
        gAcceFreePlaces > 0 Then
            '���������� ���������� ��������� ����
            gAcceFreePlaces = gAcceFreePlaces - 1
            strDisplay = CStr(gAcceFreePlaces)
            '���� ������� ��������� � ����������
            '  ��������� ���� �� �����������
        ElseIf Left(strMessage, 15) = "AcceFreePlaces=" Then
            '���������� ���������� ��������� ����
            gAcceFreePlaces = Mid(strMessage, 16)
            '���� ������� ���������� ��������� �� ���������
            '  ���������a ��������� ���� ��� �������� ���������� �������
        Else
            '������������ ���������� ��������� ���� � ���������� ��������
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
            '�������������� ���������� ��� ���������� ���������
            Else
            '������������� ������� - ���� ������
                strDisplay = Chr(CLng(CByte(161))) + _
                Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
                Chr(CLng(CByte(0)))
            '������� ������ �� �������
                prtPortDisplay.Output = strDisplay
            '����� ���������� �������� ������ �� �������
                Do
                Loop Until prtPortDisplay.OutBufferCount = 0
                Exit Function
            End If
        End If
        
            '���������� ������������� �� ������� ���������� ��������� ����
        If strDisplay - gDisplayDiscount >= 0 Then _
        strDisplay = CStr(strDisplay - gDisplayDiscount)
        
            '���������� ������ ���������� �� �������
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
        
            '������� ������ �� �������
        prtPortDisplay.Output = strDisplay
            '����� ���������� �������� ������ �� �������
        Do
        Loop Until prtPortDisplay.OutBufferCount = 0
        
    End If
    
End Function
            
            '��������� ������� TimeOut ��� "Relay"
Private Sub tmrRelay_Timer()
            
            '��������� �������� �������
    tmrRelay.Enabled = False
            '���������� ������� ������� TimeOut
    tmrRelay.Tag = 1

End Sub

            '��������� ������� TimeOut ��� "Controller'a"
Private Sub tmrTimeOut_Timer(intIndex As Integer)
            
            '���������������� ������� - "TimeOut"
    gProtocol.strProtocName = "$$$"
            '��������� ������
    gProtocol.strProtocPersonCode = ""
            '������
    gProtocol.strProtocStatus = ""
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "COMMAND TimeOut"
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
            
            '��������� �������� �������
    tmrTimeOut(intIndex).Enabled = False
            '���������� ������� ������� TimeOut
    tmrTimeOut(intIndex).Tag = 1
    
End Sub
            
            '������������ ������ ������ ������������� ����
Private Sub tmrTermContr_Timer()
            '������ �������� � ������� ��������� ���������� ����
Static intControlIndex As Integer
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������� �������
Static intCount As Integer
            
            '������ ������ ����������
    If gTermContr = 0 Then Exit Sub
            
            '���� � ������� ������� ������ - �������� � ���������� �����
    Do While prtPortC(intControlIndex).PortOpen = False
            '������������� ����� �������� ��������
            ' (������������ "��������" ��������)
            '  � ������� "������� ����������"
        intCount = 0
           '���������� ����� �������� � �������� ����������� ��������� ����
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
    Loop
    
            '������������� ����� �������� ��������
            ' (������������ "��������" ��������)
            '  � ������� "������� ����������"
    If intCount < 15 Then
        intCount = intCount + 1
    Else
        intCount = 0
           '���������� ����� �������� � �������� ����������� ��������� ����
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
            '����� �� ���������
        Exit Sub
    End If
            
            '���� ������� �� ������������ ���������� - ���������� ���� ������
            '   ����������, ����������� ��� ��� ������ ������������ (Main)
    If prtPortC(intControlIndex).Tag > 0 Then
            '������������ ���������� ������������� �����
            '  �������� ��������  (������������ "��������" ��������)
            '  - ���������� ����������� ���������� ����� � ������ ���������
        intCount = intCount - 1
        Exit Sub
    End If
            
            '����� �������� ��������
            ' � ������� "������� ����������",
    intRequest = (prtPortC(intControlIndex).CommPort - 2) * 15 + intCount
        
            ' "Controller" ��������� ������� � �� ����� �������������
            '   ��������� ���� ������� ���������� ���������� �� �����������
    If (Mid(gAddrPort(0, intRequest), 4) = "0" Or _
    Mid(gAddrPort(0, intRequest), 4) = "#") And _
    Mid(gAddrPort(0, intRequest), 1, 2) <> "00" Then
            '�������� �������� ����� �����
        prtPortC(intControlIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������ ������������� ����
            '  � ����������� �������
        prtPortC(intControlIndex).Output = Chr(CLng(CByte(176) Or CByte(intCount)))
             '����� ���������� �������� ������� ������ ������������� ����
        Do
        Loop Until prtPortC(intControlIndex).OutBufferCount = 0
            ' "Controller" ��������� �������� ��� �����  �������������
            '   ��������� � �� ������� ���������� ���������� �� �����������
    Else
            '������ ��������� "Controller", ������� ��������� �������
            '  � �� ����� ������������� ��������� ���� �������
            '  ���������� ���������� �� �����������
        Do Until (Mid(gAddrPort(0, intRequest), 4) = "0" Or _
        Mid(gAddrPort(0, intRequest), 4) = "#") And _
        Mid(gAddrPort(0, intRequest), 1, 2) <> "00"
        
            '������������� ����� �������� ��������
            ' (������������ "��������" ��������)
            '  � ������� "������� ����������"
            If intCount < 15 Then
                intCount = intCount + 1
            Else
                intCount = 0
           '���������� ����� �������� � �������� ����������� ��������� ����
                If intControlIndex < 3 Then
                    intControlIndex = intControlIndex + 1
                Else
                    intControlIndex = 0
                End If
            '����� �� ���������
                Exit Sub
            End If
            
            '����� �������� ��������
            ' � ������� "������� ����������",
            intRequest = (prtPortC(intControlIndex).CommPort - 2) * 15 + intCount
        
        Loop
            '������������ ���������� ������������� �����
            '  �������� ��������  (������������ "��������" ��������)
            '  - ���������� ����������� ���������� ����� � ������ ���������
        If intCount <> 0 Then intCount = intCount - 1
    
    End If
            
End Sub
            
            '����� ��������� ��������� ������� ����������� "������"
Public Function OpenBarrier(intIndex As Integer)
            
            '����������� ������� ����������� "������"
    Call cmdOpen_Click(intIndex)

End Function

            '��������� ��������� ������� ����������� "������"
Private Sub cmdOpen_Click(intIndex As Integer)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '��� �������� ������� ��������� ������ "Reserve" � "������� ������"
            '  ����� ����������� � ���������� �������
Dim intCode As Integer
            '������� ����������� "������" �����������
    cmdOpen(intIndex).Enabled = False
            ' "������" ������ ����� "Controller'a", ����������
            '  ������� ������������� �������� ��������� ���
            '  �������� "������" ���������� �����������
    If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Or _
    cmdOpen(intIndex).Caption <> chkTerm(intIndex).Caption _
    And optManual = True Then
            '������ ��������� ���������������� "Controller'a"
        vntAddr = cmdOpen(intIndex).Tag
        intRequest = (prtPortC(intIndex).CommPort - 2) * 15 + vntAddr
            ' "Controller" ����� ������������� ��������� - ����� �� ���������
        If Mid(gAddrPort(0, intRequest), 4) <> "0" Then Exit Sub
            ' "Controller" ��������� �������� �� ������� - ����� �� ���������
        If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then Exit Sub
            '��������� "������������������ �������� ���������"
            '  �� ����������� "������"
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "1"
            '��������� ������� "Controller'��", ������� ������������� ����������
        prtPortC(intIndex).Tag = prtPortC(intIndex).Tag + 1
            '�������� ����������� "������" ��� ����������
            '  "��������������" ���������� �����������
        If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Then
            cmdOpen(intIndex).Caption = chkTerm(intIndex).Caption
            cmdOpen(intIndex).Tag = 0
            '��������� ������ ����������� "������"
            tmrButton(intIndex).Enabled = False
            '���������� ������� ��������� ����������� ������� �����������
            If imgParkingInData(intIndex).Tag = 1 Then
            '��������� ������ "Reserve" � "������� ������"
            '  (���������� ������) ����� ����������� ������� �����������
                intCode = frmTablePerson.InputParking(intIndex)
            '���� ������������ �������� ��� ��������� ������
            '  "Reserve" � "������� ������" (����������� ������� �����������)
                If intCode <> 0 Then
            '���� �������� � �������� ��������� "������� ������" - �� �����
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
                    BeepSound
                    MsgBox "Error Parking Registration  !!!", intButtonsAndIcons, "Error"
                End If
            
            '���������� ������� ��������� ����������� ������� �����������
            ElseIf imgAccessInData(intIndex).Tag = 1 Then
            '��������� ������ "Reserve" � "������� ������"
            '  (���������� �����) ����� ����������� ������� �����������
                intCode = frmTablePerson.InputAccess(intIndex)
            '���� ������������ �������� ��� ��������� ������
            '  "Reserve" � "������� ������" (����������� ������� �����������)
                If intCode <> 0 Then
            '���� �������� � �������� ��������� "������� ������" - �� �����
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
                    BeepSound
                    MsgBox "Error Access Registration  !!!", intButtonsAndIcons, "Error"
                End If
            
            '���������� ������� ��������� ���������� ������� �����������
            ElseIf imgParkingOutData(intIndex).Tag = 1 Then
            '��������� ������ "Reserve" � "������� ������"
            '  (���������� ������) ����� ���������� ������� �����������
                intCode = frmTablePerson.OutputParking(intIndex, 6)
            
            '���������� ������� ��������� ���������� ������� �����������
            ElseIf imgAccessOutData(intIndex).Tag = 1 Then
            '��������� ������ "Reserve" � "������� ������"
            '  (���������� �����) ����� ���������� �������
                intCode = frmTablePerson.OutputAccess(intIndex, 9)
            End If
            
            '�������� �������� �������� �����������
            '   � �������� ��������
            imgParkingInData(intIndex).Tag = 0
            imgParkingOutData(intIndex).Tag = 0
            
            imgAccessInData(intIndex).Tag = 0
            imgAccessOutData(intIndex).Tag = 0
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ���������
            imgParkingInData(intIndex).Enabled = True
            imgParkingOutData(intIndex).Enabled = True
            imgParkingInfoData(intIndex).Enabled = True
            
            imgAccessInData(intIndex).Enabled = True
            imgAccessOutData(intIndex).Enabled = True
            imgAccessInfoData(intIndex).Enabled = True
        End If
        
            '���������������� ������� - ������� ����������� "������"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            '��������� ������
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "BUTTON PRESSING"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
        
    End If

End Sub

            '��������� ������� "TimeOut" ����������� "������"
Private Sub tmrButton_Timer(intIndex As Integer)
            '��������� ������ ����������� "������"
    tmrButton(intIndex).Enabled = False
            '���������� ������� "TimeOut" ��� ����������� "������"
    tmrButton(intIndex).Tag = 1
            '�������� ����������� "������" ��� ����������
            '  "��������������" ���������� �����������
    If cmdOpen(intIndex).Tag > 0 And optAutomatic = True Then
        cmdOpen(intIndex).Caption = chkTerm(intIndex).Caption
        cmdOpen(intIndex).Tag = 0
            '������� ����������� "������" �����������
        cmdOpen(intIndex).Enabled = False
    End If
            '�������� �������� �������� �����������
            '   � �������� ��������
    imgParkingInData(intIndex).Tag = 0
    imgParkingOutData(intIndex).Tag = 0
    
    imgAccessInData(intIndex).Tag = 0
    imgAccessOutData(intIndex).Tag = 0
            '������� ���������� �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ���������
    imgParkingInData(intIndex).Enabled = True
    imgParkingOutData(intIndex).Enabled = True
    imgParkingInfoData(intIndex).Enabled = True
    
    imgAccessInData(intIndex).Enabled = True
    imgAccessOutData(intIndex).Enabled = True
    imgAccessInfoData(intIndex).Enabled = True

End Sub

            '��������� �������� ������� ��� ����� ������ - ������� "TimeOut"
Private Sub tmrPasswTimeOut_Timer()
            '������ �������� ������
    BeepSound
    
                '���������������� ������� - "TimeOut" ��� ����� ������
    gProtocol.strProtocName = "????????????????"
            '��������� ������
    gProtocol.strProtocPersonCode = ""
            '������
    gProtocol.strProtocStatus = ""
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "PASSWORD TimeOut"
            '�������� ������ � ���� "������� ���������"
    WriteProtocol

            ' "�������" ���� ������ ���������
    txtPassword.Text = ""
            '������� ����������� ���� ������
    txtPassword.Enabled = False
            ' "��������" �������� "������"
    lblPassword.Enabled = False
            '�������� �������� ������� ����� ������
    tmrPasswTimeOut.Enabled = False
            '������� ��������� ������������� "����������/���������"
    chkSetup.Enabled = True
            '������� ��������� ������ "Exit"
    cmdExit.Enabled = True
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then chkDummy.SetFocus
    
End Sub

            '��������� ��������� "������ ����" �� ���� ������
Private Sub txtPassword_Click()
            '����� ��������� "Exit" ��� "SetUp" - ����� �� ���������
    If tmrPasswTimeOut.Enabled = True Then Exit Sub
            ' "�������" ������ ������ ���������
    strPassword = ""
            '���������� �������� ������� ����� ������
    tmrPasswTimeOut.Enabled = True
            '���������� ����� �� ���� ������
    txtPassword.SetFocus
           '��������� ������ ���������� �� ���� ������ �� ��� �����
           '   ��� ��������� ������� ����� ������
    Do While strPassword = "" And tmrPasswTimeOut.Enabled = True
        DoEvents
    Loop
    
            '���������������� ������� - "���� ������ ������"
    gProtocol.strProtocName = "????????????????"
            '��������� ������
    gProtocol.strProtocPersonCode = strPassword
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "New PASSWORD"
            '�������� ������ � ���� "������� ���������"
    WriteProtocol
            '���������� ����� ������ � �������� ��������
    If strPassword <> "" Then txtPassword.Tag = strPassword
            ' "�������" ���� ������ ���������
    txtPassword.Text = ""
            '������� ����������� ���� ������
    txtPassword.Enabled = False
            ' "��������" �������� "������"
    lblPassword.Enabled = False
            '������� ��������� ������������� "����������/���������"
    chkSetup.Enabled = True
    

End Sub

            '��������� ��������� ��������� ������ ����� ������
Private Sub txtPassword_GotFocus()

            '������� ����������� ������ "Exit"
    cmdExit.Enabled = False
            '������� ����������� ������������� "����������/���������"
    chkSetup.Enabled = False

            ' "��������" �������� "������"
    lblPassword.Enabled = True

End Sub

            '��������� ����� � ������� ������
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
            '������ �������
    If KeyAscii = vbKeyReturn Then
    
    
            '������ ?
        strPassword = txtPassword.Text
        
            '���������������� ������� - "���� ������"
        gProtocol.strProtocName = "????????????????"
            '��������� ������
        gProtocol.strProtocPersonCode = txtPassword.Text
            '������
        gProtocol.strProtocStatus = "04 - Manager"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "PASSWORD Input"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
        
            '������ ������������ ������
        If txtPassword.Text = txtPassword.Tag Then
            ' "�������" ���� ������ ���������
            txtPassword.Text = ""
            '������� ����������� ���� ������
            txtPassword.Enabled = False
            ' "��������" �������� "������"
            lblPassword.Enabled = False
        End If
    End If

End Sub

            '��������� "������" ���� �� �������� "imgViewClose"
Private Sub imgViewClose_click(intIndex As Integer)

'''          '����� "frmDemo" �������� � ����� "����������"
'''    If frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
           
'''            '����������� ����� ���������� �������
'''            '   (������������ ���������� ��������)
'''        If gTimeShare = 1 Then
'''            '������� ������� "imgAccessOutData"
'''            If imgAccessOutData(intIndex).Visible = True Then
'''                Call mnuAccessServData_Click
'''            '������� ������� "imgParkingOutData"
'''            ElseIf imgParkingOutData(intIndex).Visible = True Then
'''                Call mnuParkingServData_Click
'''            End If
'''            Exit Sub
'''        End If
    
'''    End If

End Sub

            '������������� ������������ ������������� ����
            '  � ������ ������ ��� ��������������� (����� ���� "Controller"
            '  � ������� "DALLAS") ��������� �������� ����������� ���
            '  �����������
Public Function AutoRegDallasButton(ByVal vntPersonCode As Variant, _
intIndex As Integer, ByVal strAddrPortType As String)
            
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ������������  "PersonCode"
Dim intDallasCode  As Integer

            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataParkingIn" - ����� �� �������
    If frmDataParkingIn.Enabled = False And _
    Mid(Trim(strAddrPortType), 4) = "ParkI" Then Exit Function
            '������� ���������� ������ ������� �� �����������
            '  ������ �� ������ ����� "frmDataAccessIn" - ����� �� �������
    If frmDataAccessIn.Enabled = False And _
    Mid(Trim(strAddrPortType), 4) = "AcceI" Then Exit Function
    
            '������� ����������� ����� "frmDataParkingIn"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    frmDataParkingIn.Enabled = False
            '������� ����������� ����� "frmDataAccessIn"
    If Mid(Trim(strAddrPortType), 4) = "AcceI" Then _
    frmDataAccessIn.Enabled = False
            
            '���� ��� ��������� ������ �����������, �� ����������
            '  ��������� ������������ "PersonCode", "Info" � ������
            '  ������ �� �����-����� (+ ����) ��� ��������������� �������
            '  ����������� ����� ����������� "Controller" � ������� "DALLAS"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    intDallasCode = frmDataParkingIn.DallasButton(strAddrPortType, intIndex)
            '���� ��� ��������� ������ �����������, �� ���������
            '  ��������� ������������ "PersonCode", "Info" � ������
            '  ������ �� �����-����� (+ ����) ��� ��������������� �������
            '  ����������� ����� ����������� "Controller" � ������� "DALLAS"
    If Mid(Trim(strAddrPortType), 4, 2) = "AcceI" Then _
    intDallasCode = frmDataAccessIn.DallasButton(strAddrPortType, intIndex)

            '������� ��������� ����� "frmDataParkingIn"
    If Mid(Trim(strAddrPortType), 4) = "ParkI" Then _
    frmDataParkingIn.Enabled = True
            '������� ��������� ����� "frmDataAccessIn"
    If Mid(Trim(strAddrPortType), 4) = "AcceI" Then _
    frmDataAccessIn.Enabled = True

End Function

            '������ ��������� (�������� �� �����-�����, ��������
            '  ���� �/��� ��������� ����)
Public Sub PrintDocument(ByVal strProtocName As String, _
            ByVal strProtocPersonCode As String, _
            ByVal strProtocStatus As String, _
            ByVal strProtocTime As String, _
            ByVal strProtocDate As String, _
            ByVal strProtocReserve As String, _
            ByRef intError As Integer)
        
            '����� ���������� �����
Dim intFileNum As Integer
            '������ ��� ���������� ����� (� ��������� "����")
Dim strPathFileName As String
            '������ ����� "�������������" ���������� ����� ��������� ��������
Dim strCashPrinter(2) As String
            '����� ������� ������ ���������� ����� ��������� ��������
Dim intRowNum As Integer
            '����� ������� ��������� ������� � ������
Dim intPosNum As Integer
            '��� �������� ��� ������ ������� "Shell"
Dim vntShell As Variant
            '����������-������ "������ ���������"
Dim strDocument As String
            '����� ����������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ����������� �������
Dim strHour As String
Dim strMinute As String
            '������� �������
Dim intCount As Integer
            '������� ������� �������� ���� ���������
            '  �������� �����-����
Dim lngTimeCount As Long
            '������� ����������
Dim intWork As Integer
            '������� ������
Dim strWork As String

            '����� ������ ������ ��������� �� �������� �����-����
Dim Buffer() As Byte
           '���������� ��� �������������� ������� ������ ���������
            ' �� ��������-�����-���� � ����������������� �������������
Dim strBuffer As String
Dim intBuffer1 As Integer
Dim intBuffer2 As Integer

            '����� �������� ������
    intError = 0

            '���� ��������� ��� "������ ���������" (1 - ������� �������
            '  �������, 2 - ������� ���������, 4 - �������� ������� �������;
            '  �������� ����������: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7 Then
        strDocument = ""
            '������������� �������� ��������
        strDocument = strDocument + Chr(CLng(CByte(27))) + Chr(CLng(CByte(64)))
            '��������� ������� �� ������� ��������
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(82))) + Chr(CLng(CByte(0)))
            '������ ������ �� ������� �������� �� 5-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(5)))
            '��������� ������� ������ �� ������� �������� 7�7
        strDocument = strDocument + Chr(CLng(CByte(27))) + Chr(CLng(CByte(77)))
            '����������� ���������� "sp" ������� ��������� - �������
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            '�������������� ���� "00"H - �������
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            '������������ �������� ��������
        intWork = Len(Trim(gPrintSIAName))
        For intCount = 1 To intWork Step 1
            '������� ������ �� �������� �������� � ������� �������
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(gPrintSIAName), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            '������ ������ �� ������� �������� �� 1-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            '����������� ���������� "sp" ������� ��������� - �������
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            '�������������� ���� "00"H - �������
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            '������������ ������������� �������� ��������
        For intCount = 1 To 16 Step 1
            '������� ������ �� ������������� � ������� �������
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid("================", intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            '������ ������ �� ������� �������� �� 1-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            '����������� ���������� "sp" ������� ��������� - �������
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            '�������������� ���� "00"H - �������
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            '������������ ������������� ����
        strWork = "#### = " + Trim(strProtocPersonCode)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            '������� ������ �� ������������� ���� � ������� �������
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strWork), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            '������ ������ �� ������� �������� �� 1-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            '����������� ���������� "sp" ������� ��������� - �������
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            '�������������� ���� "00"H - �������
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            '������������ �������/����
        strWork = Trim(strProtocTime) + "||" + Trim(strProtocDate)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            '������� ������ �� �������/���� � ������� �������
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strWork), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            '������ ������ �� ������� �������� �� 1-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(1)))
            '����������� ���������� "sp" ������� ��������� - �������
        strDocument = strDocument + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + Chr(CLng(CByte(32))) + _
        Chr(CLng(CByte(32))) + Chr(CLng(CByte(32)))
            '�������������� ���� "00"H - �������
        strDocument = strDocument + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + _
        Chr(CLng(CByte(0))) + Chr(CLng(CByte(0))) + Chr(CLng(CByte(0)))
            '������������ ����������/�����
        intWork = Len(Trim(strProtocReserve))
        For intCount = 1 To intWork Step 1
            '������� ������ �� ����������/����� � ������� �������
            strDocument = Trim(strDocument) + _
            Chr(CByte(Asc(Mid(Trim(strProtocReserve), intCount, 1)))) + _
            Chr(CLng(CByte(13)))
        Next
            '������� �������
        strDocument = strDocument + Chr(CLng(CByte(13)))
            '������ ������ �� ������� �������� �� 10-� �����
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(97))) + Chr(CLng(CByte(10)))
            '������� ������ �� ������� �������
        prtPortDocument.Output = strDocument
             '����� ���������� �������� ������ �� ������� �������
        Do
        Loop Until prtPortDocument.OutBufferCount = 0
    End If
    
            '���� ��������� ��� "������ ���������" (1 - ������� ������� �������,
            '  2 - ������� ����� ����, 4 - �������� ������� �������;
            '  �������� ����������: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            '�� ������� ����������� ������� ��� ������� ����������� -
            '  ������ �����-���� �� ������������
        If frmDataAccessIn.Tag <> 1 And frmDataParkingIn.Tag <> 1 _
        Then GoTo BarCodeOK
            '��������� ������ ��� ���������� ������ �����-���� �� ��������
        On Error GoTo BarCodeError
            
            
            '������������� �������� �����-����
        strDocument = ""
        strDocument = Chr(CLng(CByte(27))) + Chr(CLng(CByte(64)))
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            '�������� �������� ����� �����
        prtPortBarCode.InBufferCount = 0
            '��������� ���������� � ������� �������� �����-����
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(97))) + _
        Chr(CLng(CByte(15)))
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� �������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
             
            '����� �� �������� �����-���� ���� ���������
        For lngTimeCount = 1 To 99000 Step 1
            If prtPortBarCode.InBufferCount = 4 Then
                Buffer = ""
            '���������� ������ � �������� ����� ��� ���������� ���������
                Buffer = prtPortBarCode.Input
            '�������� �������� ����� �����
                prtPortBarCode.InBufferCount = 0
                Exit For
            End If
            '���������� ��������� �������
            DoEvents
        Next
            '������������ �������� �����-����
        If lngTimeCount > 99000 Then GoTo BarCodeError
            
            '������������� ������ �� ������ ��������� � ����������������� ���
        strBuffer = ""
        intCount = 0
        Do While intCount <= 3
            intBuffer1 = (CByte(Buffer(intCount)) And CByte(240)) / 16
            intBuffer2 = CByte(Buffer(intCount)) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCount = intCount + 1
        Loop
            '������������ ��� ��������� �������� �����-����
        If Trim(strBuffer) <> "00400014" Then GoTo BarCodeError
            
            '���������� ��� ��������� �������� �����-����
            
            '������������ ����/�������
        strDocument = ""
        strWork = Trim(Format(Now, "h:mm:ss")) + " || " + _
        Trim(Format(Now, "dd/mm/yyyy"))
        strWork = Trim(Format(Now, "h:mm:ss"))
            '����
        intHour = Hour(strWork)
        If intHour < 10 Then
            strHour = "0" + Trim(Str(intHour))
        Else
            strHour = Trim(Str(intHour))
        End If
            '������
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
            '������� ������ �� ����/�������
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            '������� ������
        strDocument = strDocument + Chr(CLng(CByte(10)))
             '������������ �������� �������� ��� ������ ����������
        intWork = Len("BEZMAKSAS LAIKS 2 st.")
        For intCount = 1 To intWork Step 1
            '������� ������ �� �������� ��������
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid("BEZMAKSAS LAIKS 2 st.", intCount, 1))))
        Next
            '������� ������
        strDocument = strDocument + Chr(CLng(CByte(10)))
            '������������ ������������� �������� ��������
        For intCount = 1 To intWork Step 1
            '������� ������ �� �������������
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid("================================", intCount, 1))))
        Next
            '������� ������
        strDocument = strDocument + Chr(CLng(CByte(10)))
            '������������ ������������� ����
        strWork = Right(Trim(strProtocPersonCode), 10)
        strWork = "#### = " + Trim(strWork)
        intWork = Len(strWork)
        For intCount = 1 To intWork Step 1
            '������� ������ �� ������������� ����
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            '������ ������ �� ������� �������� �� 1-� �����
        strDocument = strDocument + Chr(CLng(CByte(10)))
            '������� ������ �� ������� �������
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
    
        strDocument = ""
            '���������� ������ �����-����
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(119))) + _
        Chr(CLng(CByte(4)))
            '���������� ���� ���� "�" ��� ������ �������� ��� �����-�����
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(102))) + _
        Chr(CLng(CByte(0)))
            '�������� ����� ������ �������� ��� �����-�����
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(72))) + _
        Chr(CLng(CByte(0)))
            '������� ������ �� ������� �������
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
    
        strDocument = ""
        strWork = "99" + Right(Trim(strProtocPersonCode), 10)
            '��������� ���� "EAN13"
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(107))) + _
        Chr(CLng(CByte(2)))
            '������������ �������� ����� �����-���� EAN13
        For intCount = 1 To 12 Step 1
            '������� ������ �� ������������� ����
            strDocument = strDocument + _
            Chr(CByte(Asc(Mid(strWork, intCount, 1))))
        Next
            '������������ ����������� ����� �����-���� EAN13
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
            '����������� �����
        strDocument = strDocument + Trim(Str(intWork))
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0

            '������ ������ �� �������� �����-���� �� N �����
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(100))) + Chr(CLng(CByte(9)))
            
            
'            '��������� ������ ������
'        strDocument = strDocument + Chr(CLng(CByte(27))) + _
'        Chr(CLng(CByte(109)))
            '������ ��������� ������
        strDocument = strDocument + Chr(CLng(CByte(27))) + _
        Chr(CLng(CByte(105)))
            
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            '������ ������ �� �������� �����-���� �� N �����
            ' (������������ ������)
        strDocument = ""
        For intCount = 1 To gTalonLength Step 1
            '������ ������ �� ������� �������� �� 1-� �����
            strDocument = strDocument + Chr(CLng(CByte(10)))
        Next
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� �������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            '�������� �������� ����� �����
        prtPortBarCode.InBufferCount = 0
            '��������� ���������� � ������� �������� �����-����
        strDocument = ""
        strDocument = strDocument + Chr(CLng(CByte(29))) + Chr(CLng(CByte(97))) + _
        Chr(CLng(CByte(15)))
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� �������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
             
            '����� �� �������� �����-���� ���� ���������
        For lngTimeCount = 1 To 99000 Step 1
            If prtPortBarCode.InBufferCount = 4 Then
                Buffer = ""
            '���������� ������ � �������� ����� ��� ���������� ���������
                Buffer = prtPortBarCode.Input
            '�������� �������� ����� �����
                prtPortBarCode.InBufferCount = 0
                Exit For
            End If
            '���������� ��������� �������
            DoEvents
        Next
            '������������ �������� �����-����
        If lngTimeCount > 99000 Then GoTo BarCodeError
            
            '������������� ������ �� ������ ��������� � ����������������� ���
        strBuffer = ""
        intCount = 0
        Do While intCount <= 3
            intBuffer1 = (CByte(Buffer(intCount)) And CByte(240)) / 16
            intBuffer2 = CByte(Buffer(intCount)) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCount = intCount + 1
        Loop
            '������������ ��� ��������� �������� �����-����
        If Trim(strBuffer) <> "00400014" Then GoTo BarCodeError
            
            '���������� ��� ��������� �������� �����-���� -
            '  ���������������� ������� "������ �����-���� �� ��������"
        gProtocol.strProtocName = "Print BarCode"
            '������������ ���
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '����������
        gProtocol.strProtocReserve = "BAR_CODE BOX"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            '����� �������� ������
        intError = 0
        
        GoTo BarCodeOK
BarCodeError:
            '������ �������� ������
        BeepSound
            '������� ������� ����� ���������
        lblErrorBarCodePrinter.Visible = True
        
            '������������ ����������� ���������
        strMessage = "BarCode Printer Error !!!"
            '�������� ���������
        Call SendMessage(strMessage)
        
            '���������������� ������� - "������ �������� �����-����"
        gProtocol.strProtocName = "Print BarCode"
            '������������ ���
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "BAR_CODE ERROR"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
            ' ���� ��� �������������� � ��������� ����
        If gNetPreprocNum = 0 Then
            '�������� ������������� ���� ������������������� ����������
            '  ������� �� "������� ������"
            Call frmTablePerson.AutoDelParking(strProtocPersonCode, _
            strProtocStatus)
        End If
            
            '��������� �������� ������
        intError = 1

BarCodeOK:
        On Error GoTo 0
    End If
            
            '���� ��������� ��� "������ ���������" (1 - ������� ������� �������,
            '  2 - ������� ���������, 4 - �������� ������� �������;
            '  �������� ����������: 1+2, 1+4, 2+4, 1+2+4)
            '  � ������� ������ �������
    If (gDocument = 4 Or gDocument = 5 Or gDocument = 6 Or gDocument = 7) And _
    intError = 0 Then
            '������� ����� ������ - ������ ���� �� ������������
        If Right(Trim(strProtocReserve), 9) = "000,00 Ls" Or _
        Right(Trim(strProtocReserve), 3) <> " Ls" Then GoTo CashOK
            '��������� ������ ��� ���������� ������ ���� �� �������� ��������
        On Error GoTo CashError
            
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� ������ "����" � ���������� ����� ��������� ��������
        strPathFileName = "C:\BarCashPrinter\Rs2810s.txt"
    
            '������� ��������� ���� ��������� �������� ��� �����
        Open strPathFileName For Input As intFileNum
            '���� �� ���� ������� ���������� ����� ��������� ��������
        For intRowNum = 1 To 3
            '������ ������ ���������� ����� ��������� �������� � �����
            strDocument = Input(34, intFileNum)
            '�������������� 1-�� ������ ���������� ����� ��������� ��������
            If intRowNum = 1 Then
            '������� ��������� "text=" � ������
                intPosNum = InStr(1, strDocument, "text=")
            '�������� ���������� ������
                If intPosNum = 0 Then
                    Close intFileNum
                    GoTo CashError
            '�������� � ��� "������������ ���" �������
                Else
                    strDocument = Left(strDocument, intPosNum + 5) + _
                    Trim(strProtocPersonCode) + """" + Right(strDocument, 2)
                    strCashPrinter(intRowNum - 1) = strDocument
                End If
            '�������������� 2-�� ������ ���������� ����� ��������� ��������
            ElseIf intRowNum = 2 Then
            '������� ��������� "deptnr=" � ������
                intPosNum = InStr(1, strDocument, "deptnr=")
            '�������� ���������� ������
                If intPosNum = 0 Then
                    Close intFileNum
                    GoTo CashError
            '������ � ��� "���� ������"
                Else
            '������� (����)����������� �������
                    If frmDataAccessIn.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "1" + _
                        Mid(strDocument, intPosNum + 8)
            '������� (����)�������� �������
                    ElseIf frmDataAccessOut.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "2" + _
                        Mid(strDocument, intPosNum + 8)
            '������� (����)��������� �������
                    ElseIf frmDataAccessServ.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "3" + _
                        Mid(strDocument, intPosNum + 8)
            '������� (����)����������� ����������
                    ElseIf frmDataParkingIn.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "4" + _
                        Mid(strDocument, intPosNum + 8)
            '������� (����)�������� ����������
                    ElseIf frmDataParkingOut.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "5" + _
                        Mid(strDocument, intPosNum + 8)
            '������� (����)��������� ����������
                    ElseIf frmDataParkingServ.Tag = 1 Then
                        strDocument = Left(strDocument, intPosNum + 6) + "6" + _
                        Mid(strDocument, intPosNum + 8)
                    End If
            '������� ��������� "amount=" � ������
                    intPosNum = InStr(intPosNum + 8, strDocument, "amount=")
            '�������� ���������� ������
                    If intPosNum = 0 Then
                        Close intFileNum
                        GoTo CashError
            '������ � ��� "��������� ������"
                    Else
                        strDocument = Left(strDocument, intPosNum + 6) + _
                        Right(strDocument, 2)
            '������� ������� "," � "����������" ���������
                        intPosNum = InStr(1, strProtocReserve, ",")
            '�������� ���������� ������
                        If intPosNum = 0 Then
                            Close intFileNum
                            GoTo CashError
            '�������������� ������� ������ "��������� ������"
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
            '�� �������������� 3-�� ������ ���������� ����� ��������� ��������
            ElseIf intRowNum = 3 Then
                strCashPrinter(intRowNum - 1) = strDocument
            End If
            
        Next
            '������� ��������� ���� ��������� ��������
        Close intFileNum
            '������� ��������� ���� ��������� �������� ��� ������
            '  (� ������ ��������� ������� - ��� ���������� �� �������)
        Open strPathFileName For Binary As intFileNum
            '���� �� ���� ������� ���������� ����� ��������� ��������
        For intRowNum = 1 To 3
            '�������� ������ ������ � ��������� ���� ��������� ��������
            Put #intFileNum, , strCashPrinter(intRowNum - 1)
        Next
            '������� ��������� ���� ��������� ��������
        Close intFileNum
    
        
        vntShell = Shell("C:\BarCashPrinter\Rs2810s.bat", 0)
        If vntShell <> 0 Then
            '���������������� ������� - "������ ���� �� �������� ��������"
            gProtocol.strProtocName = "Print Check"
            '������������ ���
            gProtocol.strProtocPersonCode = ""
            '������
            gProtocol.strProtocStatus = ""
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "CASH BOX"
            '�������� ������ � ���� "������� ���������"
            WriteProtocol
            GoTo CashOK
        End If
CashError:
            '������ �������� ������
        BeepSound
            '����� ���������
        If optEnglish = True Then
            MsgBox ("The CashPrinter Error")
        Else
            MsgBox ("Nepareizs 'CashPrinter' ")
        End If
        
            '���������������� ������� - "������ ���� �� �������� ��������"
        gProtocol.strProtocName = "Print Check"
            '������������ ���
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "CASH BOX ERROR"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
CashOK:
        On Error GoTo 0
    
    End If

End Sub

            '������ ��������� (Z_������ �� �������� �����-����)
Public Sub PrintZReport(ByVal strProtocName As String, _
            ByVal strProtocPersonCode As String, _
            ByVal strProtocStatus As String, _
            ByVal strProtocTime As String, _
            ByVal strProtocDate As String, _
            ByVal strProtocReserve As String, _
            ByVal strMoney_Report As String, _
            ByRef strZ_Report As String)
            '����� ���������� �����
Dim intFileNum As Integer
            '������ ��� ���������� ����� (� ��������� "����")
Dim strPathFileName As String
            '����� ������� ��������� ������� � ������
Dim intPosNum As Integer
            '��� �������� ��� ������ �������� �� �������� �����-����
Dim vntBuffer As Variant
            '����������-������ "������ ���������"
Dim strDocument As String
            '����� ������ Z_������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ������ Z_������
Dim strHour As String
Dim strMinute As String

            '���� ��������� ��� "������ ���������" (1 - ������� ������� �������,
            '  2 - ������� ����� ����, 4 - �������� ������� �������;
            '  �������� ����������: 1+2, 1+4, 2+4, 1+2+4)
    If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            '��������� ������ ��� ���������� ������
        On Error GoTo BarCodeError
            
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� ������ "����" � ���������� ����� �������� �����-����
        strPathFileName = "C:\BarCashPrinter\ZReport.txt"
    
            '������� ��������� ���� �������� �����-���� ��� �����
        Open strPathFileName For Input As intFileNum
            '������ ���� �������� �����-���� � �����
        strDocument = Input(FileLen(strPathFileName), intFileNum)
            '������� ��������� ���� �������� �����-����
        Close intFileNum
    
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
       
            '����
        intHour = Hour(gProtocol.strProtocTime)
        If intHour < 10 Then
            strHour = "0" + Trim(Str(intHour))
        Else
            strHour = Trim(Str(intHour))
        End If
            '������
        intMinute = Minute(gProtocol.strProtocTime)
        If intMinute < 10 Then
            strMinute = "0" + Trim(Str(intMinute))
        Else
            strMinute = Trim(Str(intMinute))
        End If

            '�������������� ������ "�����-����" ���������� ����� � ������
        intPosNum = InStr(1, strDocument, """""")
            '�������� ���������� ������
        If intPosNum = 0 Then
            GoTo BarCodeError
            '�������� � Z_����� "�����-����" ��� ������
        Else
            strDocument = Left(strDocument, intPosNum) + _
            strHour + ":" + strMinute + " / " + _
            CStr(Trim(gProtocol.strProtocDate)) + Mid(strDocument, intPosNum + 1)
        End If
            '�������������� ������ "Z_�����" ���������� ����� � ������
        intPosNum = InStr(intPosNum + 7, strDocument, "Z_Report = ")
            '�������� ���������� ������
        If intPosNum = 0 Then
            GoTo BarCodeError
            '�������� ����� � Z_�����
        Else
            strDocument = Left(strDocument, intPosNum + 10) + _
            strMoney_Report + Mid(strDocument, intPosNum + 11)
        End If
            
            
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
        tmrTimeOut(1).Tag = 0
        tmrTimeOut(1).Enabled = True
            ' ���� ������ ��������� �������� �����-����
        Do While DoEvents()
            '���� ������ ���������
            If prtPortBarCode.InBufferCount > 1 Then
            '��������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
                tmrTimeOut(1).Enabled = False
           '���������� ������ � �������� ����� ��� ���������� ���������
                vntBuffer = prtPortBarCode.Input
            '���������� ����� �� �������� �����-����
                If Len(vntBuffer) <= 3 Then
            '�� ���������������� �������
                    GoTo Protocol
            '������
                Else
            '�� ��������� ������
                    GoTo BarCodeError
                End If
            '��������� ������� TimeOut ��� ��������
            '   ���� ��������� �� �������� �����-����
            ElseIf tmrTimeOut(1).Tag <> 0 Then
            '��������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
                tmrTimeOut(1).Enabled = False
            '�� ��������� ������
                GoTo BarCodeError
            End If
        Loop
            
Protocol:
            '���������������� ������� - "������ Z_������"
        gProtocol.strProtocName = "Print Z_Report"
            '������������ ���
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "BAR_CODE BOX"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
            
        GoTo BarCodeOK
BarCodeError:
            '������ �������� ������
        BeepSound
            '����� ���������
        If optEnglish = True Then
            MsgBox ("The BarCodePrinter Error")
        Else
            MsgBox ("Nepareizs 'BarCodePrinter' ")
        End If
            '�������� ��������� ������� ����� "Z_������"
        strZ_Report = ""
        On Error GoTo 0
        
            '������� ������ �� ������� �����-���� (����� ��������)
        prtPortBarCode.Output = Chr(94) + Chr(64) + Chr(13)
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            
            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
        frmDataAccessIn.tmrParoleTimeOut.Enabled = True
            ' ���� ������ ��������� �������� �����-����
        Do While DoEvents()
            '���� ������ ���������
            If prtPortBarCode.InBufferCount > 1 Then
            '��������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
                frmDataAccessIn.tmrParoleTimeOut.Enabled = False
            '�� ������ ������ ��������
                Exit Do
            '��������� ������� TimeOut ��� ��������
            '   ���� ��������� �� �������� �����-����
            ElseIf frmDataAccessIn.tmrParoleTimeOut.Enabled = False Then
            '�� ������ ������ ��������
                Exit Do
            End If
        Loop
            '������� ������������ ������ "OK" � "Cancel" ����� "frmDataAccessIn"
        frmDataAccessIn.cmdOK.Enabled = False
        frmDataAccessIn.cmdCancel.Enabled = False
            '������� ���������������� ���� ��� �������� �����-���� - ������� �����
        prtPortBarCode.PortOpen = False
            'O������ ���������������� ���� ��� �������� �����-����
        prtPortBarCode.PortOpen = True
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� ������ "����" � ���������� ����� �������� �����-����
        strPathFileName = "C:\BarCashPrinter\ZReport.txt"
            '������� ��������� ���� �������� �����-���� ��� �����
        Open strPathFileName For Input As intFileNum
            '������ ���� �������� �����-���� � �����
        strDocument = Input(FileLen(strPathFileName), intFileNum)
            '������� ��������� ���� �������� �����-����
        Close intFileNum
            '������� ������ �� ������� �����-����
        prtPortBarCode.Output = strDocument
             '����� ���������� �������� ������ �� ������� �����-����
        Do
        Loop Until prtPortBarCode.OutBufferCount = 0
            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
        frmDataAccessIn.tmrParoleTimeOut.Enabled = True
            ' ���� ������ ��������� �������� �����-����
        Do While DoEvents()
            '���� ������ ���������
            If prtPortBarCode.InBufferCount > 1 Then
            '��������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� �������� �����-����
                frmDataAccessIn.tmrParoleTimeOut.Enabled = False
            '����� �� �������� ������ ��������
                Exit Do
            '��������� ������� TimeOut ��� ��������
            '   ���� ��������� �� �������� �����-����
            ElseIf frmDataAccessIn.tmrParoleTimeOut.Enabled = False Then
            '����� �� �������� ������ ��������
                Exit Do
            End If
        Loop
            
            '���������������� ������� - "������ Z_������"
        gProtocol.strProtocName = "Print Z_Report"
            '������������ ���
        gProtocol.strProtocPersonCode = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "BAR_CODE ERROR"
            '�������� ������ � ���� "������� ���������"
        WriteProtocol
        
BarCodeOK:
        On Error GoTo 0
    End If

End Sub
            
            '��������� ������ ������ � ���� "������� ���������"
Public Sub WriteProtocol()
            
            '�������� ������ � ���� "������� ���������"
    Put gProtocFileNum, gProtocRowNum, gProtocol
            '����� ��������� ��������� ������ (������) � ����� "������� ���������"
    gProtocRowNum = gProtocRowNum + 1
            '���� "������� ���������" ������� �������������
    If gProtocRowNum > 32760 Then
            '������������� ����� "������� ���������"
        WriteProtocolToArchives
    End If
    
End Sub

            '��������� �������������� ������ "Protocol to Archives..."
Private Sub WriteProtocolToArchives()
            '������ ��� ����� ������ (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ��������� �����
Dim intFileNum As Integer
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer

            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = strPathFileName + gHost + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
            ' ���� ��� ������������
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = strPathFileName + gPreprocName + "_" + _
        Left(Format(Now, "dd/mm/yyyy"), 2) + "_" + _
        Mid(Format(Now, "dd/mm/yyyy"), 4, 2) + "_" + _
        Right(Format(Now, "dd/mm/yyyy"), 2)
    End If
            
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
            '��������� ������� � ������ ����� ����� (�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '������� "������" �������� ����, ���� �� ����������
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            '��������� ������
    On Error GoTo UnDefError
            '������� ��������� �������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
    
            '���� �� ���� ������� "������� ���������"
    For intRowNum = 1 To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
        Get gProtocFileNum, intRowNum, gProtocol
            '������� ������ "������� ���������" � �������� ����
        Put intFileNum, intRowNum, gProtocol
    Next
            '������� ���������  �������� ����
    Close intFileNum
            
            ' ���� ��� ������������
    If gPreprocName <> "" Then
            '������������ ����������� ���������
        strMessage = "Archive" + " " + Mid(strPathFileName, intSymbPos)
            '�������� ���������
        Call SendMessage(strMessage)
    End If
        
             '������� "�������" ���� "������� ���������"
    Close gProtocFileNum
           '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableProtocol.dat"
        
            '��������� ������� � ������ ����� ����� (�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '�������  "�������" ���� "������� ���������"
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
            '�������� ��������� ����� ����� ��� ����� "������� ���������"
    gProtocFileNum = FreeFile
            '������� ����� ���� "������� ���������" ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '����� ��������� "���������" ������ � ����� ����� "������� ���������"
    gProtocRowNum = 1
            '���������� ������� ���������� ��������� ������� � ������������ �����
    mnuSaveProtocol.Checked = True
    mnuSaveProtocolAs.Checked = False
    
    Exit Sub
            '�������������� ������
UnDefError:
            '������ �������� ������
    BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

End Sub
