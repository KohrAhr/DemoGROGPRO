VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTablePerson 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_person"
   ClientHeight    =   6855
   ClientLeft      =   1035
   ClientTop       =   1095
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find..."
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
      Left            =   4080
      TabIndex        =   46
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefaultPers 
      Cancel          =   -1  'True
      Caption         =   "Default from HDD"
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
      Left            =   7800
      TabIndex        =   44
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtPort 
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
      Height          =   288
      Left            =   7680
      TabIndex        =   42
      Top             =   3360
      Width           =   255
   End
   Begin VB.TextBox txtReservation 
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
      Height          =   288
      Left            =   5520
      TabIndex        =   41
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Frame fraCalendar 
      Caption         =   "Working"
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
      Height          =   612
      Left            =   3960
      TabIndex        =   36
      Top             =   2400
      Width           =   4692
      Begin VB.OptionButton optAlways 
         Caption         =   "Always"
         Height          =   252
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optStandard 
         Caption         =   "Standard"
         Height          =   252
         Left            =   1320
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   972
      End
      Begin VB.OptionButton optSpecial 
         Caption         =   "Special [Time/Ter/Cal]"
         Height          =   252
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtType 
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
      Height          =   288
      Left            =   8040
      TabIndex        =   33
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtAddress 
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
      Height          =   288
      Left            =   7200
      TabIndex        =   32
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtPersonCode 
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
      Height          =   288
      Left            =   5280
      TabIndex        =   31
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtName 
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
      Height          =   288
      Left            =   4920
      TabIndex        =   28
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Control"
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
      Height          =   612
      Left            =   3960
      TabIndex        =   23
      Top             =   1320
      Width           =   4692
      Begin VB.OptionButton optBlackCard 
         Caption         =   "Black Card"
         Height          =   252
         Left            =   2400
         TabIndex        =   27
         Top             =   240
         Width           =   1212
      End
      Begin VB.OptionButton optRelay 
         Caption         =   "Relay"
         Height          =   252
         Left            =   3600
         TabIndex        =   26
         Top             =   240
         Width           =   852
      End
      Begin VB.OptionButton optComputer 
         Caption         =   "Computer"
         Height          =   252
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1092
      End
      Begin VB.OptionButton optTerminal 
         Caption         =   "Terminal"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.HScrollBar hsbMinute 
      Enabled         =   0   'False
      Height          =   252
      LargeChange     =   10
      Left            =   6600
      Max             =   59
      TabIndex        =   18
      Top             =   2040
      Width           =   1452
   End
   Begin VB.HScrollBar hsbHour 
      Enabled         =   0   'False
      Height          =   252
      LargeChange     =   10
      Left            =   4320
      Max             =   23
      TabIndex        =   16
      Top             =   2040
      Width           =   1452
   End
   Begin VB.Frame fraColName 
      Caption         =   "Options"
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
      Height          =   3372
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   1815
      Begin VB.CheckBox chkFromToTime 
         Caption         =   "vFrom    To"
         Height          =   495
         Left            =   840
         TabIndex        =   22
         Top             =   1680
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.OptionButton optReservation 
         Caption         =   "Reservation"
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   1452
      End
      Begin VB.OptionButton optCalendar 
         Caption         =   "Calendar"
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1452
      End
      Begin VB.OptionButton optTime 
         Caption         =   "Time"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1452
      End
      Begin VB.OptionButton optStatus 
         Caption         =   "Status"
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1452
      End
      Begin VB.OptionButton optPersonCode 
         Caption         =   "PersonCode"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1452
      End
      Begin VB.OptionButton optName 
         Caption         =   "Pers. or Term."
         Height          =   312
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmdCorrection 
      Caption         =   "Correction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1092
   End
   Begin VB.ListBox lstName 
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
      Height          =   1950
      ItemData        =   "frmTablePerson.frx":0000
      Left            =   120
      List            =   "frmTablePerson.frx":0002
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "SaveAs..."
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
      Left            =   2760
      TabIndex        =   5
      Top             =   6240
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   6240
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete..."
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
      Left            =   6480
      TabIndex        =   3
      Top             =   6240
      Width           =   1092
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
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
      Left            =   5280
      TabIndex        =   2
      Top             =   6240
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid grdTablePerson 
      Height          =   2295
      Left            =   2160
      TabIndex        =   0
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   9
      Cols            =   6
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAccess 
      Alignment       =   2  'Center
      Caption         =   "0 "
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
      Height          =   255
      Left            =   480
      TabIndex        =   45
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7800
      X2              =   7800
      Y1              =   3360
      Y2              =   3120
   End
   Begin VB.Line Line17 
      BorderWidth     =   4
      X1              =   8760
      X2              =   7800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line16 
      BorderWidth     =   3
      X1              =   8160
      X2              =   7800
      Y1              =   3360
      Y2              =   3120
   End
   Begin VB.Line Line15 
      BorderWidth     =   3
      X1              =   7800
      X2              =   7440
      Y1              =   3120
      Y2              =   3360
   End
   Begin VB.Line Line14 
      BorderWidth     =   3
      X1              =   7800
      X2              =   8040
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line Line13 
      BorderWidth     =   3
      X1              =   7320
      X2              =   7560
      Y1              =   1080
      Y2              =   1320
   End
   Begin VB.Line Line12 
      BorderWidth     =   3
      X1              =   7680
      X2              =   7680
      Y1              =   1080
      Y2              =   1320
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      Caption         =   " 2-8 Port "
      Enabled         =   0   'False
      Height          =   495
      Left            =   7560
      TabIndex        =   43
      Top             =   480
      Width           =   375
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   8880
      X2              =   8880
      Y1              =   3720
      Y2              =   120
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   120
      X2              =   8880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8880
      X2              =   2040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   3720
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   120
   End
   Begin VB.Label lblReservation 
      Alignment       =   2  'Center
      Caption         =   "Reservation "
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
      Height          =   255
      Left            =   4080
      TabIndex        =   40
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   8040
      X2              =   8760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   8760
      X2              =   8760
      Y1              =   3120
      Y2              =   1080
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Caption         =   """xxxxx"" Type "
      Enabled         =   0   'False
      Height          =   495
      Left            =   8040
      TabIndex        =   35
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      Caption         =   "01-15  Addr. "
      Enabled         =   0   'False
      Height          =   495
      Left            =   6960
      TabIndex        =   34
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblPersonCode 
      Alignment       =   2  'Center
      Caption         =   "PersonCode "
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
      Height          =   255
      Left            =   3960
      TabIndex        =   30
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Caption         =   "Name "
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
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMinute0 
      Alignment       =   2  'Center
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label lblMinute59 
      Alignment       =   2  'Center
      Caption         =   "59min"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblHour23 
      Alignment       =   2  'Center
      Caption         =   "23h"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblHour0 
      Alignment       =   2  'Center
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label lblPersOrTerm 
      Alignment       =   2  'Center
      Caption         =   "Pers. or Term. "
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
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "frmTablePerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������� ����� �������������� ������ "������� ������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "������� ������"
Dim intColNumCorr As Integer
            '������� ����� �����
Dim intFileNum As Integer
            '������ "������� ������"
Dim gPerson As PersonInfo
            '������ ����������� ���������
Dim strMessage As String


            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            
            '���� �� ����������� ��������� � "������� ������"
    If gChangesTablePerson = True Then
            '������ �������� ������
        frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '���������� "T������ ������" � ����� �� ���������
            cmdSave_Click
        End If
    End If
    
            '������� ������������ �������� ���������� ���������� "������� ������"
    fraColName.Enabled = False
    optName.Enabled = False
    optPersonCode.Enabled = False
    optStatus.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optCalendar.Enabled = False
    optReservation.Enabled = False
    txtName.Enabled = False
    lstName.Enabled = False
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    optTerminal.Enabled = False
    optComputer.Enabled = False
    optBlackCard.Enabled = False
    optRelay.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    optAlways.Enabled = False
    optStandard.Enabled = False
    optSpecial.Enabled = False
    txtReservation.Enabled = False
            '�������� ��������� ����
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            '�������� ������ ����
    lstName.Clear
            '�������� ������� ��������� ��������� � "������� ������"
    gChangesTablePerson = False
            '������� ��������� ������� �����
    frmTablePerson.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '���������
Private Sub cmdCorrection_Click()
            
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
    
    Else
            '������� ���������� ��������� �������� ����������
            '   ���������� "������� ������"
        fraColName.Enabled = True
        optName.Enabled = True
        optName.Value = True
        optPersonCode.Enabled = True
        optStatus.Enabled = True
        optTime.Enabled = True
        chkFromToTime.Enabled = True
        optCalendar.Enabled = True
        optReservation.Enabled = True
        lblName.Enabled = True
        txtName.Enabled = True
        lblPersOrTerm.Enabled = True
        lstName.Enabled = True
            '�������� ��������� ����
        txtName.Text = ""
        txtPersonCode.Text = ""
        txtReservation.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtType.Text = ""
            '�������� ������ ����
        lstName.Clear
    
            '������� "Person or Terminal"
        gTablePerson.Col = 0
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNumCorr = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNumCorr
            '���������� ������ "lstName" �������� �� "������� ������"
            lstName.AddItem gTablePerson.Text
        Next
            '�������  ������� ������
        lstName.ListIndex = 0
            '����� �������������� ������ - (1)
        intRowNumCorr = 1
        gTablePerson.Row = intRowNumCorr
            '�������� �����
        optName_Click
    End If
    
End Sub
            
            '��������� �������������� ������������ ��������
            ' ���������� "������� ������"
Private Sub cmdDefaultPers_Click()
            
            '������� ������������ �������� ����������
            '  ���������� "������� ������"
    fraColName.Enabled = False
    txtName.Enabled = False
    lstName.Enabled = False
    txtName.Enabled = False
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    txtReservation.Enabled = False
            '�������� ��������� ����
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            '�������� ������ ����
    lstName.Clear
            
            '�������� ����� "������� ������"
    Form_Load
            '���������� ����� �� ������ "Correction"
    If frmTablePerson.Visible = True Then cmdCorrection.SetFocus

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "������� ������"
Private Sub grdTablePerson_Click()
            '��������� "��������"
    If lstName.Enabled = True Then
            '����� �������������� ������ "������� ������"
        intRowNumCorr = gTablePerson.RowSel
        gTablePerson.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstName.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "������� ������"
        intColNumCorr = gTablePerson.ColSel
        gTablePerson.Col = intColNumCorr
            '����� �������������� ������ "������� ������"
        lstName_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstName.Left, Y:=lstName.Top
            '����� ��������������� ������� "������� ������"
        Select Case intColNumCorr
            Case 1
            optPersonCode.Value = True
            '���������� ����� �� ��������� ���� ��� ���������
            txtPersonCode.SetFocus
            Case 2
            optStatus.Value = True
            Case 3
            optTime.Value = True
            Case 4
            optCalendar.Value = True
            Case 5
            If optReservation.Value = True Then
                optReservation_Click
            Else
                optReservation.Value = True
            End If
        End Select
    End If
        
End Sub


            '����� �������������� ������ "������� ������"
Private Sub lstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� �������������� ������ "������� ������"
        intRowNumCorr = lstName.ListIndex + 1
        gTablePerson.Row = intRowNumCorr
        gTablePerson.Col = 0
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
        txtName.Text = gTablePerson.Text
            '����� �������������� ������� "������� ������" - "PersonCode"
        gTablePerson.Col = 1
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
        txtPersonCode.Text = gTablePerson.Text
            '����� �������������� ������� "������� ������" - "Status"
        gTablePerson.Col = 2
            '�� ����������� ����� "Relay"
        If Left(gTablePerson.Text, 2) <> "03" Then
            '����� ��������������� ������� "������� ������"
            gTablePerson.Col = 5
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
            txtReservation.Text = gTablePerson.Text
            '������� ����������� ��� ��������� ��������� �����
            txtAddress.Text = ""
            txtType.Text = ""
            txtPort.Text = ""
            '����������� ����� "Relay"
        Else
            '����� ��������������� ������� "������� ������"
            gTablePerson.Col = 4
            '�������������� ��������� ������ "Calendar" � "������� ������"
            gTablePerson.Text = "00 - Always"
            '���������� �������  ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            
            '����� ��������������� ������� "������� ������"
            gTablePerson.Col = 5
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
            txtAddress.Text = Left(gTablePerson.Text, 2)
            txtPort.Text = Mid(gTablePerson.Text, 3, 1)
            txtType.Text = Mid(gTablePerson.Text, 4)
            '������� ������������ ��� ��������� ���������� ����
            txtReservation.Text = ""
        End If
            '������������ ����� ��������������� ������� "������� ������"
        gTablePerson.Col = intColNumCorr
    End If

End Sub

            '������� ����� - "Pers. or Term."
Private Sub optName_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 0
    gTablePerson.Col = intColNumCorr
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
    txtName.Text = gTablePerson.Text
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    txtName.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtName.SetFocus
            '����� ������������� ������� "������� ������"
    gTablePerson.Col = 1
            '����������� ������ "������� ������" � ��������� ���� ��� �����������
    txtPersonCode.Text = gTablePerson.Text
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    txtReservation.Enabled = False
            '����� �������������� ������� "������� ������" - "Status"
    gTablePerson.Col = 2
            '�� ����������� ����� "Relay"
    If Left(gTablePerson.Text, 2) <> "03" Then
            '����� ������������� ������� "������� ������"
        gTablePerson.Col = 5
            '����������� ������ "������� ������" � ��������� ���� ��� �����������
        txtReservation.Text = gTablePerson.Text
            '������� ����������� ��� ����������� ��������� �����
        txtAddress.Text = ""
        txtType.Text = ""
        txtPort.Text = ""
            '����������� ����� "Relay"
    Else
            '����� ������������� ������� "������� ������"
        gTablePerson.Col = 5
            '����������� ������ "������� ������" � ��������� ���� ��� �����������
        txtAddress.Text = Left(gTablePerson.Text, 2)
        txtPort.Text = Mid(gTablePerson.Text, 3, 1)
        txtType.Text = Mid(gTablePerson.Text, 4)
            '������� ������������ ��� ����������� ���������� ����
        txtReservation.Text = ""
    End If
            '������������ ����� ��������������� ������� "������� ������"
    gTablePerson.Col = intColNumCorr

End Sub
            
            '������� ����� - "PersonCode"
Private Sub optPersonCode_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 1
    gTablePerson.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = True
    txtPersonCode.Enabled = True
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
    txtPersonCode.Text = gTablePerson.Text
            '���������� ����� �� ��������� ���� ��� ���������
    txtPersonCode.SetFocus
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblType.Enabled = False
    txtType.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    optTerminal.Enabled = False
    optComputer.Enabled = False
    optBlackCard.Enabled = False
    optRelay.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    optAlways.Enabled = False
    optStandard.Enabled = False
    optSpecial.Enabled = False
    lblReservation.Enabled = False
    txtReservation.Enabled = False

End Sub


            '������� ����� - "Status"
Private Sub optStatus_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 2
    gTablePerson.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = False
    txtPersonCode.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblType.Enabled = False
    txtType.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = True
    optTerminal.Enabled = True
    optComputer.Enabled = True
            '�������� ����� "Computer"
    optComputer.Value = True
    optBlackCard.Enabled = True
    optRelay.Enabled = True
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    optAlways.Enabled = False
    optStandard.Enabled = False
    optSpecial.Enabled = False
    lblReservation.Enabled = False
    txtReservation.Enabled = False

End Sub

            '������� ����� - "Time"
Private Sub optTime_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 3
    gTablePerson.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = False
    txtPersonCode.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblType.Enabled = False
    txtType.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    optTerminal.Enabled = False
    optComputer.Enabled = False
    optBlackCard.Enabled = False
    optRelay.Enabled = False
    lblHour0.Enabled = True
    lblHour23.Enabled = True
    lblMinute0.Enabled = True
    lblMinute59.Enabled = True
    hsbHour.Enabled = True
    hsbMinute.Enabled = True
    fraCalendar.Enabled = False
    optAlways.Enabled = False
    optStandard.Enabled = False
    optSpecial.Enabled = False
    lblReservation.Enabled = False
    txtReservation.Enabled = False

End Sub
            
            '������� ����� - "Calendar"
Private Sub optCalendar_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 4
    gTablePerson.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = False
    txtPersonCode.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblType.Enabled = False
    txtType.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    optTerminal.Enabled = False
    optComputer.Enabled = False
    optBlackCard.Enabled = False
    optRelay.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = True
    optAlways.Enabled = True
    optStandard.Enabled = True
    optSpecial.Enabled = True
    lblReservation.Enabled = False
    txtReservation.Enabled = False
            '����� �������������� ������� "������� ������" - "Status"
    gTablePerson.Col = 2
            '����������� ����� "Terminal" ��� "Relay"
    If Left(gTablePerson.Text, 2) = "00" Or Left(gTablePerson.Text, 2) = "03" Then
            '�������� ����� "Standard"
        optAlways.Value = True
            '����������� ����� "Terminal" ��� "Relay"
    Else
            '�������� ����� "Standard"
        optStandard.Value = True
    End If
            '������������ ����� ��������������� ������� "������� ������"
    gTablePerson.Col = intColNumCorr


End Sub

            '������� ����� "Reservation"
Private Sub optReservation_Click()
            '����� ��������������� ������� "������� ������"
    intColNumCorr = 5
            '������� (��)���������� ��������� �������� ������. ���������� "������� ������"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = False
    txtPersonCode.Enabled = False
    fraStatus.Enabled = False
    optTerminal.Enabled = False
    optComputer.Enabled = False
    optBlackCard.Enabled = False
    optRelay.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    optAlways.Enabled = False
    optStandard.Enabled = False
    optSpecial.Enabled = False
    
                '����� �������������� ������� "������� ������" - "Status"
    gTablePerson.Col = 2
            '�� ����������� ����� "Relay"
    If Left(gTablePerson.Text, 2) <> "03" Then
            '����� ��������������� ������� "������� ������"
        gTablePerson.Col = intColNumCorr
        lblReservation.Enabled = True
        txtReservation.Enabled = True
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
        txtReservation.Text = gTablePerson.Text
            '���������� ����� �� ��������� ���� ��� ���������
        txtReservation.SetFocus
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblType.Enabled = False
        txtType.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
            '������ ����������� ��� ��������� ��������� �����
        txtAddress.Text = ""
        txtType.Text = ""
        txtPort.Text = ""
            '����������� ����� "Relay"
    Else
            '����� ��������������� ������� "������� ������"
        gTablePerson.Col = 4
            '�������������� ��������� ������ "Calendar" � "������� ������"
        gTablePerson.Text = "00 - Always"
            '���������� �������  ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            
            '����� ��������������� ������� "������� ������"
        gTablePerson.Col = intColNumCorr
        lblReservation.Enabled = False
        txtReservation.Enabled = False
        lblAddress.Enabled = True
        txtAddress.Enabled = True
            '����������� ������ "������� ������" � ��������� ���� ��� ���������
        txtAddress.Text = Left(gTablePerson.Text, 2)
        txtPort.Text = Mid(gTablePerson.Text, 3, 1)
        txtType.Text = Mid(gTablePerson.Text, 4)
            '���������� ����� �� ��������� ���� ��� ���������
        txtAddress.SetFocus
        lblType.Enabled = True
        txtType.Enabled = True
        lblPort.Enabled = True
        txtPort.Enabled = True
            '������ ������������ ��� ��������� ���������� ����
        txtReservation.Text = ""
    End If

End Sub

            '��������� ����� � ������� ��������������� ����� "Person or Terminal"
Private Sub txtName_KeyPress(KeyAscii As Integer)
            '��� �������
    If KeyAscii = vbKeyReturn Then
            '��� � ���������� ���������
        If Len(Trim(txtName.Text)) < 17 Then
            '��������� ����� "Person or Terminal" � "������� ������"
        gTablePerson.Text = Trim(txtName.Text)
            '���������� �������  ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            '�������� ����� "optPersonCode"
        optPersonCode.Value = True
            Exit Sub
            '��� � ������������ ���������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo PersonCodeError
            '������������ ��� � ���������� ���������
        If Len(Trim(txtPersonCode.Text)) = 16 Then
            '��������� ������ "PersonCode" � "������� ������"
            gTablePerson.Text = Trim(txtPersonCode.Text)
            '���������� �������  ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            '�������� ����� "optStatus"
            optStatus.Value = True
            Exit Sub
            '������������ ��� � ������������ ���������
PersonCodeError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            '������� ����� - "Terminal"
Private Sub optTerminal_GotFocus()
            '��������� ������ "Status" � "������� ������"
    gTablePerson.Text = "00 - Terminal"
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 4
            '�������������� ��������� ������ "Calendar" � "������� ������"
    gTablePerson.Text = "00 - Always"
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 2
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub

            '������� ����� - "Computer"
Private Sub optComputer_GotFocus()
            '��������� ������ "Status" � "������� ������"
    gTablePerson.Text = "01 - Computer"
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub

            '������� ����� - "BlackCard"
Private Sub optBlackCard_GotFocus()
            '��������� ������ "Status" � "������� ������"
    gTablePerson.Text = "02 - Black card"
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub

            '������� ����� - "Relay"
Private Sub optRelay_GotFocus()
            '��������� ������ "Status" � "������� ������"
    gTablePerson.Text = "03 - Relay"
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 4
            '�������������� ��������� ������ "Calendar" � "������� ������"
    gTablePerson.Text = "00 - Always"
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 2
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Hour"
Private Sub hsbHour_Scroll()
    hsbHour_Change
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Hour"
Private Sub hsbHour_Change()
            '������ ���������� ���������
    If chkFromToTime.Value = 1 Then
            '��������� ������ "Time" � "������� ������"
        If hsbHour.Value < 10 Then
            gTablePerson.Text = "0" + Trim(Str(hsbHour.Value)) + Mid(gTablePerson.Text, 3)
        Else
            gTablePerson.Text = Trim(Str(hsbHour.Value)) + Mid(gTablePerson.Text, 3)
        End If
            '����� ���������� ���������
    Else
            '��������� ������ "Time" � "������� ������"
        If hsbHour.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 6) + "0" + Trim(Str(hsbHour.Value)) _
            + Mid(gTablePerson.Text, 9)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 6) + Trim(Str(hsbHour.Value)) _
            + Mid(gTablePerson.Text, 9)
        End If
    End If
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True
    
End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Minute"
Private Sub hsbMinute_Scroll()
    hsbMinute_Change
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Minute"
Private Sub hsbMinute_Change()
            '������ ���������� ���������
    If chkFromToTime.Value = 1 Then
            '��������� ������ "Time" � "������� ������"
        If hsbMinute.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 3) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 6)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 6)
        End If
            '����� ���������� ���������
    Else
            '��������� ������ "Time" � "������� ������"
        If hsbMinute.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 9) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 12)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 9) + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 12)
        End If
    End If
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub

            '������� ����� - "Always"
Private Sub optAlways_GotFocus()
            '��������� ������ "Calendar" � "������� ������"
    gTablePerson.Text = "00 - Always"
            '���������� �������  ��������� ��������� � "������� ������"
    gChangesTablePerson = True

End Sub

            '������� ����� - "Standard"
Private Sub optStandard_GotFocus()
            '����� �������������� ������� "������� ������" - "Status"
    gTablePerson.Col = 2
            '�� ����������� ����� "Terminal" � "Relay"
    If Left(gTablePerson.Text, 2) <> "00" And Left(gTablePerson.Text, 2) <> "03" Then
            '����� ��������������� ������� "������� ������"
        gTablePerson.Col = 4
            '��������� ������ "Calendar" � "������� ������"
        gTablePerson.Text = "01 - Standard"
            '���������� �������  ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 4

End Sub

            '������� ����� - "Special"
Private Sub optSpecial_GotFocus()
            '����� �������������� ������� "������� ������" - "Status"
    gTablePerson.Col = 2
            '�� ����������� ����� "Terminal" � "Relay"
    If Left(gTablePerson.Text, 2) <> "00" And Left(gTablePerson.Text, 2) <> "03" Then
            '����� ��������������� ������� "������� ������"
        gTablePerson.Col = 4
            '��������� ������ "Calendar" � "������� ������"
        gTablePerson.Text = "02 - Special"
            '���������� �������  ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
            '����� ��������������� ������� "������� ������"
    gTablePerson.Col = 4

End Sub
            
            '��������� ����� � ������� ��������������� ���� "Reservation"
Private Sub txtReservation_KeyPress(KeyAscii As Integer)
            '����� ������� �������� "/" � �������������� ����
Dim intPosNum As Integer
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '����� �������������� ������� "������� ������" - "Calendar"
        gTablePerson.Col = 4
            '����������� ����� "Special"
        If Left(gTablePerson.Text, 2) = "02" Then
            '����� ��������������� ������� "������� ������"
            gTablePerson.Col = intColNumCorr
            '����������� �������� ..."/"..."/"... - ������. ����.: ����. ������., ������. � ���������
            intPosNum = InStr(1, Trim(txtReservation.Text), "/")
            If intPosNum <> 0 And InStr(intPosNum + 1, Trim(txtReservation.Text), "/") <> 0 _
            And Len(Trim(txtReservation.Text)) < 9 Then
            '������ �������������� �������� ������
                On Error GoTo TimeTerCalError
            '��������� ������ "������� �������"(...//) � ������ "Reservation" "������� ������"
                If intPosNum > 1 Then
            '����� "������� �������" � ���������� ��������� ������� (0/99)
                    If intPosNum < 4 And Left(Trim(txtReservation.Text), intPosNum - 1) < 100 Then
            
            '����� "������� �������" � ������������ ��������� �������
                    Else
                        GoTo TimeTerCalError
                    End If
                End If
            '��������� ������ "������� ����������"(/.../) � ������ "Reservation" "������� ������"
                If InStr(intPosNum + 1, Trim(txtReservation.Text), "/") > intPosNum + 1 Then
            '����� "������� ����������" � ���������� ��������� ������� (1/99)
                    If InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum < 4 And _
                    Mid(Trim(txtReservation.Text), intPosNum + 1, _
                    InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum - 1) > 0 And _
                    Mid(Trim(txtReservation.Text), intPosNum + 1, _
                    InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum - 1) < 100 Then
            
            '����� "������� ����������" � ������������ ��������� �������
                    Else
                        GoTo TimeTerCalError
                    End If
                End If
            '��������� ������ "������� ��������������� ���������"(//...) � ������
            '  "Reservation" "������� ������"
                If Len(Trim(txtReservation.Text)) > InStr(intPosNum + 1, _
                Trim(txtReservation.Text), "/") Then
            '����� "������� ��������������� ���������" � ���������� ���������
            '  ������� (1/99)
                    If Len(Trim(txtReservation.Text)) - InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") < 3 And _
                    Mid(Trim(txtReservation.Text), InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 1) > 0 And _
                    Mid(Trim(txtReservation.Text), InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 1) < 100 Then
            
            '����� "������� ��������������� ���������" � ������������ ��������� �������
                    Else
                            GoTo TimeTerCalError
                    End If
                End If
            
            '��������� � ������ "Reservation" "������� ������"
                gTablePerson.Text = Trim(txtReservation.Text)
            
            '�������� ���������� ����� - � �������� ������� (������ ������)
                If Len(Trim(txtReservation.Text)) > InStr(intPosNum + 1, _
                Trim(txtReservation.Text), "/") Then
                    If Len(Trim(txtReservation.Text)) = InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 2 And _
                    Mid(Trim(txtReservation.Text), InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 1, 1) = 0 Then _
                    gTablePerson.Text = Left(gTablePerson.Text, InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/")) + Mid(gTablePerson, InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 2)
                End If
                
                If InStr(intPosNum + 1, Trim(txtReservation.Text), "/") > intPosNum + 1 Then
                    If Mid(gTablePerson.Text, intPosNum + 1, 1) = 0 _
                    And InStr(intPosNum + 1, Trim(txtReservation.Text), "/") = intPosNum + 3 Then _
                    gTablePerson.Text = Left(gTablePerson.Text, intPosNum) + _
                    Mid(gTablePerson, intPosNum + 2)
                End If
                
                If intPosNum > 1 Then
                    If intPosNum = 3 And Left(gTablePerson.Text, 1) = 0 Then _
                    gTablePerson.Text = Mid(gTablePerson.Text, 2)
                End If
                
            '���������� �������  ��������� ��������� � "������� ������"
                gChangesTablePerson = True
            '���������� ����� �� ������ "Save"
                cmdSave.SetFocus
                Exit Sub
            End If
            '������ "������� �������", "������� ����������" ���
            '  "������� ��������������� ���������" � ������������ ��������� �������
TimeTerCalError:
            frmDemo.BeepSound
            '�� ����������� ����� "Special"
        Else
                    '����� ��������������� ������� "������� ������"
            gTablePerson.Col = intColNumCorr
            If Len(Trim(txtReservation.Text)) < 9 Then
            '��������� ���� "Reservation" � "������� ������"
                gTablePerson.Text = Trim(txtReservation.Text)
            '���������� �������  ��������� ��������� � "������� ������"
                gChangesTablePerson = True
            '���������� ����� �� ������ "Save"
                cmdSave.SetFocus
            '�������� ������ ������
            Else
                frmDemo.BeepSound
            End If
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Reservation - Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            '����� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo AddressError
            '����� � ���������� ��������� ������� (01/15,  00 - ��������� �����)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text > 0 And txtAddress.Text < 16 Then
            '��������� ������ "Reservation" � "������� ������"
            If Len(Trim(gTablePerson.Text)) < 4 Then
                txtPort.Text = "2"
                txtType.Text = "CONTR"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Trim(txtAddress.Text) + Mid(gTablePerson.Text, 3)
            End If
            '���������� �������  ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            '���������� ����� �� ��������� ���� "Port"
            txtPort.SetFocus
            Exit Sub
            '������ � ������������ ��������� �������
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            '��������� ����� � ������� ��������������� "Reservation - Port"
Private Sub txtPort_KeyPress(KeyAscii As Integer)
            '����� ����� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo PortError
            '����� ����� � ���������� ��������� (2/8)
        If Len(Trim(txtPort.Text)) = 1 And txtPort.Text > 1 And txtPort.Text < 9 Then
            '��������� ������ "Reservation" � "������� ������"
            If Len(Trim(gTablePerson.Text)) < 4 Then
                txtAddress.Text = "01"
                txtType.Text = "CONTR"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Left(gTablePerson.Text, 2) + Trim(txtPort.Text) + _
                Mid(gTablePerson.Text, 4)
            End If
            '���������� �������  ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            '���������� ����� �� ��������� ���� "Type"
            txtType.SetFocus
            Exit Sub
            '����� ����� � ������������ ���������
PortError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Reservation - Type"
Private Sub txtType_KeyPress(KeyAscii As Integer)
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo TypeError
            '��� � ���������� ��������� ����� ("XXXXX")
        If Len(Trim(txtType.Text)) <= 5 Then
            '��������� ������ "Reservation" � "������� ������"
            If Len(Trim(gTablePerson.Text)) = 0 Then
                txtAddress.Text = "01"
                txtPort.Text = "2"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(txtType.Text)
            End If
            '���������� �������  ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            '���������� ����� �� ������ "Save"
            cmdSave.SetFocus
            Exit Sub
            '��� � ������������ ��������� �����
TypeError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '���������� ������ � "������� ������"
Private Sub cmdAdd_Click()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '��� � ������������ ��� � "������� ������"
Dim strName As String
Dim strPersonCode As String

    strName = ""
    strPersonCode = ""
    
            '������� ������������ �������� ����������
            '  ���������� "������� ������"
    fraColName.Enabled = False
    txtName.Enabled = False
    lstName.Enabled = False
    txtName.Enabled = False
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    txtReservation.Enabled = False
            '�������� ������ ����
    lstName.Clear
    
            
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ ��� �������
    strName = InputBox("Name: 1 -- 16 Characters !!!", "Add ...")
    If Len(Trim(strName)) > 16 Then strName = Left(Trim(strName), 16)
    frmDemo.BeepSound
            '�������� �� ������������ ������������ ���
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Add ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strPersonCode)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            '��� ��� ������������ ��� �� �������
    If strName = "" Or strPersonCode = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The Name Or PersonCode isn't selected"
            
            '��� � ������������ ��� �������
    Else
        '������� ������� "������� ������" = 1 (������������ ���)
        gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
            If Trim(gTablePerson.Text) = strPersonCode Then
            '��������� ����� �� �����
                Exit For
            End If
        Next
            '��������� ������������ ��� ��� ���� � "������� ������"
        If intRowNum < gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            MsgBox ("Duplicated PersonCode")
            '���������� ������������� ���� ��� � "������� ������"
        Else
            '���������� ������ � ����� "������� ������"
            gTablePerson.AddItem strPersonCode
            gTablePerson.Row = gTablePerson.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ������"
            gTablePerson.Col = 0
            gTablePerson.Text = Trim(strName)
            '��������� ������ "PersonCode" � "������� ������"
            gTablePerson.Col = 1
            gTablePerson.Text = Trim(strPersonCode)
            '������������ ������� ��� ��������� �������
            gTablePerson.Col = 3
            gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
            "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            '���������� ��������/���������� ����� � "������� ������"
            gTablePerson.Tag = gTablePerson.Tag + 1
            '���������� ������� ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            '  ��������������� ������ �������
            gProtocol.strProtocName = strName
            gProtocol.strProtocPersonCode = strPersonCode
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "PersonCode Addition"
            '�������� ������ � ���� "������� ���������"
            frmDemo.WriteProtocol
        End If
    End If
    
            '���������� ����� �� ������ "Add"
    If frmTablePerson.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            '����� ������ � "������� ������"
Private Sub cmdFind_Click()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '��� � ������������ ��� � "������� ������"
Dim strName As String
Dim strPersonCode As String

    strName = ""
    strPersonCode = ""
    
            '������� ������������ �������� ����������
            '  ���������� "������� ������"
    fraColName.Enabled = False
    txtName.Enabled = False
    lstName.Enabled = False
    txtName.Enabled = False
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    txtReservation.Enabled = False
            '�������� ������ ����
    lstName.Clear
    
            '�������� �� ������������ ��� �������
    strName = InputBox("Name: 1 -- 16 Characters !!!", "Find ...")
    If Len(Trim(strName)) > 16 Then strName = Left(Trim(strName), 16)
    frmDemo.BeepSound
            '�������� �� ������������ ������������ ���
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Find ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strPersonCode)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            '��� ��� ������������ ��� �� �������
    If strName = "" Or strPersonCode = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The Name Or PersonCode isn't selected"
            
            '��� � ������������ ��� �������
    Else
            
            '������� ������� "������� ������" = 1 (������������ ���)
        gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
            If Trim(gTablePerson.Text) = strPersonCode Then
            '������� ������� "������� ������" = 0 (���)
                gTablePerson.Col = 0
            '���������� ������������ ��� ���� � "������� ������"
                If InStr(1, Trim(gTablePerson.Text), strName) <> 0 Then
            '���������� ������ ��������� ���� "������� ������"
                    txtPersonCode.Text = gTablePerson.Text
                    gTablePerson.Col = 0
                    txtName.Text = gTablePerson.Text
                    gTablePerson.Col = 5
                    txtReservation.Text = gTablePerson.Text
            '��������� ����� �� �����
                    Exit For
                End If
            End If
        Next
            '����� ��� ������������� ���� ��� � "������� ������"
        If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            MsgBox ("Unexistent Name Or PersonCode")
            '���������� ����� �� ������ "Correction"
            If frmTablePerson.Visible = True Then cmdCorrection.SetFocus
            
            '��� � ������������ ��� ���� � "������� ������"
        Else
            '������� ���������� ��������� �������� ����������
            '   ���������� "������� ������"
            fraColName.Enabled = True
            txtName.Enabled = True
            '�������� ��������� ����
            txtName.Text = ""
            txtPersonCode.Text = ""
            txtReservation.Text = ""
            txtAddress.Text = ""
            txtPort.Text = ""
            txtType.Text = ""
            '�������� ������ ����
            lstName.Clear
            '�������� �����
            optName_Click
        End If
    
    End If
    
End Sub

            
            '��������������� ������������� ���� � "������� ������" ��� �����������
            '   ��� ��������: 0 - ��������������� ��������� �������;
            '                 1 - � ��������������� ��������.
Public Function AutoRegParking(ByVal vntPersonCode As Variant, _
ByVal strName As String, ByVal strStatus As String, ByVal strReserve As String, _
ByVal strTime As String)
            '����� ������� ������ � "������� ������"
Dim intRowNum As Integer
            '����� ������� �������� "/" � ������������� ����
Dim intPosNum As Integer
            '������ ���������� ��������� - ����
Dim intHourStart As Integer
            '������ ���������� ��������� - ������
Dim intMinuteStart As Integer
            '����� ���������� ��������� - ����
Dim intHourFinish As Integer
            '����� ���������� ��������� - ������
Dim intMinuteFinish As Integer
            '������ ������ "������� ������"
Dim strPerson As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ��������������� ��������
    AutoRegParking = 1
    
            '��������� ������������ ��� ��� ����������
            '  ��� ���� � "������� ������"
    If intRowNum < gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode or Info")
        Else
            MsgBox ("Person. kods vai Info jau ir")
        End If
            '���������� ������������� ���� � ����������
            '  ��� � "������� ������"
    Else
            '���������� ������ � ����� "������� ������"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ������"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            '��������� ������ "Status" � "������� ������"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            '��������� ������ "Reserve" � "������� ������"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            '������������ �������� ������ "Calendar" � "������� ������"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultParkCale
            '������������ ������� ��� ��������� �������
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
        "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            '��������� �������� ������� � ����������� (������ ��� ���������������)
        If Trim(strTime) = "Day" Then
            strTime = Trim(gParkingTimeD)
        ElseIf Trim(strTime) = "DayNight" Then
            strTime = Trim(gDefaultParkTime)
        ElseIf Trim(strTime) = "Night" Then
            strTime = Mid(Trim(gParkingTimeD), 7) + "/" + Left(Trim(gParkingTimeD), 5)
        End If
            
            '������ ���������� ��������� - ����
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            '������ ���������� ��������� - ������
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ����
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ������
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            '��������� ������ "Time" � "������� ������" - ����
        If intHourStart < 10 Then
            gTablePerson.Text = "0" + Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        Else
            gTablePerson.Text = Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        End If
        If intHourFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 6) + "0" + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 6) + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        End If
            '��������� ������ "Time" � "������� ������" - ������
        If intMinuteStart < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 3) + "0" + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        End If
        If intMinuteFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 9) + "0" + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 9) + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        End If
            
            '������ �������� ���������
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultParkCale + _
        Chr(7) + strReserve
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
    
            '���������� ��������/���������� ����� � "������� ������"
        gTablePerson.Tag = gTablePerson.Tag + 1
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            '��������������� ��������� �������
        AutoRegParking = 0
    
    End If
    
End Function
            
            '��������� ������ "Reserve" � "������� ������"
Public Function InputParking(intIndex As Integer)
            '���� "Name" � "������� ������"
Dim strName  As String
            '������
Dim strStatus As String
            '��������� "��������"
Dim strCheckingInfo As String * 8
            '��������� "����������� ���� � �����"
Dim strDateInfo As String
            '�����
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) �����
Dim strHour As String
Dim strMinute As String
            '������� � ������
Dim intPosNum As Integer
            '������� ����������
Dim intCount As Integer
            '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
Dim intCellLimit As Integer

            '���������� ������� ��������� ����������� ������� �����������
            '  (���������� ������) ����� ����������� ������� �����������
    If frmDemo.imgParkingInData(intIndex).Tag = 1 Then
            '����� ������ � ����� "������� ������" (��������� �����������)
        gTablePerson.Row = gTablePerson.Rows - 1
            '����� ������� � "������� ������" ("Reserve")
        gTablePerson.Col = 5
            '���������� �������� - ���������� ���������������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
            '��������� ������ "Reserve" � "������� ������" (���������� ������)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputParking = 0
            '������������ ��������
        Else
            InputParking = 1
        End If
            '������� ��������� ����������� ������� ����������� �� ���������� -
            '  (���������� ������) �� ����� ����������� ��� ����������� �������
    Else
            '����� ������� � ������� ������ "������� ������" ("Reserve")
        gTablePerson.Col = 5
            '���������� �������� - ���������� ��������������� ��� ����� �������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "1" And _
        Mid(Trim(gTablePerson.Text), 8, 1) <> "E" Then
            '��������� ������ "Reserve" � "������� ������" (���������� ������)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputParking = 0
            
            '����������� � ������������ ������� ������������ ����������
            If gParkTimeLimit > 0 Then
            '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
                intCellLimit = gParkingCellLimit
            '������� ������� "������� ������" = 2 (������)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            '���������� ������
                If Left(Trim(strStatus), 2) = "05" Then
            '������� ��������� "��������"
                    strCheckingInfo = ""
            
            '����� ������ �������
                    strDateInfo = Format(Now, "h:mm:ss")
            '����
                    intHour = Hour(strDateInfo)
            '������
                    intMinute = Minute(strDateInfo)
            '���� ������ �������
                    strDateInfo = Format(Now, "dd/mm/yyyy")
                    strDateInfo = Left(Trim(strDateInfo), 2) + Mid(Trim(strDateInfo), 4, 2) + _
                    Right(Trim(strDateInfo), 4)
                        
            
            '��������� "���������" ����� � ���� ������ �����������
            '  �������, �� ������� ��� ����� �������� ���������� �����
        
            '��������� ������� ����
                    If (intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell) > 59 Then
                        If (gParkingTimeCell * intCellLimit + gParkingTimeCell) > 1440 Then
                            intHour = intHour + Int((intMinute + _
                            gParkingTimeCell * intCellLimit) / 60)
                            intMinute = intMinute + gParkingTimeCell * intCellLimit - _
                            Int((intMinute + gParkingTimeCell * intCellLimit) / 60) * 60
                        Else
                            intHour = intHour + Int((intMinute + gParkingTimeCell + _
                            gParkingTimeCell * intCellLimit) / 60)
                            intMinute = intMinute + gParkingTimeCell + _
                            gParkingTimeCell * intCellLimit - _
                            Int((intMinute + gParkingTimeCell + _
                            gParkingTimeCell * intCellLimit) / 60) * 60
                        End If
            
            '��������� ������� ����
                        If intHour >= 24 Then
                            intHour = intHour - 24
            '��������� "���������" �� ����, ��������� �� �������
                            frmTableCalendar.comCalendar.Today
                            frmTableCalendar.comCalendar.NextDay
            '���������  �����
                            If frmTableCalendar.comCalendar.Day > 9 Then
                                strDateInfo = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            Else
                                strDateInfo = "0" + _
                                Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            End If
            '���������  ������ �, ��������, ����
                            If frmTableCalendar.comCalendar.Day = 1 Then
                                If frmTableCalendar.comCalendar.Month > 9 Then
                                    strDateInfo = "01" + _
                                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                                    Right(strDateInfo, 4)
                                Else
                                    strDateInfo = "010" + _
                                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                                    Right(strDateInfo, 4)
                                End If
                            End If
                        End If
            
            '�� ��������� ������� ����
                    Else
                        intMinute = intMinute + gParkingTimeCell * intCellLimit + _
                        gParkingTimeCell
                    End If
            
            '"���������" ����
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '"���������" ������
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
    
            '������������ ����������� ��������� "��������"
                    For intCount = 1 To 7 Step 2
            '����
                        strCheckingInfo = Trim(strCheckingInfo) + _
                        Chr(CByte(CInt(Mid(strDateInfo, intCount, 2))))
                    Next
            '����
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            '������
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            '�������� ��������� "��������"
                    Call frmTablePerson.Pack(strCheckingInfo)
            
            '��������� ���� "Name" � "������� ������"
                    gTablePerson.Col = 0
                    gTablePerson.Text = Left(strCheckingInfo, 6) + _
                    Mid(Trim(gTablePerson.Text), 7)
                    
                End If
            End If
            
            '������������ ��������
        Else
            InputParking = 1
        End If
    End If
            'K��������� ��������
    If InputParking = 0 Then
            
            '������ �������� ���������
        strMessage = "Cor "
        gTablePerson.Col = 0
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 1
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 2
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 5
        strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '��������� ������������� ���� �� "������� ������" ��� �����������
            '   ��� ��������: 0 - ��������� �������� �������;
            '                 1 - ��������� ����������.
Public Function AutoFindParking(ByVal vntPersonCode As Variant, strProtocName As String, _
                                                strProtocStatus As String, strChecking As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strProtocStatus = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 5 (������)
            gTablePerson.Col = 5
            strChecking = Trim(gTablePerson.Text)
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '��������� ����������
    AutoFindParking = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            '���������� ������������ ��� ���� � "������� ������"
    Else
            '��������� �������� �������
        AutoFindParking = 0
    End If
    
End Function
            
            '������������ (����������) ������������� ���� �����������
            '  �� "������� ������"
            '  ��� ��������: 0 - ������������ ��������� �������;
            '                1 - � ������������ ��������.
Public Function AutoDelParking(ByVal vntPersonCode As Variant, strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ������������ ��������
    AutoDelParking = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            '���������� ������������ ��� ���� � "������� ������"
    Else
            
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "��"
        If strResponse = vbYes Then
'������������ ����� ������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '����������� ��������
            If gParkingDeletion = 1 Then
            
            '������� ������� "������� ������" = 1 (������������ ���)
                gTablePerson.Col = 1
            
            '������ �������� ���������
                strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
                gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
                gRealDelPerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                Call frmDemo.SendMessage(strMessage)
            
            '���������� �� ������/������ ��������
            Else
            '������� ������� "������� ������" = 5 (������)
                gTablePerson.Col = 5
            '������ ����� ��� ������ - ������������� ��������
            ' (����������� ��������)
                If Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
            '������� ������� "������� ������" = 1 (������������ ���)
                    gTablePerson.Col = 1
            
            '������ �������� ���������
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
                    gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
                    gRealDelPerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            ' ���������� ��������
                Else
                    gTablePerson.Text = Left(Trim(gTablePerson.Text), 7) + "E"
            
            '������ �������� ���������
                    strMessage = "Cor "
                    gTablePerson.Col = 0
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 1
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 2
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 3
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 4
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 5
                    strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
            
            '������������ ��������� �������
            AutoDelParking = 0
            
            '���������� ������� ��������� ��������� � "������� ������"
            gChangesTablePerson = True
        End If
    End If
    
End Function
            
            '��������� ������ "Reserve" ��� ���������� ������ � "������� ������"
Public Function OutputParking(intIndex As Integer, intStatusCode As Integer)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            
            '����� ������� � ������� ������ "������� ������" ("Reserve")
    gTablePerson.Col = 5
            '���������� ������� ������ ���������� ���������� ������� ���
            '  �������������� (��� ��������) ������ ����������� �������
    If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Or intStatusCode = 6 Then
            '����� ������� � ������� ������ "������� ������" ("Person")
        gTablePerson.Col = 1
            
            '������ �������� ���������
        strMessage = "Del " + Trim(gTablePerson.Text)
            '��������� ������� ������ �� "������� ������"
        gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
        gRealDelPerson = True
            
            '���������� �������� - ������ ��������� �� "������� ������"
        OutputParking = 0
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������� ��������� �������� ������� �� ���������� -
            '  (���������� ������) �� ����� ����������� ��� ����������� �������
    Else
            '���������� �������� - ������ ��������������� ��� ����� �������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
            '��������� ������ "Reserve" � "������� ������" (������ ������)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "1" + _
            Right(Trim(gTablePerson.Text), 1)
            
            '������ �������� ���������
            strMessage = "Cor "
            gTablePerson.Col = 0
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 1
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 2
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 3
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 4
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 5
            strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call frmDemo.SendMessage(strMessage)
            
            '���������� ��������
            OutputParking = 0
            
            '������������ ��������
        Else
            OutputParking = 1
        End If
    End If
            
            'K��������� ��������
    If OutputParking = 0 Then
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '������ "������� ������" ��� ����������� (�������� ����� �
            '  ����������� �� ������������ ��������� �� � ��������� ����
            '  ��������:     "0" -  ������ ��������� �������;
            '                "1" -  �� ����������� ������������ AM, ������� ������
            '                       ���� ������������ ����a�� ����� ������ ��������;
            '                       ������ ��������;
            '                "2" -  ������ ����������.
Public Function AutoPresParking()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '��������� ���� �������� �������
    AutoPresParking = 0
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ��������� ����������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            '��������� ���� �������� �������
        AutoPresParking = 2
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� ������� �����������
            If Left(Trim(gTablePerson.Text), 2) = "07" Or _
            Left(Trim(gTablePerson.Text), 2) = "05" Or _
            Left(Trim(gTablePerson.Text), 2) = "06" Then
            '������� - "Reserve"
                gTablePerson.Col = 5
            '������� ��, ������� ������ ����
            '  ������������ ������� ����� ������
                If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Then
            '����� ������� � ������� ������ "������� ������" ("Person")
                    gTablePerson.Col = 1
            '������������ ��, ������� ������ ����
            '  ������������ ������� ����� ������, �� �� �������
                    If gTablePerson.Text <> "Deleted" Then AutoPresParking = 1
                End If
            End If
        Next
    End If

End Function
            
            '������������� ���������� ����� "������� ������" ��� �����������
            '   ��� ��������: 0 - ������������� ��������� �������;
            '                 1 - ������������� ����������.
Public Function AutoCorParking(ByVal vntPersonCode As Variant, ByVal strName _
As String, ByVal strChecking As String, ByRef strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            gTablePerson.Text = strName
            '������� ������� "������� ������" = 5 (������)
            gTablePerson.Col = 5
            gTablePerson.Text = strChecking
            
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '������������� ����������
    AutoCorParking = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        MsgBox "Correction impossible !"
            '���������� ������������ ��� ���� � "������� ������"
    Else
            
            '������ �������� ���������
        strMessage = "Cor " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7)
            '������� ������� "������� ������" = 3 (�����)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            '������� ������� "������� ������" = 4 (����)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7) + strChecking
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������������� ��������
        AutoCorParking = 0
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '��������������� ������������� ���� � "������� ������" ��� �����������
            '   ��� ��������: 0 - ��������������� ��������� �������;
            '                 1 - � ��������������� ��������.
Public Function AutoRegAccess(ByVal vntPersonCode As Variant, _
ByVal strName As String, ByVal strStatus As String, ByVal strReserve As String, _
ByVal strTime As String)
            '����� ������� ������ � "������� ������"
Dim intRowNum As Integer
            '����� ������� �������� "/" � ������������� ����
Dim intPosNum As Integer
            '������ ���������� ��������� - ����
Dim intHourStart As Integer
            '������ ���������� ��������� - ������
Dim intMinuteStart As Integer
            '����� ���������� ��������� - ����
Dim intHourFinish As Integer
            '����� ���������� ��������� - ������
Dim intMinuteFinish As Integer
            '������ ������ "������� ������"
Dim strPerson As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ��������������� ��������
    AutoRegAccess = 1
    
            '��������� ������������ ��� ��� ����������
            '  ��� ���� � "������� ������"
    If intRowNum < gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode")
        Else
            MsgBox ("Person. kods jau ir")
        End If
            '���������� ������������� ���� � ����������
            '  ��� � "������� ������"
    Else
            '���������� ������ � ����� "������� ������"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ������"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            '��������� ������ "Status" � "������� ������"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            '��������� ������ "Reserve" � "������� ������"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            '������������ �������� ������ "Calendar" � "������� ������"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultAcceCale
            '������������ ������� ��� ��������� �������
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
        "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            '��������� �������� ������� (������ ��� ���������������)
        If Trim(strTime) = "Day" Then
            strTime = Trim(gAccessTimeD)
        ElseIf Trim(strTime) = "DayNight" Then
            strTime = Trim(gDefaultAcceTime)
        ElseIf Trim(strTime) = "Night" Then
            strTime = Mid(Trim(gAccessTimeD), 7) + "/" + Left(Trim(gAccessTimeD), 5)
        End If
            
            '������ ���������� ��������� - ����
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            '������ ���������� ��������� - ������
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ����
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ������
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            '��������� ������ "Time" � "������� ������" - ����
        If intHourStart < 10 Then
            gTablePerson.Text = "0" + Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        Else
            gTablePerson.Text = Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        End If
        If intHourFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 6) + "0" + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 6) + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        End If
            '��������� ������ "Time" � "������� ������" - ������
        If intMinuteStart < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 3) + "0" + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        End If
        If intMinuteFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 9) + "0" + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 9) + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        End If
            
            '������ �������� ���������
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultAcceCale + _
        Chr(7) + strReserve
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ��������/���������� ����� � "������� ������"
        gTablePerson.Tag = gTablePerson.Tag + 1
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            '��������������� ��������� �������
        AutoRegAccess = 0
        
    End If
    
End Function
            
            '��������� ������ "Reserve" � "������� ������"
Public Function InputAccess(intIndex As Integer)
            '���� "Name" � "������� ������"
Dim strName  As String
            '������
Dim strStatus As String
            '��������� "��������"
Dim strCheckingInfo As String * 8
            '��������� "����������� ���� � �����"
Dim strDateInfo As String
            '�����
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) �����
Dim strHour As String
Dim strMinute As String
            '������� � ������
Dim intPosNum As Integer
            '������� ����������
Dim intCount As Integer
            '���������� ����� �������, � ������� �������� �����������
            '  ����������� ������� ���������� ���������� �� �����������
Dim intCellLimit As Integer


            '���������� ������� ��������� ����������� �������
            '  (���������� �����) ����� ����������� �������
    If frmDemo.imgAccessInData(intIndex).Tag = 1 Then
            '����� ������ � ����� "������� ������" (��������� �����������)
        gTablePerson.Row = gTablePerson.Rows - 1
            '����� ������� � "������� ������" ("Reserve")
        gTablePerson.Col = 5
            '���������� �������� - ���������� ���������������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
            '��������� ������ "Reserve" � "������� ������" (���������� �����)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputAccess = 0
            '������������ ��������
        Else
            InputAccess = 1
        End If
            '������� ��������� ����������� ������� �� ���������� -
            '  (���������� �����) �� ����� ����������� ��� ����������� �������
    Else
            '����� ������� � ������� ������ "������� ������" ("Reserve")
        gTablePerson.Col = 5
            '���������� �������� - ���������� ��������������� ��� ����� �������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "1" And _
        Mid(Trim(gTablePerson.Text), 8, 1) <> "E" Then
            '��������� ������ "Reserve" � "������� ������" (���������� �����)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputAccess = 0
            
            '����������� � ������������ ������� ������������ ����������
            If gAcceTimeLimit > 0 Then
            '���������� ����� �������, � ������� �������� �����������
            '  ����������� ������� ���������� ���������� �� �����������
                intCellLimit = gAccessCellLimit
            '������� ������� "������� ������" = 2 (������)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            '���������� ������
                If Left(Trim(strStatus), 2) = "08" Then
            '������� ��������� "��������"
                    strCheckingInfo = ""
            
            '����� ����� �������
                    strDateInfo = Format(Now, "h:mm:ss")
            '����
                    intHour = Hour(strDateInfo)
            '������
                    intMinute = Minute(strDateInfo)
            '���� ����� �������
                    strDateInfo = Format(Now, "dd/mm/yyyy")
                    strDateInfo = Left(Trim(strDateInfo), 2) + Mid(Trim(strDateInfo), 4, 2) + _
                    Right(Trim(strDateInfo), 4)
                        
            
            '��������� "���������" ����� � ���� ����� �����������
            '  �������, �� ������� ��� ����� �������� ���������� �����
        
            '��������� ������� ����
                    If (intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell) > 59 Then
                        If (gAccessTimeCell * intCellLimit + gAccessTimeCell) > 1440 Then
                            intHour = intHour + Int((intMinute + _
                            gAccessTimeCell * intCellLimit) / 60)
                            intMinute = intMinute + gAccessTimeCell * intCellLimit - _
                            Int((intMinute + gAccessTimeCell * intCellLimit) / 60) * 60
                        Else
                            intHour = intHour + Int((intMinute + gAccessTimeCell + _
                            gAccessTimeCell * intCellLimit) / 60)
                            intMinute = intMinute + gAccessTimeCell + _
                            gAccessTimeCell * intCellLimit - _
                            Int((intMinute + gAccessTimeCell + _
                            gAccessTimeCell * intCellLimit) / 60) * 60
                        End If
            
            '��������� ������� ����
                        If intHour >= 24 Then
                            intHour = intHour - 24
            '��������� "���������" �� ����, ��������� �� �������
                            frmTableCalendar.comCalendar.Today
                            frmTableCalendar.comCalendar.NextDay
            '���������  �����
                            If frmTableCalendar.comCalendar.Day > 9 Then
                                strDateInfo = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            Else
                                strDateInfo = "0" + _
                                Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            End If
            '���������  ������ �, ��������, ����
                            If frmTableCalendar.comCalendar.Day = 1 Then
                                If frmTableCalendar.comCalendar.Month > 9 Then
                                    strDateInfo = "01" + _
                                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                                    Right(strDateInfo, 4)
                                Else
                                    strDateInfo = "010" + _
                                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                                    Right(strDateInfo, 4)
                                End If
                            End If
                        End If
            
            '�� ��������� ������� ����
                    Else
                        intMinute = intMinute + gAccessTimeCell * intCellLimit + _
                        gAccessTimeCell
                    End If
            
            '"���������" ����
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '"���������" ������
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
    
            '������������ ����������� ��������� "��������"
                    For intCount = 1 To 7 Step 2
            '����
                        strCheckingInfo = Trim(strCheckingInfo) + _
                        Chr(CByte(CInt(Mid(strDateInfo, intCount, 2))))
                    Next
            '����
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            '������
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            '�������� ��������� "��������"
                    Call frmTablePerson.Pack(strCheckingInfo)
            
            '��������� ���� "Name" � "������� ������"
                    gTablePerson.Col = 0
                    gTablePerson.Text = Left(strCheckingInfo, 6) + _
                    Mid(Trim(gTablePerson.Text), 7)
                    
                End If
            End If
            
            '������������ ��������
        Else
            InputAccess = 1
        End If
    End If
            'K��������� ��������
    If InputAccess = 0 Then
            
            '������ �������� ���������
        strMessage = "Cor "
        gTablePerson.Col = 0
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 1
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 2
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 5
        strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '��������� ������������� ���� �� "������� ������" ��� �����������
            '   ��� ��������: 0 - ��������� �������� �������;
            '                 1 - ��������� ����������.
Public Function AutoFindAccess(ByVal vntPersonCode As Variant, strProtocName As String, _
                                                strProtocStatus As String, strChecking As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strProtocStatus = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 5 (������)
            gTablePerson.Col = 5
            strChecking = Trim(gTablePerson.Text)
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '��������� ����������
    AutoFindAccess = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            '���������� ������������ ��� ���� � "������� ������"
    Else
            '��������� �������� �������
        AutoFindAccess = 0
    End If
    
End Function
            
            '������������ (����������) ������������� ���� ����������
            '  �� "������� ������"
            '  ��� ��������: 0 - ������������ ��������� �������;
            '                1 - � ������������ ��������.
Public Function AutoDelAccess(ByVal vntPersonCode As Variant, strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ������������ ��������
    AutoDelAccess = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            '���������� ������������ ��� ���� � "������� ������"
    Else
            
            '��������� ���������� � ������� ���������
        If frmLease.Tag <> "Exit" Then _
        '������� ������� "������� ������" = 0 (����������)
            gTablePerson.Col = 0
            gTablePerson.Text = Trim(frmDataAccessOut.txtInfo.Text)
        End If
            
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "��"
        If strResponse = vbYes Then
'������������ ����� ������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '����������� ��������
            If gAccessDeletion = 1 Then
        '������� ������� "������� ������" = 1 (������������ ���)
                gTablePerson.Col = 1
            
            '������ �������� ���������
                strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
                gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
                gRealDelPerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                Call frmDemo.SendMessage(strMessage)
            
            '���������� �� ������/������ ��������
            Else
            '������� ������� "������� ������" = 5 (������)
                gTablePerson.Col = 5
            '���������� ����� ��� ����� - ������������� ��������
            ' (����������� ��������)
                If Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
            '������� ������� "������� ������" = 1 (������������ ���)
                    gTablePerson.Col = 1
            
            '������ �������� ���������
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
                    gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
                    gRealDelPerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            ' ���������� ��������
                Else
                    gTablePerson.Text = Left(Trim(gTablePerson.Text), 7) + "E"
            
            '������ �������� ���������
                    strMessage = "Cor "
                    gTablePerson.Col = 0
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 1
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 2
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 3
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 4
                    strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
                    gTablePerson.Col = 5
                    strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
            
            '������������ ��������� �������
            AutoDelAccess = 0
            '���������� ������� ��������� ��������� � "������� ������"
            gChangesTablePerson = True
        End If
    End If
    
End Function
            
            '��������� ������ "Reserve" ��� ���������� ������ � "������� ������"
Public Function OutputAccess(intIndex As Integer, intStatusCode As Integer)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
            '����� ������� � ������� ������ "������� ������" ("Reserve")
    gTablePerson.Col = 5
            '���������� ������� ������ ���������� ���������� ���
            '  �������������� (��� ��������) ������ ����������� ����������
    If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Or intStatusCode = 9 Then
            '����� ������� � ������� ������ "������� ������" ("Person")
        gTablePerson.Col = 1
            
            '������ �������� ���������
        strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
        gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
        gRealDelPerson = True
            '���������� �������� - ���������� ��������� �� "������� ������"
        OutputAccess = 0
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������� ��������� �������� ������� �� ���������� -
            '  (���������� �����) �� ����� ����������� ��� ����������� �������
    Else
            '���������� �������� - ���������� ��������������� ��� ����� ������
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
            '��������� ������ "Reserve" � "������� ������" (���������� �����)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "1" + _
            Right(Trim(gTablePerson.Text), 1)
            
            '������ �������� ���������
            strMessage = "Cor "
            gTablePerson.Col = 0
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 1
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 2
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 3
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 4
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 5
            strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call frmDemo.SendMessage(strMessage)
            
            'K��������� ��������
            OutputAccess = 0
            '������������ ��������
        Else
            OutputAccess = 1
        End If
    End If
            
            'K��������� ��������
    If OutputAccess = 0 Then
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '������ "������� ������" ��� ����������� (�������� ����� �
            '  ����������� �� ������������ �������� ����������� � ��������� ����
            '  ��������:     "0" -  ������ ��������� �������;
            '                "1" -  ������������ ����������, ������� ������
            '                       ���� ������������ ������ ����� ������;
            '                "2" -  ������ ����������.
Public Function AutoPresAccess()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '��������� ���� �������� �������
    AutoPresAccess = 0
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ��������� ����������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            '��������� ���� �������� �������
        AutoPresAccess = 2
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� �������
            If Left(Trim(gTablePerson.Text), 2) = "10" Or _
            Left(Trim(gTablePerson.Text), 2) = "08" Or _
            Left(Trim(gTablePerson.Text), 2) = "09" Then
            '������� - "Reserve"
                gTablePerson.Col = 5
            '������� ����������, ������� ������ ���
            '  ������������ ����� ����� ������
                If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Then
            '����� ������� � ������� ������ "������� ������" ("Person")
                    gTablePerson.Col = 1
            '������������ ����������, ������� ������ ���
            '  ������������ ����� ����� ������, �� �� �����
                    If gTablePerson.Text <> "Deleted" Then AutoPresAccess = 1
                End If
            End If
        Next
    End If

End Function
            
            '������������� ���������� ����� "������� ������"
            '   ��� ��������: 0 - ������������� ��������� �������;
            '                 1 - ������������� ����������.
Public Function AutoCorAccess(ByVal vntPersonCode As Variant, ByVal strName _
As String, ByVal strChecking As String, ByRef strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            gTablePerson.Text = strName
            '������� ������� "������� ������" = 5 (������)
            gTablePerson.Col = 5
            gTablePerson.Text = strChecking
            
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '������������� ����������
    AutoCorAccess = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        MsgBox "Correction impossible !"
            '���������� ������������ ��� ���� � "������� ������"
    Else
            
            '������ �������� ���������
        strMessage = "Cor " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7)
            '������� ������� "������� ������" = 3 (�����)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            '������� ������� "������� ������" = 4 (����)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7) + strChecking
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������������� ��������
        AutoCorAccess = 0
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    
    End If
    
End Function
            
            '������� "������� ������" ��� ����������� (�������� ����� �
            '  ����������� � ��������� ����������� � ��������� ����
            '  ��������:     "0" -  ������� ��������� �������;
            '                "1" -  ������ ����������.
Public Function CleaningAccess()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            
            '��������� ���� �������� �������
    CleaningAccess = 0
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� �������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            '��������� ���� �������� �������
        CleaningAccess = 1
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� ������� - ���������
            If Left(Trim(gTablePerson.Text), 2) = "09" Then
            '������� - "PersonCode"
                gTablePerson.Col = 1
            '������ ����������� ��������� ������� ��� ����������� ��������
                frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
                intButtonsAndIcons = vbYesNo + vbQuestion
                If frmDemo.optEnglish = True Then
                    strResponse = MsgBox("Deletion PersonCode = " + _
                    Trim(gTablePerson.Text), intButtonsAndIcons, "Cancel")
                Else
                    strResponse = MsgBox("Izslegt person. kods = " + _
                    Trim(gTablePerson.Text), intButtonsAndIcons, "Cancel")
                End If
            '������ ������ "��"
                If strResponse = vbYes Then
'������������ ����� ������� ������ "������� ������"
                    gTablePerson.Row = intRowNum
            
            '������ �������� ���������
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
                    gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
                    gRealDelPerson = True
            '���������� ������� ��������� ��������� � "������� ������"
                    gChangesTablePerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
        Next
    End If

End Function
            
            '��������������� ������������� ���� ��������� � "������� ������"
            '   ��� ��������: 0 - ��������������� ��������� �������;
            '                 1 - � ��������������� ��������.
Public Function AutoRegEmploye(ByVal vntPersonCode As Variant, _
ByVal strName As String)
            '����� ������� ������ � "������� ������"
Dim intRowNum As Integer
            '����� ������� �������� "/" � ������������� ����
Dim intPosNum As Integer
            '������ ���������� ��������� - ����
Dim intHourStart As Integer
            '������ ���������� ��������� - ������
Dim intMinuteStart As Integer
            '����� ���������� ��������� - ����
Dim intHourFinish As Integer
            '����� ���������� ��������� - ������
Dim intMinuteFinish As Integer
            '������ ������ "������� ������"
Dim strPerson As String
Dim strTime As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ��������������� ��������
    AutoRegEmploye = 1

            '��������� ������������ ��� ��� ���� � "������� ������"
    If intRowNum < gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode")
        Else
            MsgBox ("Person. kods jau ir")
        End If
            '���������� ������������� ���� ��� � "������� ������"
    Else
            '���������� ������ � ����� "������� ������"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ������"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            '��������� ������ "Status" � "������� ������"
        gTablePerson.Col = 2
        gTablePerson.Text = gDefaultStatus
            '������������ �������� ������ "Calendar" � "������� ������"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultCalendar
            '������������ ������� ��� ��������� �������
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
            "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            '��������� �������� ������� (������������)
        strTime = gDefaultTime
            
            '������ ���������� ��������� - ����
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            '������ ���������� ��������� - ������
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ����
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            '����� ���������� ��������� - ������
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            '��������� ������ "Time" � "������� ������" - ����
        If intHourStart < 10 Then
            gTablePerson.Text = "0" + Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        Else
            gTablePerson.Text = Trim(Str(intHourStart)) + Mid(gTablePerson.Text, 3)
        End If
        If intHourFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 6) + "0" + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 6) + Trim(Str(intHourFinish)) _
            + Mid(gTablePerson.Text, 9)
        End If
            '��������� ������ "Time" � "������� ������" - ������
        If intMinuteStart < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 3) + "0" + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(Str(intMinuteStart)) _
            + Mid(gTablePerson.Text, 6)
        End If
        If intMinuteFinish < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 9) + "0" + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 9) + Trim(Str(intMinuteFinish)) _
            + Mid(gTablePerson.Text, 12)
        End If
            
            '������ �������� ���������
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        gDefaultStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultCalendar + _
        Chr(7) + " "
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ��������/���������� ����� � "������� ������"
        gTablePerson.Tag = gTablePerson.Tag + 1
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            '��������������� ��������� �������
        AutoRegEmploye = 0
    
    End If
               
End Function
            
            '��������� ������ "Name" � "������� ������"
Public Function InputEmploye(intIndex As Integer)
            '����� ������� � ������� ������ "������� ������" ("Name")
    gTablePerson.Col = 0
            '���������� �������� - �������� ��������������� ��� ����� �������
    If Len(Trim(gTablePerson.Text)) < 16 Or _
    Right(Trim(gTablePerson.Text), 1) = "-" Then
            '��������� ������ "Name" � "������� ������" (C������� �����)
        If Len(Left(Trim(gTablePerson.Text), 15)) < 15 Then
            gTablePerson.Text = Trim(gTablePerson.Text) + _
            Left("              ", 15 - Len(Trim(gTablePerson.Text))) + "+"
        Else
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 15) + "+"
        End If
        InputEmploye = 0
            '������������ ��������
    Else
            InputEmploye = 1
    End If
            'K��������� ��������
    If InputEmploye = 0 Then
            
            '������ �������� ���������
        strMessage = "Cor " + Trim(gTablePerson.Text) + Chr(7)
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 2
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
        gTablePerson.Col = 5
        strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '��������� ������������� ���� ��� ����� ��������� � "������� ������"
            '   ��� ��������: 0 - ��������� �������� �������;
            '                 1 - ��������� ����������.
Public Function AutoFindEmploye(vntPersonCode As Variant, _
vntInfo As Variant, strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
    
            '��������� ����������
    AutoFindEmploye = 1
        
        '����� �� �������������� ����
    If vntPersonCode <> "" Then
        '������� ������� "������� ������" = 1 (������������ ���)
        gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
            If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
                gTablePerson.Col = 0
                frmDataEmployeOut.txtInfo = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 2 (������)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            '��������� �������� �������
                AutoFindEmploye = 0
            '��������� ����� �� �����
                Exit For
            End If
        Next
        
        '����� �� �����
    ElseIf vntInfo <> "" Then
        '������� ������� "������� ������" = 0 (���)
        gTablePerson.Col = 0
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '���������� ��� ���� � "������� ������"
            If InStr(1, Trim(gTablePerson.Text), Trim(vntInfo)) <> 0 Then
            '������� ������� "������� ������" = 1 (������������ ���)
                gTablePerson.Col = 1
                frmDataEmployeOut.txtPersonCode = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 2 (������)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            '��������� �������� �������
                AutoFindEmploye = 0
            '��������� ����� �� �����
                Exit For
            End If
        Next
    End If
    
            '����������� ������������� ���� ��� � "������� ������"
    If AutoFindEmploye = 1 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
    End If
    
End Function
            
            '������������ (����������) ������������� ���� ���������
            '  �� "������� ������"
            '  ��� ��������: 0 - ������������ ��������� �������;
            '                1 - � ������������ ��������.
Public Function AutoDelEmploye(vntPersonCode As Variant, strStatus As String)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
    
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '������� ������� "������� ������" = 2 (������)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            '��������� ����� �� �����
            Exit For
        End If
    Next
            '� ������������ ��������
    AutoDelEmploye = 1
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            '���������� ������������ ��� ���� � "������� ������"
    Else
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "��"
        If strResponse = vbYes Then
'������������ ����� ������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� ������� "������� ������" = 1 (������������ ���)
            gTablePerson.Col = 1
            
            '������ �������� ���������
            strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
            gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
            gRealDelPerson = True
            
            '������������ ��������� �������
            AutoDelEmploye = 0
            '���������� ������� ��������� ��������� � "������� ������"
            gChangesTablePerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        End If
    End If
    
End Function
            
            '��������� ������ "Name" � "������� ������"
Public Function OutputEmploye(intIndex As Integer)
            '����� ������� � ������� ������ "������� ������" ("Name")
    gTablePerson.Col = 0
            '���������� �������� - �������� ��������������� ��� ����� ������
    If Len(Trim(gTablePerson.Text)) < 16 Or _
    Right(Trim(gTablePerson.Text), 1) = "+" Then
            '����� �����
        If Left(Trim(gTablePerson.Text), 1) = gVisitor Then
            '����� ������� � ������� ������ "������� ������" ("Person")
            gTablePerson.Col = 1
            
            '������ �������� ���������
            strMessage = "Del " + Trim(gTablePerson.Text)
            
            '��������� ������� ������ �� "������� ������"
            gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
            gRealDelPerson = True
            
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        
            '����� �� �����
        Else
            '��������� ������ "Name" � "������� ������" (C������� �����)
            If Len(Left(Trim(gTablePerson.Text), 15)) < 15 Then
                gTablePerson.Text = Trim(gTablePerson.Text) + _
                Left("              ", 15 - Len(Trim(gTablePerson.Text))) + "-"
            Else
                gTablePerson.Text = Left(Trim(gTablePerson.Text), 15) + "-"
            End If
            
            '������ �������� ���������
            strMessage = "Cor "
            gTablePerson.Col = 0
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 1
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 2
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 3
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 4
            strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            gTablePerson.Col = 5
            strMessage = strMessage + Trim(gTablePerson.Text)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        End If
        
            'K��������� ��������
        OutputEmploye = 0
        
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
        
            '������������ ��������
    Else
        OutputEmploye = 1
    End If
    
End Function
            
            '������ "������� ������" ��� �������� (����� ������ � ��������� ����
            '  ��������:     "0" -  ������ ��������� �������;
            '                "1" -  ������������ �����;
            '                "2" -  ������ ����������.
Public Function AutoPresEmploye()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '��������� ���� �������� �������
    AutoPresEmploye = 0
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ��������� ����������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            '��������� ���� �������� �������
        AutoPresEmploye = 2
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� ���������
            If Left(Trim(gTablePerson.Text), 2) = "00" Or _
            Left(Trim(gTablePerson.Text), 2) = "01" Then
            '������� - "Name"
                gTablePerson.Col = 0
            '������������ ����� - ��������� ���� �������� �������
                If Left(Trim(gTablePerson.Text), 1) = gVisitor Then AutoPresEmploye = 1
            End If
        Next
    End If

End Function
            
            '���������� �������� ��������� ��������� ����� '������� ������"
Public Function RealDelPerson()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            
            '��� �������� ��� ��������� � �������� ������� "Mutex"
Dim lngRetCode As Long
            
            '��������� ������ �� ����� ���� ������� - ����� �� ���������
    If gTablePerson.Rows = 2 Then Exit Function
    
            '���� "������a ������" ��� ��� �������� ��� ������������ �����
    If gTablePerson.Access < 1 Then
            '����� ������������ ������� "Mutex"
        lngRetCode = WaitForSingleObject(gMutex, 15000)
            '������� ������� "������� ������" = 1 (������������ ���)
        gTablePerson.Col = 1
            '������� ������ "������� ������"
        gTablePerson.Row = 1
            
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������ ��������� �������
            If Trim(gTablePerson.Text) = "Deleted" Then
            '���������� �������� ������ �� "������� ������"
                gTablePerson.RemoveItem gTablePerson.Row
            '���������� ��������/���������� ����� � "������� ������"
                gTablePerson.Tag = gTablePerson.Tag - 1
            Else
            '������� ������ "������� ������"
                If gTablePerson.Row < gTablePerson.Rows - 1 Then _
                gTablePerson.Row = gTablePerson.Row + 1
            End If
        Next
            '���������� ������ "Mutex"
        lngRetCode = ReleaseMutex(gMutex)
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
            '����� ������ �� �������� �������� ����� �� "������� ������"
        gRealDelPerson = False
    End If
    
End Function
            
            '�������� ������ �� "������� ������"
Private Sub cmdDelete_Click()
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '������������ ��� � "������� ����������"
Dim strPersonCode As String
            '������� ��������� ��������
Dim intRealDel As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '������� ������������ �������� ����������
            '  ���������� "������� ������"
    fraColName.Enabled = False
    txtName.Enabled = False
    lstName.Enabled = False
    txtName.Enabled = False
    txtPersonCode.Enabled = False
    txtAddress.Enabled = False
    txtType.Enabled = False
    txtPort.Enabled = False
    fraStatus.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    fraCalendar.Enabled = False
    txtReservation.Enabled = False
            '�������� ��������� ����
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            '�������� ������ ����
    lstName.Clear
    
            '������� ����������� ��������
    intRealDel = 1
            
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ ������������ ���
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Delete ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strPersonCode)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            '������������ ��� �� ������
    If strPersonCode = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The PersonCode isn't selected"
            
            '������������ ��� ������
    Else
            '������� "PersonCode"
        gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
            If Trim(gTablePerson.Text) = strPersonCode Then
            '����� �������� �������
                intRealDel = 0
            '������ ����������� ��������� ������� ��� ����������� ��������
                frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
                intButtonsAndIcons = vbYesNo + vbQuestion
                strResponse = MsgBox("Deletion PersonCode ?", _
                intButtonsAndIcons, "Cancel")
            '������ ������ "��"
                If strResponse = vbYes Then
                    gTablePerson.Col = 0
                    gProtocol.strProtocName = Trim(gTablePerson.Text)
            '�������� ������
                    gTablePerson.RemoveItem intRowNum
            '���������� ��������/���������� ����� � "������� ������"
                    gTablePerson.Tag = gTablePerson.Tag - 1
            '���������� ������� ��������� ��������� � "������� ������"
                    gChangesTablePerson = True
            '  ��������������� ������ �������
                    gProtocol.strProtocPersonCode = strPersonCode
                    gProtocol.strProtocStatus = "04 - Manager"
            '�����
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "PersonCode Deletion"
            '�������� ������ � ���� "������� ���������"
                    frmDemo.WriteProtocol
                End If
            '��������� ����� �� �����
                Exit For
            End If
        Next
            
            '������������� ���� ��� � "������� ������"
        If intRealDel = 1 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            MsgBox ("Unexistent PersonCode")
        End If
    End If
    
            '���������� ����� �� ������ "Delete"
    If frmTablePerson.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            '���������� "������� ������" � ����� �� ���������
Public Function SaveTablePerson()
    Call cmdSave_Click
    SaveTablePerson = 0
    
End Function
            
            '���������� "������� ������" � ����� �� ���������
Private Sub cmdSave_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ������"
Dim intColNum As Integer
            
            '��������� ����� ������ (������) "������� ������"
    lngRecordLen = Len(gPerson)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TablePerson.dat"
    
            '�����, ��������� �� "������� ������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
    If gTablePerson.Tag < 0 Then
            '������� "������" ������������ ����
        Kill strPathFileName
    End If
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ������" � ����
            Select Case intColNum
                Case 0
                gPerson.strName = gTablePerson.Text
                Case 1
            'C�������� ���������� � ���������� ��������� ����:
            '  - �� �����������
                If Trim(gPerson.strName) = "ParkFreePlaces" Then
                    gPerson.strPersonCode = Left("000000000000000", 16 - _
                    Len(CStr(gParkFreePlaces))) + CStr(gParkFreePlaces)
                    gTablePerson.Text = gPerson.strPersonCode
            '  - �� �����������
                ElseIf Trim(gPerson.strName) = "AcceFreePlaces" Then
                    gPerson.strPersonCode = Left("000000000000000", 16 - _
                    Len(CStr(gAcceFreePlaces))) + CStr(gAcceFreePlaces)
                    gTablePerson.Text = gPerson.strPersonCode
                End If
                gPerson.strPersonCode = gTablePerson.Text
                Case 2
                gPerson.strStatus = Left(gTablePerson.Text, 2)
                Case 3
                gPerson.strTime = Left(gTablePerson.Text, 2) + Mid(gTablePerson.Text, 4, 2) + _
                Mid(gTablePerson.Text, 7, 2) + Mid(gTablePerson.Text, 10, 2)
                Case 4
                gPerson.strCalendar = Left(gTablePerson.Text, 2)
                Case 5
                gPerson.strReserve = gTablePerson.Text
            End Select
        Next
            '�������� ������ "������� ������" � ����
        Put intFileNum, intRowNum, gPerson
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� ������"
    gTablePerson.Tag = 0
            '�������� ������� ��������� ��������� � "������� ������"
    gChangesTablePerson = False
            '���������� ����� �� ������ "Cancel"
    If frmTablePerson.Visible = True Then cmdCancel.SetFocus
            
End Sub
            
            '���������� "������� ������" � ���������� �����
Private Sub cmdSaveAs_Click()
            '������ ��� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ������"
Dim intColNum As Integer

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
        frmDemo.BeepSound
        MsgBox "The file isn't selected !"
            '������ "������� ������" � ��������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� ������"
        lngRecordLen = Len(gPerson)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '�����, ��������� �� "������� ������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
        If gTablePerson.Tag < 0 Then
            '������� "������" ����, ���� �� ����������
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            '������� ��������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '�� ���� �������� "������� ������"
            For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
                gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ������" � ����
                Select Case intColNum
                    Case 0
                    gPerson.strName = gTablePerson.Text
                    Case 1
            '��������� ���������� � ���������� ��������� ����:
            '  - �� �����������
                    If Trim(gPerson.strName) = "ParkFreePlaces" Then
                        gPerson.strPersonCode = Left("000000000000000", 16 - _
                        Len(CStr(gParkFreePlaces))) + CStr(gParkFreePlaces)
                        gTablePerson.Text = gPerson.strPersonCode
            '  - �� �����������
                    ElseIf Trim(gPerson.strName) = "AcceFreePlaces" Then
                        gPerson.strPersonCode = Left("000000000000000", 16 - _
                        Len(CStr(gAcceFreePlaces))) + CStr(gAcceFreePlaces)
                        gTablePerson.Text = gPerson.strPersonCode
                    End If
                    gPerson.strPersonCode = gTablePerson.Text
                    Case 2
                    gPerson.strStatus = Left(gTablePerson.Text, 2)
                    Case 3
                    gPerson.strTime = Left(gTablePerson.Text, 2) + Mid(gTablePerson.Text, 4, 2) + _
                    Mid(gTablePerson.Text, 7, 2) + Mid(gTablePerson.Text, 10, 2)
                    Case 4
                    gPerson.strCalendar = Left(gTablePerson.Text, 2)
                    Case 5
                    gPerson.strReserve = gTablePerson.Text
                End Select
            Next
            '�������� ������ "������� ������" � ����
            Put intFileNum, intRowNum, gPerson
        Next
            '������� ��������� ����
        Close intFileNum
             '���������� ��������/���������� ����� � "������� ������"
        gTablePerson.Tag = 0
               '�������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = False
    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '���������� ����� �� ������ "Cancel"
    If frmTablePerson.Visible = True Then cmdCancel.SetFocus
    
End Sub

            '�������� ����� "������� ������"
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ������"
Dim intColNum As Integer

            '���������� ������ ��������
    SetColWidth
            '������� ������ = 0 (��������� ��������)
    grdTablePerson.Row = 0
    grdTablePerson.Col = 0
    grdTablePerson.Text = "Name"
            '�������� � ������ (������ 0, ������� 1)
    grdTablePerson.Col = 1
    grdTablePerson.Text = "PersonCode"
            '�������� � ������ (������ 0, ������� 2)
    grdTablePerson.Col = 2
    grdTablePerson.Text = "Status"
            '�������� � ������ (������ 0, ������� 3)
    grdTablePerson.Col = 3
    grdTablePerson.Text = "Time"
            '�������� � ������ (������ 0, ������� 4)
    grdTablePerson.Col = 4
    grdTablePerson.Text = "Calendar"
            '�������� � ������ (������ 0, ������� 5)
    grdTablePerson.Col = 5
    grdTablePerson.Text = "Reservation"
    
            
            '���������� "������� ������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� ������"
    lngRecordLen = Len(gPerson)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TablePerson.dat"
                
            '���� ����������� - ?
    On Error GoTo ErrorTablePerson
                '���������� ����� "������� ������" ����� ������� ����� �� ��������� +1
    grdTablePerson.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To grdTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        grdTablePerson.Row = intRowNum
            '������ ������ "������� ������" �� ����� � �����
        Get intFileNum, intRowNum, gPerson
            '�� ���� �������� "������� ������"
        For intColNum = 0 To grdTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            grdTablePerson.Col = intColNum
            '���������� ������� ������ "������� ������" �� ������
            Select Case intColNum
                Case 0
                grdTablePerson.Text = gPerson.strName
            '������� ������������� ������ "������� ������":
            '   ��������������� ������ � "Host Computer'e" � � ��� �������
            '   � "�������������", ����� ��������� ���������� ����
            '   ����������� "������� ������" - "���������� ������a ������" -
            '   �������� ���������� � ���������� ��������� ����:
                If gCompresTablPers = 1 Then
                    If Trim(gPerson.strName) = "ParkFreePlaces" Then _
            '�� �����������
                        gParkFreePlaces = gPerson.strPersonCode
                    ElseIf Trim(gPerson.strName) = "AcceFreePlaces" Then _
            '�� �����������
                        gAcceFreePlaces = gPerson.strPersonCode
                    End If
                End If
                Case 1
                grdTablePerson.Text = gPerson.strPersonCode
                Case 2
                grdTablePerson.Text = gPerson.strStatus
                If gPerson.strStatus = "00" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Terminal"
                If gPerson.strStatus = "01" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Computer"
                If gPerson.strStatus = "02" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Black card"
                If gPerson.strStatus = "03" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Relay"
                If gPerson.strStatus = "05" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Parking/Calen."
                If gPerson.strStatus = "06" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Parking/Time"
                If gPerson.strStatus = "07" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Parking/Free"
                If gPerson.strStatus = "08" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Access/Calen."
                If gPerson.strStatus = "09" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Access/Time"
                If gPerson.strStatus = "10" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Access/Free"
                Case 3
                 grdTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
                 "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
                Case 4
                grdTablePerson.Text = Left(gPerson.strCalendar, 2)
                If gPerson.strCalendar = "00" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Always"
                If gPerson.strCalendar = "01" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Standard"
                If gPerson.strCalendar = "02" Then grdTablePerson.Text = grdTablePerson.Text + _
                " - Special"
                If gPerson.strCalendar <> "00" And gPerson.strCalendar <> "01" _
                And Trim(gPerson.strCalendar) <> "02" Then grdTablePerson.Text = ""
                Case 5
                grdTablePerson.Text = gPerson.strReserve
            End Select
        Next
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� ������"
    grdTablePerson.Tag = 0
            '�������� ������� ��������� ��������� � "������� ������"
    gChangesTablePerson = False
            '����� ������ �� ���������� �������� ����� �� "������� ������"
    gRealDelPerson = False
    
    Exit Sub
ErrorTablePerson:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TablePerson Error !")
            '���������� ��������/���������� ����� � "������� ������"
    grdTablePerson.Tag = 0
    
End Sub

            '��������� ������� ������ "Reserve" ������� ������ "������� ������"
            '   ��� ������/������ ���������� � ���������� �������� �����������
Public Function AnalysisParking(ByVal vntWork As Variant)
           '������
Dim strStatus As String
            '����������� ������ "Reserve" ��� ��������� "��������" ����
            '   "����������" ������� ������ "������� ������"
Dim strChecking As String * 8
            '������������� ������ "Reserve" ��� ��������� "��������" ����
            '  "����������" ������� ������ "������� ������"
Dim strCheckingUnPack As String
            '������� ����������
Dim strDate As String
            '����� ���������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������� �������
Dim strHour As String
Dim strMinute As String

            '���������� ������� ��� �������� - ������ ����� �������� ��������
    AnalysisParking = 0
            '������� ������� "������� ������" = 2 (������)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 5 (������)
    gTablePerson.Col = 5
    strChecking = Trim(gTablePerson.Text)
            '�������� ����������� ��������� ������ ����������� �
            '   ����������� ������/������ ����������
    If (vntWork = 0 And (Mid(Trim(strChecking), 7, 1) = "1" Or _
    (Mid(Trim(strChecking), 8, 1) <> "E" And Mid(Trim(strChecking), 7, 1) = "2"))) Or _
    (vntWork = 1 And (Mid(Trim(strChecking), 7, 1) = "0" Or _
    Mid(Trim(strChecking), 7, 1) = "2")) Then
            
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
        If Left(Trim(strStatus), 2) <> "05" And Left(Trim(strStatus), 2) <> "06" And _
        Left(Trim(strStatus), 2) <> "07" Then
            GoTo AnalysisError
        End If
            '���������� ������ - ������ ����� �������� �������� ��� ������
        If Left(Trim(strStatus), 2) = "07" Then Exit Function
            
            '���������� ������ "��������"
        Call UnPack(strDate, strChecking)
            
            '������������ ������������� ������ "��������"
        strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/������/������ ������� � ������
        strCheckingUnPack = Trim(strCheckingUnPack) + Mid(Trim(strChecking), 7, 2)

            '���� ���������� �������
        strDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(strDate)
            
            '��������� ���������� ���� �������� ��� �� ��������
        If ((CInt(Mid(strCheckingUnPack, 7, 4)) = CInt(Right(strDate, 4))) And _
        (CInt(Mid(strCheckingUnPack, 4, 2)) > CInt(Mid(strDate, 4, 2))) Or _
        (CInt(Mid(strCheckingUnPack, 4, 2)) = _
        CInt(Mid(strDate, 4, 2)) And CInt(Left(strCheckingUnPack, 2)) >= _
        CInt(Left(strDate, 2)))) Or _
        (CInt(Mid(strCheckingUnPack, 7, 4)) > CInt(Right(strDate, 4))) Then
            '���������� ������ � ���������� �������/�������
            If Left(Trim(strStatus), 2) = "05" Then
            '����������� � ������������ ������� ������������ ����������
            '  � �� �������� � �����������
                If gParkTimeLimit > 0 And vntWork = 1 Then
            '������� ������� "������� ������" = 0 (����������)
                    gTablePerson.Col = 0
                    strChecking = Left(Trim(gTablePerson.Text), 6)
            '������������� ��������� "��������" ���� "����������"
                    Call UnPack(strDate, strChecking)
            '������������ ������������� ��������� "��������"
                    strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
                    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '��������� ����� ������ �������
                    strDate = Format(Now, "h:mm:ss")
            '����
                    intHour = Hour(strDate)
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '������
                    intMinute = Minute(strDate)
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
            '��������� ���������� ��� � ������ ��� �� ���������
                    If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
                    (CInt(Mid(strCheckingUnPack, 12, 2)) = _
                    CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
                    CInt(strMinute)) Then
                        Exit Function
                    End If
            '���������� ��� �������� = 2 (����� �������)
                    AnalysisParking = 2
                End If
                Exit Function
            '��������� ������
            Else
                GoTo Continue
            End If
            '��������� ���������� ���� �����
        Else
            GoTo AnalysisError
        End If
Continue:
            
            '��������� ������
        If Left(Trim(strStatus), 2) = "06" Then
            '��������� ����� ���������� �������
            strDate = Format(Now, "h:mm:ss")
            '����
           intHour = Hour(strDate)
            If intHour < 10 Then
                strHour = "0" + Trim(Str(intHour))
            Else
                strHour = Trim(Str(intHour))
            End If
            '������
            intMinute = Minute(strDate)
            If intMinute < 10 Then
                strMinute = "0" + Trim(Str(intMinute))
            Else
                strMinute = Trim(Str(intMinute))
            End If
            
            '��������� ���������� ��� � ������ ��� �� ���������
            If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
            (CInt(Mid(strCheckingUnPack, 12, 2)) = _
            CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
            CInt(strMinute)) Then
                Exit Function
            End If
        End If
            
    End If
    
            '�������� ������ ������� ��� ������������ �����������
            '  ���� ����� ������/������
AnalysisError:
            '���������� �� ������� ��� �������� - ������ ����� �������� �� ��������
    AnalysisParking = 1

End Function

            '��������� ������� ������ "Reserve" ������� ������ "������� ������"
            '   ��� �����/������ ���������� � ���������� ��������
Public Function AnalysisAccess(ByVal vntWork As Variant)
           '������
Dim strStatus As String
            '����������� ������ "Reserve" ��� ��������� "��������" ����
            '   "����������" ������� ������ "������� ������"
Dim strChecking As String * 8
            '������������� ������ "Reserve" ��� ��������� "��������" ����
            '  "����������" ������� ������ "������� ������"
Dim strCheckingUnPack As String
            '������� ����������
Dim strDate As String
            '����� ���������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������� �������
Dim strHour As String
Dim strMinute As String

            '���������� ������� ��� �������� - ������ ����� �������� ��������
    AnalysisAccess = 0
            '������� ������� "������� ������" = 2 (������)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 5 (������)
    gTablePerson.Col = 5
    strChecking = Trim(gTablePerson.Text)
            '�������� ����������� ��������� ������ ����������� �
            '   ����������� �����/������ ����������
    If (vntWork = 0 And (Mid(Trim(strChecking), 7, 1) = "1" Or _
    (Mid(Trim(strChecking), 8, 1) <> "E" And Mid(Trim(strChecking), 7, 1) = "2"))) Or _
    (vntWork = 1 And (Mid(Trim(strChecking), 7, 1) = "0" Or _
    Mid(Trim(strChecking), 7, 1) = "2")) Then
            
            '������ ������� �������
            
            '������������ ��� ����������� ������ �������
        If Left(Trim(strStatus), 2) <> "08" And Left(Trim(strStatus), 2) <> "09" And _
        Left(Trim(strStatus), 2) <> "10" Then
            GoTo AnalysisError
        End If
            '���������� ������ - ������ ����� �������� �������� ��� ������
        If Left(Trim(strStatus), 2) = "10" Then Exit Function
            
            '���������� ������ "��������"
        Call UnPack(strDate, strChecking)
        
            '������������ ������������� ������ "��������"
        strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/������/������ ������� � ������
        strCheckingUnPack = Trim(strCheckingUnPack) + Mid(Trim(strChecking), 7, 2)

            '���� ���������� �������
        strDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(strDate)
            
            '��������� ���������� ���� �������� ��� �� ��������
        If ((CInt(Mid(strCheckingUnPack, 7, 4)) = CInt(Right(strDate, 4))) And _
        (CInt(Mid(strCheckingUnPack, 4, 2)) > CInt(Mid(strDate, 4, 2))) Or _
        (CInt(Mid(strCheckingUnPack, 4, 2)) = _
        CInt(Mid(strDate, 4, 2)) And CInt(Left(strCheckingUnPack, 2)) >= _
        CInt(Left(strDate, 2)))) Or _
        (CInt(Mid(strCheckingUnPack, 7, 4)) > CInt(Right(strDate, 4))) Then
            '���������� ������ � ���������� ������-�������
            If Left(Trim(strStatus), 2) = "08" Then
            '����������� � ������������ ������� ������������ ����������
            '  � ������ ������� � �����������
                If gAcceTimeLimit > 0 And vntWork = 1 Then
            '������� ������� "������� ������" = 0 (����������)
                    gTablePerson.Col = 0
                    strChecking = Left(Trim(gTablePerson.Text), 6)
            '������������� ��������� "��������" ���� "����������"
                    Call UnPack(strDate, strChecking)
            '������������ ������������� ��������� "��������"
                    strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
                    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '��������� ����� ������ �������
                    strDate = Format(Now, "h:mm:ss")
            '����
                    intHour = Hour(strDate)
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '������
                    intMinute = Minute(strDate)
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
            '��������� ���������� ��� � ������ ��� �� ���������
                    If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
                    (CInt(Mid(strCheckingUnPack, 12, 2)) = _
                    CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
                    CInt(strMinute)) Then
                        Exit Function
                    End If
            '���������� ��� �������� = 2 (����� �������)
                    AnalysisAccess = 2
                End If
                Exit Function
            '��������� ������
            Else
                GoTo Continue
            End If
            '��������� ���������� ���� �����
        Else
            GoTo AnalysisError
        End If
Continue:
            
            '��������� ������
        If Left(Trim(strStatus), 2) = "09" Then
            '��������� ����� ���������� �������
            strDate = Format(Now, "h:mm:ss")
            '����
           intHour = Hour(strDate)
            If intHour < 10 Then
                strHour = "0" + Trim(Str(intHour))
            Else
                strHour = Trim(Str(intHour))
            End If
            '������
            intMinute = Minute(strDate)
            If intMinute < 10 Then
                strMinute = "0" + Trim(Str(intMinute))
            Else
                strMinute = Trim(Str(intMinute))
            End If
            
            '��������� ���������� ��� � ������ ��� �� ���������
            If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
            (CInt(Mid(strCheckingUnPack, 12, 2)) = _
            CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
            CInt(strMinute)) Then
            '����� �����������="1" - ���������� ������� � �����������
                If vntWork = 1 Then
            '������� ������� "������� ������" = 0 (����������)
                    gTablePerson.Col = 0
            '��������� �� ������ �������� ��� ���������
                    If Left(Trim(gTablePerson.Text), 4) = "0000" Then Exit Function
            '����� �����������="0" - ���������� ������ �� �����������
                Else
                    Exit Function
                End If
            End If
        End If
            
    End If
    
            '�������� ������ ������� ��� ������������ �����������
            '  ���� ����� �����/������, ���� �� ��������� ��������� ���������
AnalysisError:
            '���������� �� ������� ��� �������� - ������ ����� �������� �� ��������
    AnalysisAccess = 1

End Function

            '��������� ������� ������ "Name" ������� ������ "������� ������"
            '   ��� �����/������ ��������
Public Function AnalysisEmploye(ByVal vntWork As Variant)
Dim strStatus As String

            '���������� ������� ��� �������� - ������ ����� �������� ��������
    AnalysisEmploye = 0
            '������� ������� "������� ������" = 2 (������)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            '������� ������� "������� ������" = 0 (���)
    gTablePerson.Col = 0
            '�������� ����������� ��������� ������ ����������� �
            '   ����������� �����/������ ���������
    If (vntWork = 0 And Right(Trim(gTablePerson.Text), 1) = "-") Or _
    (vntWork = 1 And Right(Trim(gTablePerson.Text), 1) = "+") Or _
    Len(Trim(gTablePerson.Text)) < 16 Then
            
            '������ ������� ���������
            
            '������������ ������
        If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            GoTo AnalysisError
        End If
        Exit Function
    End If
    
            '�������� ������ ��������� ��� ������������ ����������� �����/������
AnalysisError:
            '���������� �� ������� ��� �������� - ������ ����� �������� �� ��������
    AnalysisEmploye = 1

End Function
            
            '���������� ������ "��������"
Public Sub UnPack(ByRef strDate As String, ByVal strChecking As String)
             '����� ������� ����������� � ������������� ����
Dim intPosNum As Integer
            '������� �������
Dim intCount As Integer
        
    strDate = ""
            '����� ������a "z" (7AH) � ������ ���
            '  ������� ����� (������������� � ����� - "")
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "z")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("00"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� ������a "x" (78H) � ������ ���
            '  �� "09" (������������� ����� - "")
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "x")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("09"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� ������a "y" (79H) � ������ ���
            '  �� "10" (������������� ������� �������)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "y")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("10"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� ������a "w" (77H) � ������ ���
            '  �� "13" (������������� ������� �������)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "w")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("13"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� ������a "r" (72H) � ������ ���
            '  �� "32" (������������� ������� �������)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "r")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("32"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '���� � �����
    For intCount = 1 To 6 Step 1
        If Asc(Mid(strChecking, intCount, 1)) < 10 Then
            strDate = strDate + "0" + _
            Trim(Str(Asc(Mid(strChecking, intCount, 1))))
        Else
            strDate = strDate + _
            Trim(Str(Asc(Mid(strChecking, intCount, 1))))
        End If
    Next

End Sub
            
            '�������� ������ "��������"
Public Sub Pack(ByRef strChecking As String)
             '����� ������� ����������� � ������������� ����
Dim intPosNum As Integer
            '������� �������
Dim intCount As Integer
            
            '����� �������� ���� (������������ � ����� - "")
            '  � ������ ��� �������� "z" (7AH)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("00"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "z" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� "09" (������������ � ����� - "")
            '  � ������ ��� �������� "x" (78H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("09"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "x" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� "10" (������������ � ������� �������)
            '  � ������ ��� �������� "y" (79H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("10"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "y" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� "13" (������������ � ������� �������)
            '  � ������ ��� �������� "w" (77H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("13"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "w" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            '����� "32" (������������ ����� - "")
            '  � ������ ��� �������� "r" (72H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("32"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "r" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next

End Sub
            
            '���������� ������ � �������� ������ � "������� ������"
            '  �� ��������� MSMQ, ���������� �� ����
Public Function MSMQReg(ByVal strMessage As String)
            '����� ������� ������ � "������� ������"
Dim intRowNum As Integer
            '������ ������ "������� ������"
Dim strPerson As String
            '���� ������ "Person or Terminal" � "������� ������"
Dim strName As String
            '���� ������ "PersonCode" � "������� ������"
Dim vntPersonCode As Variant
            '���� ������ "Status" � "������� ������"
Dim strStatus As String
            '���� ������ "Time" � "������� ������"
Dim strTime As String
            '���� ������ "Calendar" � "������� ������"
Dim strCalendar As String
            '���� ������ "Reserve" � "������� ������"
Dim strReserve As String
            '����� ������������� ������ � ������ ���������
Dim intNumber As Integer
            
            '������� ������������� ������ "������� ������" �� ����������:
            '   "������������" ���������� "������� ������" "Host Computer'�"
            '   - ����� �� ���������
    If gCompresTablPers = 0 Then Exit Function
    
            '����� ������������� ������ � ������ ���������
    intNumber = 1
            '������ � ������ ��������� ������� "07H" - ����������� �������
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            '���������� ������ "Person or Terminal" ��� "������� ������"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            '���������� ������ "PersonCode" ��� "������� ������"
            vntPersonCode = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            '���������� ������ "Status" ��� "������� ������"
            strStatus = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            '���������� ������ "Time" ��� "������� ������"
            strTime = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            '���������� ������ "Calendar" ��� "������� ������"
            strCalendar = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            '���������� ������ "Reserve" ��� "������� ������"
            strReserve = strMessage
            '�������������� �����, �.�. ������ "Reserve" ����� �������� "07H"
            Exit Do
        End If
    Loop
        
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
    
            '����������� ������������� ���� ��� � "������� ������"
    If intRowNum = gTablePerson.Rows Then
            '���������� ������ � ����� "������� ������"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ������"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            '��������� ������ "Status" � "������� ������"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            '��������� ������ "Time" � "������� ������"
        gTablePerson.Col = 3
        gTablePerson.Text = strTime
            '��������� ������ "Calendar" � "������� ������"
        gTablePerson.Col = 4
        gTablePerson.Text = strCalendar
            '��������� ������ "Reserve" � "������� ������"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            
    
            '���������� ��������/���������� ����� � "������� ������"
        gTablePerson.Tag = gTablePerson.Tag + 1
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    
    End If
    
End Function
            
            '�������� (����������) ������ � �������� ������������ �����
            '  �� "������� ������" �� ��������� MSMQ, ���������� �� ����
Public Function MSMQDel(ByVal vntPersonCode As Variant)
            '������� ����� ��������������� ������ "������� ������"
Dim intRowNum As Integer
            
            '������� ������������� ������ "������� ������" �� ����������:
            '   "������������" ���������� "������� ������" "Host Computer'�"
            '   - ����� �� ���������
    If gCompresTablPers = 0 Then Exit Function
        
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
    
            '���������� ������������ ��� ���� � "������� ������"
    If intRowNum < gTablePerson.Rows Then
            '��������� ������� ������ �� "������� ������"
        gTablePerson.Text = "Deleted"
            '���������� ������ �� �������� ��������
        gRealDelPerson = True
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    End If
    
End Function
            
            '��������� �������� ����� ������ "������� ������"
            '  �� ��������� MSMQ, ���������� �� ����
Public Function MSMQCor(ByVal strMessage As String)
            '����� ������� ������ "������� ������"
Dim intRowNum As Integer
            '������ ������ "������� ������"
Dim strPerson As String
            '���� ������ "Person or Terminal" � "������� ������"
Dim strName As String
            '���� ������ "PersonCode" � "������� ������"
Dim vntPersonCode As Variant
            '���� ������ "Status" � "������� ������"
Dim strStatus As String
            '���� ������ "Time" � "������� ������"
Dim strTime As String
            '���� ������ "Calendar" � "������� ������"
Dim strCalendar As String
            '���� ������ "Reserve" � "������� ������"
Dim strReserve As String
            '����� ������������� ������ � ������ ���������
Dim intNumber As Integer
            
            '������� ������������� ������ "������� ������" �� ����������:
            '   "������������" ���������� "������� ������" "Host Computer'�"
            '   - ����� �� ���������
    If gCompresTablPers = 0 Then Exit Function
    
            '����� ������������� ������ � ������ ���������
    intNumber = 1
            '������ � ������ ��������� ������� "07H" - ����������� �������
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            '���������� ������ "Person or Terminal" ��� "������� ������"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            '���������� ������ "PersonCode" ��� "������� ������"
            vntPersonCode = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            '���������� ������ "Status" ��� "������� ������"
            strStatus = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            '���������� ������ "Time" ��� "������� ������"
            strTime = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            '���������� ������ "Calendar" ��� "������� ������"
            strCalendar = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            '���������� ������ "Reserve" ��� "������� ������"
            strReserve = strMessage
            '�������������� �����, �.�. ������ "Reserve" ����� �������� "07H"
            Exit Do
        End If
    Loop
        
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
    
            '���������� ������������ ��� ���� � "������� ������"
    If intRowNum < gTablePerson.Rows Then
        gTablePerson.Row = intRowNum
            '��������� ������ "Person or Terminal" � "������� ������"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            '��������� ������ "PersonCode" � "������� ������"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            '��������� ������ "Status" � "������� ������"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            '��������� ������ "Time" � "������� ������"
        gTablePerson.Col = 3
        gTablePerson.Text = strTime
            '��������� ������ "Calendar" � "������� ������"
        gTablePerson.Col = 4
        gTablePerson.Text = strCalendar
            '��������� ������ "Reserve" � "������� ������"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            
            '���������� ������� ��������� ��������� � "������� ������"
        gChangesTablePerson = True
    
    End If
    
End Function

            '��������� ��������� ������ � ������������ �������� "������� ������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '���� �� ���� ��������
    For intColNumber = 0 To grdTablePerson.Cols - 1 Step 1
        grdTablePerson.ColWidth(intColNumber) = 1650
        grdTablePerson.ColAlignment(intColNumber) = 0
    Next
            '���������� ������� 0-�� � 1-�� �������� (��� � ������������ ���)
    intColNumber = 0
    grdTablePerson.ColWidth(intColNumber) = 2500
    intColNumber = 1
    grdTablePerson.ColWidth(intColNumber) = 2500
    
End Sub

