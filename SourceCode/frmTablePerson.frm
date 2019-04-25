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
            'Текущий номер корректируемой строки "Таблицы персон"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Таблицы персон"
Dim intColNumCorr As Integer
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Таблицы персон"
Dim gPerson As PersonInfo
            'Строка отсылаемого сообщения
Dim strMessage As String


            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            
            'Были не сохраненные изменения в "Таблице персон"
    If gChangesTablePerson = True Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы персон" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранение "Tаблицы персон" в файле по умолчанию
            cmdSave_Click
        End If
    End If
    
            'Сделать недоступными элементы управления Коррекцией "Таблицы персон"
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
            'Очистить текстовые поля
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            'Очистить список имен
    lstName.Clear
            'Сбросить признак внесенных изменений в "Таблицу персон"
    gChangesTablePerson = False
            'Сделать невидимой текущую форму
    frmTablePerson.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()
            
            ' "Таблица персон" не содержит нефиксированных строк
    If gTablePerson.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
    
    Else
            'Сделать доступными некоторые элементы управления
            '   Коррекцией "Таблицы персон"
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
            'Очистить текстовые поля
        txtName.Text = ""
        txtPersonCode.Text = ""
        txtReservation.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtType.Text = ""
            'Очистить список имен
        lstName.Clear
    
            'Столбец "Person or Terminal"
        gTablePerson.Col = 0
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNumCorr = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNumCorr
            'Заполнение списка "lstName" записями из "Таблицы персон"
            lstName.AddItem gTablePerson.Text
        Next
            'Выбрать  элемент списка
        lstName.ListIndex = 0
            'Номер корректируемой строки - (1)
        intRowNumCorr = 1
        gTablePerson.Row = intRowNumCorr
            'Включить опцию
        optName_Click
    End If
    
End Sub
            
            'Процедура восстановления умалчиваемых значений
            ' параметров "Таблицы персон"
Private Sub cmdDefaultPers_Click()
            
            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы персон"
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
            'Очистить текстовые поля
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            'Очистить список имен
    lstName.Clear
            
            'Загрузка формы "Таблица персон"
    Form_Load
            'Установить фокус на кнопке "Correction"
    If frmTablePerson.Visible = True Then cmdCorrection.SetFocus

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Таблицы персон"
Private Sub grdTablePerson_Click()
            'Коррекция "включена"
    If lstName.Enabled = True Then
            'Номер корректируемой строки "Таблицы персон"
        intRowNumCorr = gTablePerson.RowSel
        gTablePerson.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstName.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Таблицы персон"
        intColNumCorr = gTablePerson.ColSel
        gTablePerson.Col = intColNumCorr
            'Выбор корректируемой строки "Таблицы персон"
        lstName_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstName.Left, Y:=lstName.Top
            'Выбор корректируемого столбца "Таблицы персон"
        Select Case intColNumCorr
            Case 1
            optPersonCode.Value = True
            'Установить фокус на текстовом поле для Коррекции
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


            'Выбор корректируемой строки "Таблицы персон"
Private Sub lstName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемой строки "Таблицы персон"
        intRowNumCorr = lstName.ListIndex + 1
        gTablePerson.Row = intRowNumCorr
        gTablePerson.Col = 0
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
        txtName.Text = gTablePerson.Text
            'Номер анализируемого столбца "Таблицы персон" - "PersonCode"
        gTablePerson.Col = 1
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
        txtPersonCode.Text = gTablePerson.Text
            'Номер анализируемого столбца "Таблицы персон" - "Status"
        gTablePerson.Col = 2
            'Не установлена опция "Relay"
        If Left(gTablePerson.Text, 2) <> "03" Then
            'Номер корректируемого столбца "Таблицы персон"
            gTablePerson.Col = 5
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
            txtReservation.Text = gTablePerson.Text
            'Очистка недоступных для Коррекции текстовых полей
            txtAddress.Text = ""
            txtType.Text = ""
            txtPort.Text = ""
            'Установлена опция "Relay"
        Else
            'Номер корректируемого столбца "Таблицы персон"
            gTablePerson.Col = 4
            'Принудительное изменение ячейки "Calendar" в "Таблице персон"
            gTablePerson.Text = "00 - Always"
            'Установить признак  внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            
            'Номер корректируемого столбца "Таблицы персон"
            gTablePerson.Col = 5
            'Копирование ячейки "Таблицы персон" в текстовые поля для Коррекции
            txtAddress.Text = Left(gTablePerson.Text, 2)
            txtPort.Text = Mid(gTablePerson.Text, 3, 1)
            txtType.Text = Mid(gTablePerson.Text, 4)
            'Очистка недоступного для Коррекции текстового поля
            txtReservation.Text = ""
        End If
            'Восстановить номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = intColNumCorr
    End If

End Sub

            'Выбрана опция - "Pers. or Term."
Private Sub optName_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 0
    gTablePerson.Col = intColNumCorr
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
    txtName.Text = gTablePerson.Text
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
    txtName.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtName.SetFocus
            'Номер отображаемого столбца "Таблицы персон"
    gTablePerson.Col = 1
            'Копирование ячейки "Таблицы персон" в текстовое поле для Отображения
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
            'Номер анализируемого столбца "Таблицы персон" - "Status"
    gTablePerson.Col = 2
            'Не установлена опция "Relay"
    If Left(gTablePerson.Text, 2) <> "03" Then
            'Номер отображаемого столбца "Таблицы персон"
        gTablePerson.Col = 5
            'Копирование ячейки "Таблицы персон" в текстовое поле для Отображения
        txtReservation.Text = gTablePerson.Text
            'Очистка недоступных для Отображения текстовых полей
        txtAddress.Text = ""
        txtType.Text = ""
        txtPort.Text = ""
            'Установлена опция "Relay"
    Else
            'Номер отображаемого столбца "Таблицы персон"
        gTablePerson.Col = 5
            'Копирование ячейки "Таблицы персон" в текстовые поля для Отображения
        txtAddress.Text = Left(gTablePerson.Text, 2)
        txtPort.Text = Mid(gTablePerson.Text, 3, 1)
        txtType.Text = Mid(gTablePerson.Text, 4)
            'Очистка недоступного для Отображения текстового поля
        txtReservation.Text = ""
    End If
            'Восстановить номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "PersonCode"
Private Sub optPersonCode_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 1
    gTablePerson.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
    lblName.Enabled = False
    txtName.Enabled = False
    lblPersonCode.Enabled = True
    txtPersonCode.Enabled = True
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
    txtPersonCode.Text = gTablePerson.Text
            'Установить фокус на текстовом поле для Коррекции
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


            'Выбрана опция - "Status"
Private Sub optStatus_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 2
    gTablePerson.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
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
            'Включить опцию "Computer"
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

            'Выбрана опция - "Time"
Private Sub optTime_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 3
    gTablePerson.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
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
            
            'Выбрана опция - "Calendar"
Private Sub optCalendar_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 4
    gTablePerson.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
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
            'Номер анализируемого столбца "Таблицы персон" - "Status"
    gTablePerson.Col = 2
            'Установлены опции "Terminal" или "Relay"
    If Left(gTablePerson.Text, 2) = "00" Or Left(gTablePerson.Text, 2) = "03" Then
            'Включить опцию "Standard"
        optAlways.Value = True
            'Установлены опции "Terminal" или "Relay"
    Else
            'Включить опцию "Standard"
        optStandard.Value = True
    End If
            'Восстановить номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = intColNumCorr


End Sub

            'Выбрана опция "Reservation"
Private Sub optReservation_Click()
            'Номер корректируемого столбца "Таблицы персон"
    intColNumCorr = 5
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы персон"
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
    
                'Номер анализируемого столбца "Таблицы персон" - "Status"
    gTablePerson.Col = 2
            'Не установлена опция "Relay"
    If Left(gTablePerson.Text, 2) <> "03" Then
            'Номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = intColNumCorr
        lblReservation.Enabled = True
        txtReservation.Enabled = True
            'Копирование ячейки "Таблицы персон" в текстовое поле для Коррекции
        txtReservation.Text = gTablePerson.Text
            'Установить фокус на текстовом поле для Коррекции
        txtReservation.SetFocus
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblType.Enabled = False
        txtType.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
            'Очиска недоступных для Коррекции текстовых полей
        txtAddress.Text = ""
        txtType.Text = ""
        txtPort.Text = ""
            'Установлена опция "Relay"
    Else
            'Номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = 4
            'Принудительное изменение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Text = "00 - Always"
            'Установить признак  внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            
            'Номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = intColNumCorr
        lblReservation.Enabled = False
        txtReservation.Enabled = False
        lblAddress.Enabled = True
        txtAddress.Enabled = True
            'Копирование ячейки "Таблицы персон" в текстовые поля для Коррекции
        txtAddress.Text = Left(gTablePerson.Text, 2)
        txtPort.Text = Mid(gTablePerson.Text, 3, 1)
        txtType.Text = Mid(gTablePerson.Text, 4)
            'Установить фокус на текстовом поле для Коррекции
        txtAddress.SetFocus
        lblType.Enabled = True
        txtType.Enabled = True
        lblPort.Enabled = True
        txtPort.Enabled = True
            'Очиска недоступного для Коррекции текстового поля
        txtReservation.Text = ""
    End If

End Sub

            'Процедура ввода и анализа Корректируемого имени "Person or Terminal"
Private Sub txtName_KeyPress(KeyAscii As Integer)
            'Имя введено
    If KeyAscii = vbKeyReturn Then
            'Имя в допустимом диапазоне
        If Len(Trim(txtName.Text)) < 17 Then
            'Изменение имени "Person or Terminal" в "Таблице персон"
        gTablePerson.Text = Trim(txtName.Text)
            'Установить признак  внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            'Включить опцию "optPersonCode"
        optPersonCode.Value = True
            Exit Sub
            'Имя в недопустимом диапазоне
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo PersonCodeError
            'Персональный код в допустимом диапазоне
        If Len(Trim(txtPersonCode.Text)) = 16 Then
            'Изменение ячейки "PersonCode" в "Таблице персон"
            gTablePerson.Text = Trim(txtPersonCode.Text)
            'Установить признак  внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            'Включить опцию "optStatus"
            optStatus.Value = True
            Exit Sub
            'Персональный код в недопустимом диапазоне
PersonCodeError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            'Выбрана опция - "Terminal"
Private Sub optTerminal_GotFocus()
            'Изменение ячейки "Status" в "Таблице персон"
    gTablePerson.Text = "00 - Terminal"
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 4
            'Принудительное изменение ячейки "Calendar" в "Таблице персон"
    gTablePerson.Text = "00 - Always"
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 2
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub

            'Выбрана опция - "Computer"
Private Sub optComputer_GotFocus()
            'Изменение ячейки "Status" в "Таблице персон"
    gTablePerson.Text = "01 - Computer"
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub

            'Выбрана опция - "BlackCard"
Private Sub optBlackCard_GotFocus()
            'Изменение ячейки "Status" в "Таблице персон"
    gTablePerson.Text = "02 - Black card"
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub

            'Выбрана опция - "Relay"
Private Sub optRelay_GotFocus()
            'Изменение ячейки "Status" в "Таблице персон"
    gTablePerson.Text = "03 - Relay"
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 4
            'Принудительное изменение ячейки "Calendar" в "Таблице персон"
    gTablePerson.Text = "00 - Always"
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 2
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Hour"
Private Sub hsbHour_Scroll()
    hsbHour_Change
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Hour"
Private Sub hsbHour_Change()
            'Начало временного интервала
    If chkFromToTime.Value = 1 Then
            'Изменение ячейки "Time" в "Таблице персон"
        If hsbHour.Value < 10 Then
            gTablePerson.Text = "0" + Trim(Str(hsbHour.Value)) + Mid(gTablePerson.Text, 3)
        Else
            gTablePerson.Text = Trim(Str(hsbHour.Value)) + Mid(gTablePerson.Text, 3)
        End If
            'Конец временного интервала
    Else
            'Изменение ячейки "Time" в "Таблице персон"
        If hsbHour.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 6) + "0" + Trim(Str(hsbHour.Value)) _
            + Mid(gTablePerson.Text, 9)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 6) + Trim(Str(hsbHour.Value)) _
            + Mid(gTablePerson.Text, 9)
        End If
    End If
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True
    
End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Minute"
Private Sub hsbMinute_Scroll()
    hsbMinute_Change
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Minute"
Private Sub hsbMinute_Change()
            'Начало временного интервала
    If chkFromToTime.Value = 1 Then
            'Изменение ячейки "Time" в "Таблице персон"
        If hsbMinute.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 3) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 6)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 6)
        End If
            'Конец временного интервала
    Else
            'Изменение ячейки "Time" в "Таблице персон"
        If hsbMinute.Value < 10 Then
            gTablePerson.Text = Left(gTablePerson.Text, 9) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 12)
        Else
            gTablePerson.Text = Left(gTablePerson.Text, 9) + Trim(Str(hsbMinute.Value)) _
            + Mid(gTablePerson.Text, 12)
        End If
    End If
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub

            'Выбрана опция - "Always"
Private Sub optAlways_GotFocus()
            'Изменение ячейки "Calendar" в "Таблице персон"
    gTablePerson.Text = "00 - Always"
            'Установить признак  внесенных изменений в "Таблицу персон"
    gChangesTablePerson = True

End Sub

            'Выбрана опция - "Standard"
Private Sub optStandard_GotFocus()
            'Номер анализируемого столбца "Таблицы персон" - "Status"
    gTablePerson.Col = 2
            'Не установлены опции "Terminal" и "Relay"
    If Left(gTablePerson.Text, 2) <> "00" And Left(gTablePerson.Text, 2) <> "03" Then
            'Номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = 4
            'Изменение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Text = "01 - Standard"
            'Установить признак  внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 4

End Sub

            'Выбрана опция - "Special"
Private Sub optSpecial_GotFocus()
            'Номер анализируемого столбца "Таблицы персон" - "Status"
    gTablePerson.Col = 2
            'Не установлены опции "Terminal" и "Relay"
    If Left(gTablePerson.Text, 2) <> "00" And Left(gTablePerson.Text, 2) <> "03" Then
            'Номер корректируемого столбца "Таблицы персон"
        gTablePerson.Col = 4
            'Изменение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Text = "02 - Special"
            'Установить признак  внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
            'Номер корректируемого столбца "Таблицы персон"
    gTablePerson.Col = 4

End Sub
            
            'Процедура ввода и анализа Корректируемого поля "Reservation"
Private Sub txtReservation_KeyPress(KeyAscii As Integer)
            'Номер позиции признака "/" в Корректируемом поле
Dim intPosNum As Integer
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Номер анализируемого столбца "Таблицы персон" - "Calendar"
        gTablePerson.Col = 4
            'Установлена опция "Special"
        If Left(gTablePerson.Text, 2) = "02" Then
            'Номер корректируемого столбца "Таблицы персон"
            gTablePerson.Col = intColNumCorr
            'Установлены признаки ..."/"..."/"... - Дополн. табл.: врем. интерв., термин. и календаря
            intPosNum = InStr(1, Trim(txtReservation.Text), "/")
            If intPosNum <> 0 And InStr(intPosNum + 1, Trim(txtReservation.Text), "/") <> 0 _
            And Len(Trim(txtReservation.Text)) < 9 Then
            'Ошибка преобразования форматов данных
                On Error GoTo TimeTerCalError
            'Коррекция номера "Таблицы времени"(...//) в ячейке "Reservation" "Таблицы персон"
                If intPosNum > 1 Then
            'Номер "Таблицы времени" в допустимом диапазоне номеров (0/99)
                    If intPosNum < 4 And Left(Trim(txtReservation.Text), intPosNum - 1) < 100 Then
            
            'Номер "Таблицы времени" в недопустимом диапазоне номеров
                    Else
                        GoTo TimeTerCalError
                    End If
                End If
            'Коррекция номера "Таблицы Терминалов"(/.../) в ячейке "Reservation" "Таблицы персон"
                If InStr(intPosNum + 1, Trim(txtReservation.Text), "/") > intPosNum + 1 Then
            'Номер "Таблицы терминалов" в допустимом диапазоне номеров (1/99)
                    If InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum < 4 And _
                    Mid(Trim(txtReservation.Text), intPosNum + 1, _
                    InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum - 1) > 0 And _
                    Mid(Trim(txtReservation.Text), intPosNum + 1, _
                    InStr(intPosNum + 1, Trim(txtReservation.Text), "/") - intPosNum - 1) < 100 Then
            
            'Номер "Таблицы терминалов" в недопустимом диапазоне номеров
                    Else
                        GoTo TimeTerCalError
                    End If
                End If
            'Коррекция номера "Таблицы индивидуального календаря"(//...) в ячейке
            '  "Reservation" "Таблицы персон"
                If Len(Trim(txtReservation.Text)) > InStr(intPosNum + 1, _
                Trim(txtReservation.Text), "/") Then
            'Номер "Таблицы индивидуального календаря" в допустимом диапазоне
            '  номеров (1/99)
                    If Len(Trim(txtReservation.Text)) - InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") < 3 And _
                    Mid(Trim(txtReservation.Text), InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 1) > 0 And _
                    Mid(Trim(txtReservation.Text), InStr(intPosNum + 1, _
                    Trim(txtReservation.Text), "/") + 1) < 100 Then
            
            'Номер "Таблицы индивидуального календаря" в недопустимом диапазоне номеров
                    Else
                            GoTo TimeTerCalError
                    End If
                End If
            
            'Изменение в ячейке "Reservation" "Таблицы персон"
                gTablePerson.Text = Trim(txtReservation.Text)
            
            'Удаление незначащих нулей - в обратном порядке (справа налево)
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
                
            'Установить признак  внесенных изменений в "Таблицу персон"
                gChangesTablePerson = True
            'Установить фокус на кнопке "Save"
                cmdSave.SetFocus
                Exit Sub
            End If
            'Номера "Таблицы времени", "Таблицы терминалов" или
            '  "Таблицы индивидуального календаря" в недопустимом диапазоне номеров
TimeTerCalError:
            frmDemo.BeepSound
            'Не установлена опция "Special"
        Else
                    'Номер корректируемого столбца "Таблицы персон"
            gTablePerson.Col = intColNumCorr
            If Len(Trim(txtReservation.Text)) < 9 Then
            'Изменение поля "Reservation" в "Таблице персон"
                gTablePerson.Text = Trim(txtReservation.Text)
            'Установить признак  внесенных изменений в "Таблицу персон"
                gChangesTablePerson = True
            'Установить фокус на кнопке "Save"
                cmdSave.SetFocus
            'Неверный формат данных
            Else
                frmDemo.BeepSound
            End If
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Reservation - Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            'Адрес введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo AddressError
            'Адрес в допустимом диапазоне адресов (01/15,  00 - групповой адрес)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text > 0 And txtAddress.Text < 16 Then
            'Изменение ячейки "Reservation" в "Таблице персон"
            If Len(Trim(gTablePerson.Text)) < 4 Then
                txtPort.Text = "2"
                txtType.Text = "CONTR"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Trim(txtAddress.Text) + Mid(gTablePerson.Text, 3)
            End If
            'Установить признак  внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            'Установить фокус на текстовом поле "Port"
            txtPort.SetFocus
            Exit Sub
            'Адреса в недопустимом диапазоне адресов
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            'Процедура ввода и анализа Корректируемого "Reservation - Port"
Private Sub txtPort_KeyPress(KeyAscii As Integer)
            'Номер порта введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo PortError
            'Номер порта в допустимом диапазоне (2/8)
        If Len(Trim(txtPort.Text)) = 1 And txtPort.Text > 1 And txtPort.Text < 9 Then
            'Изменение ячейки "Reservation" в "Таблице персон"
            If Len(Trim(gTablePerson.Text)) < 4 Then
                txtAddress.Text = "01"
                txtType.Text = "CONTR"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Left(gTablePerson.Text, 2) + Trim(txtPort.Text) + _
                Mid(gTablePerson.Text, 4)
            End If
            'Установить признак  внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            'Установить фокус на текстовом поле "Type"
            txtType.SetFocus
            Exit Sub
            'Номер порта в недопустимом диапазоне
PortError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Reservation - Type"
Private Sub txtType_KeyPress(KeyAscii As Integer)
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo TypeError
            'Код в допустимом диапазоне кодов ("XXXXX")
        If Len(Trim(txtType.Text)) <= 5 Then
            'Изменение ячейки "Reservation" в "Таблице персон"
            If Len(Trim(gTablePerson.Text)) = 0 Then
                txtAddress.Text = "01"
                txtPort.Text = "2"
                gTablePerson.Text = Trim(txtAddress.Text) + Trim(txtPort.Text) + _
                Trim(txtType.Text)
            Else
                gTablePerson.Text = Left(gTablePerson.Text, 3) + Trim(txtType.Text)
            End If
            'Установить признак  внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            'Установить фокус на кнопке "Save"
            cmdSave.SetFocus
            Exit Sub
            'Код в недопустимом диапазоне кодов
TypeError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Добавление строки в "Таблицу персон"
Private Sub cmdAdd_Click()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Имя и Персональный код в "Таблице персон"
Dim strName As String
Dim strPersonCode As String

    strName = ""
    strPersonCode = ""
    
            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы персон"
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
            'Очистить список имен
    lstName.Clear
    
            
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя Имя персоны
    strName = InputBox("Name: 1 -- 16 Characters !!!", "Add ...")
    If Len(Trim(strName)) > 16 Then strName = Left(Trim(strName), 16)
    frmDemo.BeepSound
            'Получить от пользователя Персональный код
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Add ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strPersonCode)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            'Имя или Персональный код не выбраны
    If strName = "" Or strPersonCode = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The Name Or PersonCode isn't selected"
            
            'Имя и Персональный код выбраны
    Else
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
            If Trim(gTablePerson.Text) = strPersonCode Then
            'Досрочный выход из цикла
                Exit For
            End If
        Next
            'Введенный ПЕРСОНАЛЬНЫЙ КОД уже есть в "Таблице персон"
        If intRowNum < gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            MsgBox ("Duplicated PersonCode")
            'Введенного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
        Else
            'Добавление строки в конец "Таблицы персон"
            gTablePerson.AddItem strPersonCode
            gTablePerson.Row = gTablePerson.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
            gTablePerson.Col = 0
            gTablePerson.Text = Trim(strName)
            'Изменение ячейки "PersonCode" в "Таблице персон"
            gTablePerson.Col = 1
            gTablePerson.Text = Trim(strPersonCode)
            'Формирование шаблона для интервала времени
            gTablePerson.Col = 3
            gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
            "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            'Количество удалений/добавлений строк в "Таблице персон"
            gTablePerson.Tag = gTablePerson.Tag + 1
            'Установить признак внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            '  Протоколировать данное событие
            gProtocol.strProtocName = strName
            gProtocol.strProtocPersonCode = strPersonCode
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
            gProtocol.strProtocReserve = "PersonCode Addition"
            'Записать строку в файл "Таблицы протокола"
            frmDemo.WriteProtocol
        End If
    End If
    
            'Установить фокус на кнопке "Add"
    If frmTablePerson.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            'Поиск строки в "Таблице персон"
Private Sub cmdFind_Click()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Имя и Персональный код в "Таблице Персон"
Dim strName As String
Dim strPersonCode As String

    strName = ""
    strPersonCode = ""
    
            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы персон"
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
            'Очистить список имен
    lstName.Clear
    
            'Получить от пользователя Имя персоны
    strName = InputBox("Name: 1 -- 16 Characters !!!", "Find ...")
    If Len(Trim(strName)) > 16 Then strName = Left(Trim(strName), 16)
    frmDemo.BeepSound
            'Получить от пользователя Персональный код
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Find ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strPersonCode)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            'Имя или Персональный код не выбраны
    If strName = "" Or strPersonCode = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The Name Or PersonCode isn't selected"
            
            'Имя и Персональный код выбраны
    Else
            
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
            If Trim(gTablePerson.Text) = strPersonCode Then
            'Текущий столбец "Таблицы Персон" = 0 (Имя)
                gTablePerson.Col = 0
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
                If InStr(1, Trim(gTablePerson.Text), strName) <> 0 Then
            'Отобразить ТОЛЬКО ТЕКСТОВЫЕ ПОЛЯ "Таблицы персон"
                    txtPersonCode.Text = gTablePerson.Text
                    gTablePerson.Col = 0
                    txtName.Text = gTablePerson.Text
                    gTablePerson.Col = 5
                    txtReservation.Text = gTablePerson.Text
            'Досрочный выход из цикла
                    Exit For
                End If
            End If
        Next
            'ИМЕНИ или ПЕРСОНАЛЬНОГО КОДА нет в "Таблице Персон"
        If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            MsgBox ("Unexistent Name Or PersonCode")
            'Установить фокус на кнопке "Correction"
            If frmTablePerson.Visible = True Then cmdCorrection.SetFocus
            
            'ИМЯ и ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице Персон"
        Else
            'Сделать доступными некоторые элементы управления
            '   Коррекцией "Таблицы персон"
            fraColName.Enabled = True
            txtName.Enabled = True
            'Очистить текстовые поля
            txtName.Text = ""
            txtPersonCode.Text = ""
            txtReservation.Text = ""
            txtAddress.Text = ""
            txtPort.Text = ""
            txtType.Text = ""
            'Очистить список имен
            lstName.Clear
            'Включить опцию
            optName_Click
        End If
    
    End If
    
End Sub

            
            'АвтоРегистрация персонального кода в "Таблице персон" для Автостоянки
            '   Код возврата: 0 - АвтоРегистрация выполнена успешно;
            '                 1 - в АвтоРегистрации отказано.
Public Function AutoRegParking(ByVal vntPersonCode As Variant, _
ByVal strName As String, ByVal strStatus As String, ByVal strReserve As String, _
ByVal strTime As String)
            'Номер текущей строки в "Таблице персон"
Dim intRowNum As Integer
            'Номер позиции признака "/" в анализируемом поле
Dim intPosNum As Integer
            'Начало временного интервала - Часы
Dim intHourStart As Integer
            'Начало временного интервала - Минуты
Dim intMinuteStart As Integer
            'Конец временного интервала - Часы
Dim intHourFinish As Integer
            'Конец временного интервала - Минуты
Dim intMinuteFinish As Integer
            'Пустая строка "Таблицы персон"
Dim strPerson As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоРегистрации отказано
    AutoRegParking = 1
    
            'Введенные ПЕРСОНАЛЬНЫЙ КОД или ИНФОРМАЦИЯ
            '  уже есть в "Таблице персон"
    If intRowNum < gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode or Info")
        Else
            MsgBox ("Person. kods vai Info jau ir")
        End If
            'Введенного ПЕРСОНАЛЬНОГО КОДА и ИНФОРМАЦИИ
            '  нет в "Таблице персон"
    Else
            'Добавление строки в конец "Таблицы персон"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            'Изменение ячейки "PersonCode" в "Таблице персон"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            'Изменение ячейки "Status" в "Таблице персон"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            'Изменение ячейки "Reserve" в "Таблице персон"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            'Умалчиваемое значение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultParkCale
            'Формирование шаблона для интервала времени
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
        "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            'Временной интервал допуска к Автостоянке (выбран при АвтоРегистрации)
        If Trim(strTime) = "Day" Then
            strTime = Trim(gParkingTimeD)
        ElseIf Trim(strTime) = "DayNight" Then
            strTime = Trim(gDefaultParkTime)
        ElseIf Trim(strTime) = "Night" Then
            strTime = Mid(Trim(gParkingTimeD), 7) + "/" + Left(Trim(gParkingTimeD), 5)
        End If
            
            'Начало временного интервала - Часы
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            'Начало временного интервала - Минуты
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Часы
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Минуты
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            'Изменение ячейки "Time" в "Таблице персон" - Часы
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
            'Изменение ячейки "Time" в "Таблице персон" - Минуты
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
            
            'Строка передачи сообщения
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultParkCale + _
        Chr(7) + strReserve
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
    
            'Количество удалений/добавлений строк в "Таблице персон"
        gTablePerson.Tag = gTablePerson.Tag + 1
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            'АвтоРегистрация выполнена успешно
        AutoRegParking = 0
    
    End If
    
End Function
            
            'Коррекция ячейки "Reserve" в "Таблице персон"
Public Function InputParking(intIndex As Integer)
            'Поле "Name" в "Таблице персон"
Dim strName  As String
            'Статус
Dim strStatus As String
            'Подстрока "Контроль"
Dim strCheckingInfo As String * 8
            'Подстрока "Контрольные Дата и Время"
Dim strDateInfo As String
            'Время
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время
Dim strHour As String
Dim strMinute As String
            'Позиция в строке
Dim intPosNum As Integer
            'Рабочая переменная
Dim intCount As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
Dim intCellLimit As Integer

            'Установлен признак процедуры Регистрации Клиента Автостоянки
            '  (Автомобиль въехал) после Регистрации Клиента Автостоянки
    If frmDemo.imgParkingInData(intIndex).Tag = 1 Then
            'Номер строки в конце "Таблицы персон" (Последняя добавленная)
        gTablePerson.Row = gTablePerson.Rows - 1
            'Номер столбца в "Таблице персон" ("Reserve")
        gTablePerson.Col = 5
            'Корректная ситуация - Автомобиль Зарегистрирован
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Автомобиль въехал)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputParking = 0
            'Некорректная ситуация
        Else
            InputParking = 1
        End If
            'Признак процедуры Регистрации Клиента Автостоянки не установлен -
            '  (Автомобиль въехал) по Ключу Постоянного или Бесплатного Клиента
    Else
            'Номер столбца в текущей строке "Таблицы персон" ("Reserve")
        gTablePerson.Col = 5
            'Корректная ситуация - Автомобиль Зарегистрирован или ранее выезжал
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "1" And _
        Mid(Trim(gTablePerson.Text), 8, 1) <> "E" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Автомобиль въехал)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputParking = 0
            
            'Автостоянка с ограничением времени непрерывного пребывания
            If gParkTimeLimit > 0 Then
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
                intCellLimit = gParkingCellLimit
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            'Постоянный Клиент
                If Left(Trim(strStatus), 2) = "05" Then
            'Очистка подстроки "Контроль"
                    strCheckingInfo = ""
            
            'Время въезда Клиента
                    strDateInfo = Format(Now, "h:mm:ss")
            'Часы
                    intHour = Hour(strDateInfo)
            'Минуты
                    intMinute = Minute(strDateInfo)
            'Дата въезда Клиента
                    strDateInfo = Format(Now, "dd/mm/yyyy")
                    strDateInfo = Left(Trim(strDateInfo), 2) + Mid(Trim(strDateInfo), 4, 2) + _
                    Right(Trim(strDateInfo), 4)
                        
            
            'Вычислить "сдвинутые" время и дату въезда Постоянного
            '  Клиента, до которых ему будет разрешен бесплатный выезд
        
            'Требуется переход часа
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
            
            'Требуется переход даты
                        If intHour >= 24 Then
                            intHour = intHour - 24
            'Установка "Календаря" на дату, следующую за текущей
                            frmTableCalendar.comCalendar.Today
                            frmTableCalendar.comCalendar.NextDay
            'Изменение  Числа
                            If frmTableCalendar.comCalendar.Day > 9 Then
                                strDateInfo = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            Else
                                strDateInfo = "0" + _
                                Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            End If
            'Изменение  Месяца и, возможно, Года
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
            
            'Не требуется переход часа
                    Else
                        intMinute = intMinute + gParkingTimeCell * intCellLimit + _
                        gParkingTimeCell
                    End If
            
            '"Сдвинутые" Часы
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '"Сдвинутые" Минуты
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
    
            'Формирование упакованной подстроки "Контроль"
                    For intCount = 1 To 7 Step 2
            'Дата
                        strCheckingInfo = Trim(strCheckingInfo) + _
                        Chr(CByte(CInt(Mid(strDateInfo, intCount, 2))))
                    Next
            'Часы
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            'Минуты
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            'Упаковка подстроки "Контроль"
                    Call frmTablePerson.Pack(strCheckingInfo)
            
            'Коррекция поля "Name" в "Таблице персон"
                    gTablePerson.Col = 0
                    gTablePerson.Text = Left(strCheckingInfo, 6) + _
                    Mid(Trim(gTablePerson.Text), 7)
                    
                End If
            End If
            
            'Некорректная ситуация
        Else
            InputParking = 1
        End If
    End If
            'Kорректная ситуация
    If InputParking = 0 Then
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'АвтоПоиск персонального кода из "Таблицы персон" для Автостоянки
            '   Код возврата: 0 - АвтоПоиск выполнен успешно;
            '                 1 - АвтоПоиск неуспешный.
Public Function AutoFindParking(ByVal vntPersonCode As Variant, strProtocName As String, _
                                                strProtocStatus As String, strChecking As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strProtocStatus = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
            gTablePerson.Col = 5
            strChecking = Trim(gTablePerson.Text)
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'АвтоПоиск неуспешный
    AutoFindParking = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            'АвтоПоиск выполнен успешно
        AutoFindParking = 0
    End If
    
End Function
            
            'АвтоУдаление (ЛОГИЧЕСКОЕ) персонального кода АвтоКлиента
            '  из "Таблицы персон"
            '  Код возврата: 0 - АвтоУдаление выполнено успешно;
            '                1 - в АвтоУдалении отказано.
Public Function AutoDelParking(ByVal vntPersonCode As Variant, strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоУдалении отказано
    AutoDelParking = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
'Восстановить номер текущей строки "Таблицы персон"
            gTablePerson.Row = intRowNum
            'НЕМЕДЛЕННОЕ Удаление
            If gParkingDeletion = 1 Then
            
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                gTablePerson.Col = 1
            
            'Строка передачи сообщения
                strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
                gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
                gRealDelPerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                Call frmDemo.SendMessage(strMessage)
            
            'ОТЛОЖЕННОЕ до выхода/выезда Удаление
            Else
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
                gTablePerson.Col = 5
            'Клиент ранее уже выехал - нестандартная ситуация
            ' (НЕМЕДЛЕННОЕ Удаление)
                If Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                    gTablePerson.Col = 1
            
            'Строка передачи сообщения
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
                    gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
                    gRealDelPerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            ' ОТЛОЖЕННОЕ Удаление
                Else
                    gTablePerson.Text = Left(Trim(gTablePerson.Text), 7) + "E"
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
            
            'АвтоУдаление выполнено успешно
            AutoDelParking = 0
            
            'Установить признак внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
        End If
    End If
    
End Function
            
            'Коррекция ячейки "Reserve" или исключение строки в "Таблице персон"
Public Function OutputParking(intIndex As Integer, intStatusCode As Integer)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            
            'Номер столбца в текущей строке "Таблицы персон" ("Reserve")
    gTablePerson.Col = 5
            'Установлен признак выезда Автомобиля Временного Клиента или
            '  Окончательного (без возврата) выезда Постоянного Клиента
    If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Or intStatusCode = 6 Then
            'Номер столбца в текущей строке "Таблицы персон" ("Person")
        gTablePerson.Col = 1
            
            'Строка передачи сообщения
        strMessage = "Del " + Trim(gTablePerson.Text)
            'Логически удалить строку из "Таблицы персон"
        gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
        gRealDelPerson = True
            
            'Корректная ситуация - Клиент Удаляется из "Таблицы персон"
        OutputParking = 0
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Признак процедуры Удаления Клиента не установлен -
            '  (Автомобиль выехал) по Ключу Постоянного или Бесплатного Клиента
    Else
            'Корректная ситуация - Клиент Зарегистрирован или ранее въезжал
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Клиент выехал)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "1" + _
            Right(Trim(gTablePerson.Text), 1)
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call frmDemo.SendMessage(strMessage)
            
            'Корректная ситуация
            OutputParking = 0
            
            'Некорректная ситуация
        Else
            OutputParking = 1
        End If
    End If
            
            'Kорректная ситуация
    If OutputParking = 0 Then
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'Сжатие "Таблицы персон" для Автостоянки (Удаление строк с
            '  информацией об Окончательно выехавших АМ и Установка Кода
            '  возврата:     "0" -  Сжатие выполнено успешно;
            '                "1" -  на Автостоянке присутствуют AM, которые должны
            '                       были Окончательно выехaть после оплаты парковки;
            '                       оплаты парковки;
            '                "2" -  Сжатие невозможно.
Public Function AutoPresParking()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Установка Кода возврата функции
    AutoPresParking = 0
            ' "Таблица персон" не содержит нефиксированных строк
    If gTablePerson.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности получения информации
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            'Установка Кода возврата функции
        AutoPresParking = 2
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Клиента Автостоянки
            If Left(Trim(gTablePerson.Text), 2) = "07" Or _
            Left(Trim(gTablePerson.Text), 2) = "05" Or _
            Left(Trim(gTablePerson.Text), 2) = "06" Then
            'Столбец - "Reserve"
                gTablePerson.Col = 5
            'Признак АМ, которая должна была
            '  Окончательно выехать после оплаты
                If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Then
            'Номер столбца в текущей строке "Таблицы персон" ("Person")
                    gTablePerson.Col = 1
            'Присутствует АМ, которая должна была
            '  Окончательно выехать после оплаты, но не выехала
                    If gTablePerson.Text <> "Deleted" Then AutoPresParking = 1
                End If
            End If
        Next
    End If

End Function
            
            'АвтоКоррекция измененных ячеек "Таблицы персон" для Автостоянки
            '   Код возврата: 0 - АвтоКоррекция выполнена успешно;
            '                 1 - АвтоКоррекция неуспешная.
Public Function AutoCorParking(ByVal vntPersonCode As Variant, ByVal strName _
As String, ByVal strChecking As String, ByRef strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            gTablePerson.Text = strName
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
            gTablePerson.Col = 5
            gTablePerson.Text = strChecking
            
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'АвтоКоррекция неуспешная
    AutoCorParking = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        MsgBox "Correction impossible !"
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            
            'Строка передачи сообщения
        strMessage = "Cor " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7)
            'Текущий столбец "Таблицы персон" = 3 (Время)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            'Текущий столбец "Таблицы персон" = 4 (Дата)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7) + strChecking
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'АвтоКоррекция успешная
        AutoCorParking = 0
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'АвтоРегистрация персонального кода в "Таблице персон" для Предприятия
            '   Код возврата: 0 - АвтоРегистрация выполнена успешно;
            '                 1 - в АвтоРегистрации отказано.
Public Function AutoRegAccess(ByVal vntPersonCode As Variant, _
ByVal strName As String, ByVal strStatus As String, ByVal strReserve As String, _
ByVal strTime As String)
            'Номер текущей строки в "Таблице персон"
Dim intRowNum As Integer
            'Номер позиции признака "/" в анализируемом поле
Dim intPosNum As Integer
            'Начало временного интервала - Часы
Dim intHourStart As Integer
            'Начало временного интервала - Минуты
Dim intMinuteStart As Integer
            'Конец временного интервала - Часы
Dim intHourFinish As Integer
            'Конец временного интервала - Минуты
Dim intMinuteFinish As Integer
            'Пустая строка "Таблицы персон"
Dim strPerson As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоРегистрации отказано
    AutoRegAccess = 1
    
            'Введенные ПЕРСОНАЛЬНЫЙ КОД или ИНФОРМАЦИЯ
            '  уже есть в "Таблице персон"
    If intRowNum < gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode")
        Else
            MsgBox ("Person. kods jau ir")
        End If
            'Введенного ПЕРСОНАЛЬНОГО КОДА и ИНФОРМАЦИИ
            '  нет в "Таблице персон"
    Else
            'Добавление строки в конец "Таблицы персон"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            'Изменение ячейки "PersonCode" в "Таблице персон"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            'Изменение ячейки "Status" в "Таблице персон"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            'Изменение ячейки "Reserve" в "Таблице персон"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            'Умалчиваемое значение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultAcceCale
            'Формирование шаблона для интервала времени
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
        "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            'Временной интервал допуска (выбран при АвтоРегистрации)
        If Trim(strTime) = "Day" Then
            strTime = Trim(gAccessTimeD)
        ElseIf Trim(strTime) = "DayNight" Then
            strTime = Trim(gDefaultAcceTime)
        ElseIf Trim(strTime) = "Night" Then
            strTime = Mid(Trim(gAccessTimeD), 7) + "/" + Left(Trim(gAccessTimeD), 5)
        End If
            
            'Начало временного интервала - Часы
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            'Начало временного интервала - Минуты
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Часы
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Минуты
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            'Изменение ячейки "Time" в "Таблице персон" - Часы
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
            'Изменение ячейки "Time" в "Таблице персон" - Минуты
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
            
            'Строка передачи сообщения
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultAcceCale + _
        Chr(7) + strReserve
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Количество удалений/добавлений строк в "Таблице персон"
        gTablePerson.Tag = gTablePerson.Tag + 1
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            'АвтоРегистрация выполнена успешно
        AutoRegAccess = 0
        
    End If
    
End Function
            
            'Коррекция ячейки "Reserve" в "Таблице персон"
Public Function InputAccess(intIndex As Integer)
            'Поле "Name" в "Таблице персон"
Dim strName  As String
            'Статус
Dim strStatus As String
            'Подстрока "Контроль"
Dim strCheckingInfo As String * 8
            'Подстрока "Контрольные Дата и Время"
Dim strDateInfo As String
            'Время
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время
Dim strHour As String
Dim strMinute As String
            'Позиция в строке
Dim intPosNum As Integer
            'Рабочая переменная
Dim intCount As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  Постоянному Клиенту непрерывно находиться на Предприятии
Dim intCellLimit As Integer


            'Установлен признак процедуры Регистрации Клиента
            '  (Посетитель вошел) после Регистрации Клиента
    If frmDemo.imgAccessInData(intIndex).Tag = 1 Then
            'Номер строки в конце "Таблицы персон" (Последняя добавленная)
        gTablePerson.Row = gTablePerson.Rows - 1
            'Номер столбца в "Таблице персон" ("Reserve")
        gTablePerson.Col = 5
            'Корректная ситуация - Посетитель Зарегистрирован
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Посетитель вошел)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputAccess = 0
            'Некорректная ситуация
        Else
            InputAccess = 1
        End If
            'Признак процедуры Регистрации Клиента не установлен -
            '  (Посетитель вошел) по Ключу Постоянного или Бесплатного Клиента
    Else
            'Номер столбца в текущей строке "Таблицы персон" ("Reserve")
        gTablePerson.Col = 5
            'Корректная ситуация - Посетитель Зарегистрирован или ранее выходил
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "1" And _
        Mid(Trim(gTablePerson.Text), 8, 1) <> "E" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Посетитель вошел)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "0" + _
            Right(Trim(gTablePerson.Text), 1)
            InputAccess = 0
            
            'Предприятие с ограничением времени непрерывного пребывания
            If gAcceTimeLimit > 0 Then
            'Количество ячеек времени, в течение которого разрешается
            '  Постоянному Клиенту непрерывно находиться на Предприятии
                intCellLimit = gAccessCellLimit
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            'Постоянный Клиент
                If Left(Trim(strStatus), 2) = "08" Then
            'Очистка подстроки "Контроль"
                    strCheckingInfo = ""
            
            'Время входа Клиента
                    strDateInfo = Format(Now, "h:mm:ss")
            'Часы
                    intHour = Hour(strDateInfo)
            'Минуты
                    intMinute = Minute(strDateInfo)
            'Дата входа Клиента
                    strDateInfo = Format(Now, "dd/mm/yyyy")
                    strDateInfo = Left(Trim(strDateInfo), 2) + Mid(Trim(strDateInfo), 4, 2) + _
                    Right(Trim(strDateInfo), 4)
                        
            
            'Вычислить "сдвинутые" время и дату входа Постоянного
            '  Клиента, до которых ему будет разрешен бесплатный выход
        
            'Требуется переход часа
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
            
            'Требуется переход даты
                        If intHour >= 24 Then
                            intHour = intHour - 24
            'Установка "Календаря" на дату, следующую за текущей
                            frmTableCalendar.comCalendar.Today
                            frmTableCalendar.comCalendar.NextDay
            'Изменение  Числа
                            If frmTableCalendar.comCalendar.Day > 9 Then
                                strDateInfo = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            Else
                                strDateInfo = "0" + _
                                Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                                Right(strDateInfo, 6)
                            End If
            'Изменение  Месяца и, возможно, Года
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
            
            'Не требуется переход часа
                    Else
                        intMinute = intMinute + gAccessTimeCell * intCellLimit + _
                        gAccessTimeCell
                    End If
            
            '"Сдвинутые" Часы
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            '"Сдвинутые" Минуты
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
    
            'Формирование упакованной подстроки "Контроль"
                    For intCount = 1 To 7 Step 2
            'Дата
                        strCheckingInfo = Trim(strCheckingInfo) + _
                        Chr(CByte(CInt(Mid(strDateInfo, intCount, 2))))
                    Next
            'Часы
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            'Минуты
                    strCheckingInfo = Trim(strCheckingInfo) + _
                    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            'Упаковка подстроки "Контроль"
                    Call frmTablePerson.Pack(strCheckingInfo)
            
            'Коррекция поля "Name" в "Таблице персон"
                    gTablePerson.Col = 0
                    gTablePerson.Text = Left(strCheckingInfo, 6) + _
                    Mid(Trim(gTablePerson.Text), 7)
                    
                End If
            End If
            
            'Некорректная ситуация
        Else
            InputAccess = 1
        End If
    End If
            'Kорректная ситуация
    If InputAccess = 0 Then
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'АвтоПоиск персонального кода из "Таблицы персон" для Предприятия
            '   Код возврата: 0 - АвтоПоиск выполнен успешно;
            '                 1 - АвтоПоиск неуспешный.
Public Function AutoFindAccess(ByVal vntPersonCode As Variant, strProtocName As String, _
                                                strProtocStatus As String, strChecking As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strProtocStatus = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
            gTablePerson.Col = 5
            strChecking = Trim(gTablePerson.Text)
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'АвтоПоиск неуспешный
    AutoFindAccess = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            'АвтоПоиск выполнен успешно
        AutoFindAccess = 0
    End If
    
End Function
            
            'АвтоУдаление (ЛОГИЧЕСКОЕ) персонального кода Посетителя
            '  из "Таблицы персон"
            '  Код возврата: 0 - АвтоУдаление выполнено успешно;
            '                1 - в АвтоУдалении отказано.
Public Function AutoDelAccess(ByVal vntPersonCode As Variant, strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоУдалении отказано
    AutoDelAccess = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            
            'Коррекция информации о ПРОКАТЕ ИНВЕНТАРЯ
        If frmLease.Tag <> "Exit" Then _
        'Текущий столбец "Таблицы персон" = 0 (Информация)
            gTablePerson.Col = 0
            gTablePerson.Text = Trim(frmDataAccessOut.txtInfo.Text)
        End If
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
'Восстановить номер текущей строки "Таблицы персон"
            gTablePerson.Row = intRowNum
            'НЕМЕДЛЕННОЕ Удаление
            If gAccessDeletion = 1 Then
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                gTablePerson.Col = 1
            
            'Строка передачи сообщения
                strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
                gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
                gRealDelPerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                Call frmDemo.SendMessage(strMessage)
            
            'ОТЛОЖЕННОЕ до выхода/выезда Удаление
            Else
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
                gTablePerson.Col = 5
            'Посетитель ранее уже вышел - нестандартная ситуация
            ' (НЕМЕДЛЕННОЕ Удаление)
                If Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                    gTablePerson.Col = 1
            
            'Строка передачи сообщения
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
                    gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
                    gRealDelPerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            ' ОТЛОЖЕННОЕ Удаление
                Else
                    gTablePerson.Text = Left(Trim(gTablePerson.Text), 7) + "E"
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
            
            'АвтоУдаление выполнено успешно
            AutoDelAccess = 0
            'Установить признак внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
        End If
    End If
    
End Function
            
            'Коррекция ячейки "Reserve" или исключение строки в "Таблице персон"
Public Function OutputAccess(intIndex As Integer, intStatusCode As Integer)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
            'Номер столбца в текущей строке "Таблицы персон" ("Reserve")
    gTablePerson.Col = 5
            'Установлен признак выхода Временного Посетителя или
            '  Окончательного (без возврата) выхода Постоянного Посетителя
    If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Or intStatusCode = 9 Then
            'Номер столбца в текущей строке "Таблицы персон" ("Person")
        gTablePerson.Col = 1
            
            'Строка передачи сообщения
        strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
        gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
        gRealDelPerson = True
            'Корректная ситуация - Посетитель Удаляется из "Таблицы персон"
        OutputAccess = 0
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Признак процедуры Удаления Клиента не установлен -
            '  (Посетитель вышел) по Ключу Постоянного или Бесплатного Клиента
    Else
            'Корректная ситуация - Посетитель Зарегистрирован или ранее входил
        If Mid(Trim(gTablePerson.Text), 7, 1) = "2" Or _
        Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
            'Изменение ячейки "Reserve" в "Таблице персон" (Посетитель вышел)
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 6) + "1" + _
            Right(Trim(gTablePerson.Text), 1)
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call frmDemo.SendMessage(strMessage)
            
            'Kорректная ситуация
            OutputAccess = 0
            'Некорректная ситуация
        Else
            OutputAccess = 1
        End If
    End If
            
            'Kорректная ситуация
    If OutputAccess = 0 Then
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'Сжатие "Таблицы персон" для Предприятия (Удаление строк с
            '  информацией об Окончательно вышедших Посетителях и Установка Кода
            '  возврата:     "0" -  Сжатие выполнено успешно;
            '                "1" -  присутствуют Посетители, которые должны
            '                       были Окончательно выйдти после оплаты;
            '                "2" -  Сжатие невозможно.
Public Function AutoPresAccess()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Установка Кода возврата функции
    AutoPresAccess = 0
            ' "Таблица персон" не содержит нефиксированных строк
    If gTablePerson.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности получения информации
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            'Установка Кода возврата функции
        AutoPresAccess = 2
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Клиента
            If Left(Trim(gTablePerson.Text), 2) = "10" Or _
            Left(Trim(gTablePerson.Text), 2) = "08" Or _
            Left(Trim(gTablePerson.Text), 2) = "09" Then
            'Столбец - "Reserve"
                gTablePerson.Col = 5
            'Признак Посетителя, который должен был
            '  Окончательно выйти после оплаты
                If Mid(Trim(gTablePerson.Text), 8, 1) = "E" Then
            'Номер столбца в текущей строке "Таблицы персон" ("Person")
                    gTablePerson.Col = 1
            'Присутствуют Посетитель, который должен был
            '  Окончательно выйти после оплаты, но не вышел
                    If gTablePerson.Text <> "Deleted" Then AutoPresAccess = 1
                End If
            End If
        Next
    End If

End Function
            
            'АвтоКоррекция измененных ячеек "Таблицы персон"
            '   Код возврата: 0 - АвтоКоррекция выполнена успешно;
            '                 1 - АвтоКоррекция неуспешная.
Public Function AutoCorAccess(ByVal vntPersonCode As Variant, ByVal strName _
As String, ByVal strChecking As String, ByRef strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
            gTablePerson.Col = 0
            gTablePerson.Text = strName
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
            gTablePerson.Col = 5
            gTablePerson.Text = strChecking
            
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'АвтоКоррекция неуспешная
    AutoCorAccess = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        MsgBox "Correction impossible !"
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            
            'Строка передачи сообщения
        strMessage = "Cor " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        strStatus + Chr(7)
            'Текущий столбец "Таблицы персон" = 3 (Время)
        gTablePerson.Col = 3
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7)
            'Текущий столбец "Таблицы персон" = 4 (Дата)
        gTablePerson.Col = 4
        strMessage = strMessage + Trim(gTablePerson.Text) + Chr(7) + strChecking
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'АвтоКоррекция успешная
        AutoCorAccess = 0
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    
    End If
    
End Function
            
            'Очистка "Таблицы персон" для Предприятия (Удаление строк с
            '  информацией о Временных Посетителях и Установка Кода
            '  возврата:     "0" -  Очистка выполнено успешно;
            '                "1" -  Очиска невозможна.
Public Function CleaningAccess()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            
            'Установка Кода возврата функции
    CleaningAccess = 0
            ' "Таблица персон" не содержит нефиксированных строк
    If gTablePerson.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности Очистки
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            'Установка Кода возврата функции
        CleaningAccess = 1
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Клиента - Временный
            If Left(Trim(gTablePerson.Text), 2) = "09" Then
            'Столбец - "PersonCode"
                gTablePerson.Col = 1
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
                intButtonsAndIcons = vbYesNo + vbQuestion
                If frmDemo.optEnglish = True Then
                    strResponse = MsgBox("Deletion PersonCode = " + _
                    Trim(gTablePerson.Text), intButtonsAndIcons, "Cancel")
                Else
                    strResponse = MsgBox("Izslegt person. kods = " + _
                    Trim(gTablePerson.Text), intButtonsAndIcons, "Cancel")
                End If
            'Нажата кнопка "Да"
                If strResponse = vbYes Then
'Восстановить номер текущей строки "Таблицы персон"
                    gTablePerson.Row = intRowNum
            
            'Строка передачи сообщения
                    strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
                    gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
                    gRealDelPerson = True
            'Установить признак внесенных изменений в "Таблицу персон"
                    gChangesTablePerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
                End If
            End If
        Next
    End If

End Function
            
            'АвтоРегистрация персонального кода Служащего в "Таблице персон"
            '   Код возврата: 0 - АвтоРегистрация выполнена успешно;
            '                 1 - в АвтоРегистрации отказано.
Public Function AutoRegEmploye(ByVal vntPersonCode As Variant, _
ByVal strName As String)
            'Номер текущей строки в "Таблице персон"
Dim intRowNum As Integer
            'Номер позиции признака "/" в анализируемом поле
Dim intPosNum As Integer
            'Начало временного интервала - Часы
Dim intHourStart As Integer
            'Начало временного интервала - Минуты
Dim intMinuteStart As Integer
            'Конец временного интервала - Часы
Dim intHourFinish As Integer
            'Конец временного интервала - Минуты
Dim intMinuteFinish As Integer
            'Пустая строка "Таблицы персон"
Dim strPerson As String
Dim strTime As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоРегистрации отказано
    AutoRegEmploye = 1

            'Введенные ПЕРСОНАЛЬНЫЙ КОД уже есть в "Таблице персон"
    If intRowNum < gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Duplicated PersonCode")
        Else
            MsgBox ("Person. kods jau ir")
        End If
            'Введенного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    Else
            'Добавление строки в конец "Таблицы персон"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            'Изменение ячейки "PersonCode" в "Таблице персон"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            'Изменение ячейки "Status" в "Таблице персон"
        gTablePerson.Col = 2
        gTablePerson.Text = gDefaultStatus
            'Умалчиваемое значение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Col = 4
        gTablePerson.Text = gDefaultCalendar
            'Формирование шаблона для интервала времени
        gTablePerson.Col = 3
        gTablePerson.Text = Left(gPerson.strTime, 2) + "." + Mid(gPerson.strTime, 3, 2) + _
            "-" + Mid(gPerson.strTime, 5, 2) + "." + Mid(gPerson.strTime, 7, 2) + " - Inter."
            
            'Временной интервал допуска (Умалчиваемый)
        strTime = gDefaultTime
            
            'Начало временного интервала - Часы
        intPosNum = InStr(2, strTime, "/")
        intHourStart = Left(strTime, intPosNum - 1)
            'Начало временного интервала - Минуты
        intMinuteStart = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Часы
        intHourFinish = Mid(strTime, intPosNum + 1, _
        InStr(intPosNum + 1, strTime, "/") - intPosNum - 1)
        intPosNum = InStr(intPosNum + 1, strTime, "/")
            'Конец временного интервала - Минуты
        intMinuteFinish = Right(strTime, _
        Len(strTime) - intPosNum)
            
            'Изменение ячейки "Time" в "Таблице персон" - Часы
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
            'Изменение ячейки "Time" в "Таблице персон" - Минуты
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
            
            'Строка передачи сообщения
        strMessage = "Reg " + strName + Chr(7) + Trim(vntPersonCode) + Chr(7) + _
        gDefaultStatus + Chr(7) + gTablePerson.Text + Chr(7) + gDefaultCalendar + _
        Chr(7) + " "
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Количество удалений/добавлений строк в "Таблице персон"
        gTablePerson.Tag = gTablePerson.Tag + 1
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            'АвтоРегистрация выполнена успешно
        AutoRegEmploye = 0
    
    End If
               
End Function
            
            'Коррекция ячейки "Name" в "Таблице персон"
Public Function InputEmploye(intIndex As Integer)
            'Номер столбца в текущей строке "Таблицы персон" ("Name")
    gTablePerson.Col = 0
            'Корректная ситуация - Служащий Зарегистрирован или ранее выходил
    If Len(Trim(gTablePerson.Text)) < 16 Or _
    Right(Trim(gTablePerson.Text), 1) = "-" Then
            'Изменение ячейки "Name" в "Таблице персон" (Cлужащий вошел)
        If Len(Left(Trim(gTablePerson.Text), 15)) < 15 Then
            gTablePerson.Text = Trim(gTablePerson.Text) + _
            Left("              ", 15 - Len(Trim(gTablePerson.Text))) + "+"
        Else
            gTablePerson.Text = Left(Trim(gTablePerson.Text), 15) + "+"
        End If
        InputEmploye = 0
            'Некорректная ситуация
    Else
            InputEmploye = 1
    End If
            'Kорректная ситуация
    If InputEmploye = 0 Then
            
            'Строка передачи сообщения
        strMessage = "Cor " + Trim(gTablePerson.Text) + Chr(7)
            'Изменение ячейки "PersonCode" в "Таблице персон"
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'АвтоПоиск Персонального кода или Имени Служащего в "Таблице персон"
            '   Код возврата: 0 - АвтоПоиск выполнен успешно;
            '                 1 - АвтоПоиск неуспешный.
Public Function AutoFindEmploye(vntPersonCode As Variant, _
vntInfo As Variant, strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
    
            'АвтоПоиск неуспешный
    AutoFindEmploye = 1
        
        'Поиск по Персональномуй коду
    If vntPersonCode <> "" Then
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
            If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
                gTablePerson.Col = 0
                frmDataEmployeOut.txtInfo = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            'АвтоПоиск выполнен успешно
                AutoFindEmploye = 0
            'Досрочный выход из цикла
                Exit For
            End If
        Next
        
        'Поиск по Имени
    ElseIf vntInfo <> "" Then
        'Текущий столбец "Таблицы персон" = 0 (Имя)
        gTablePerson.Col = 0
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Полученное ИМЯ есть в "Таблице персон"
            If InStr(1, Trim(gTablePerson.Text), Trim(vntInfo)) <> 0 Then
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
                gTablePerson.Col = 1
                frmDataEmployeOut.txtPersonCode = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 2 (Статус)
                gTablePerson.Col = 2
                strStatus = Trim(gTablePerson.Text)
            'АвтоПоиск выполнен успешно
                AutoFindEmploye = 0
            'Досрочный выход из цикла
                Exit For
            End If
        Next
    End If
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If AutoFindEmploye = 1 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
    End If
    
End Function
            
            'АвтоУдаление (ЛОГИЧЕСКОЕ) персонального кода Служащего
            '  из "Таблицы персон"
            '  Код возврата: 0 - АвтоУдаление выполнено успешно;
            '                1 - в АвтоУдалении отказано.
Public Function AutoDelEmploye(vntPersonCode As Variant, strStatus As String)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
    
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Текущий столбец "Таблицы персон" = 2 (Статус)
            gTablePerson.Col = 2
            strStatus = Trim(gTablePerson.Text)
            'Досрочный выход из цикла
            Exit For
        End If
    Next
            'В АвтоУдалении отказано
    AutoDelEmploye = 1
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent PersonCode")
        Else
            MsgBox ("Nav person. koda")
        End If
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    Else
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Deletion PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Izslegt person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
'Восстановить номер текущей строки "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
            gTablePerson.Col = 1
            
            'Строка передачи сообщения
            strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
            gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
            gRealDelPerson = True
            
            'АвтоУдаление выполнено успешно
            AutoDelEmploye = 0
            'Установить признак внесенных изменений в "Таблицу персон"
            gChangesTablePerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        End If
    End If
    
End Function
            
            'Коррекция ячейки "Name" в "Таблице персон"
Public Function OutputEmploye(intIndex As Integer)
            'Номер столбца в текущей строке "Таблицы персон" ("Name")
    gTablePerson.Col = 0
            'Корректная ситуация - Служащий Зарегистрирован или ранее входил
    If Len(Trim(gTablePerson.Text)) < 16 Or _
    Right(Trim(gTablePerson.Text), 1) = "+" Then
            'Вышел Гость
        If Left(Trim(gTablePerson.Text), 1) = gVisitor Then
            'Номер столбца в текущей строке "Таблицы персон" ("Person")
            gTablePerson.Col = 1
            
            'Строка передачи сообщения
            strMessage = "Del " + Trim(gTablePerson.Text)
            
            'Логически удалить строку из "Таблицы персон"
            gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
            gRealDelPerson = True
            
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        
            'Вышел НЕ Гость
        Else
            'Изменение ячейки "Name" в "Таблице персон" (Cлужащий вошел)
            If Len(Left(Trim(gTablePerson.Text), 15)) < 15 Then
                gTablePerson.Text = Trim(gTablePerson.Text) + _
                Left("              ", 15 - Len(Trim(gTablePerson.Text))) + "-"
            Else
                gTablePerson.Text = Left(Trim(gTablePerson.Text), 15) + "-"
            End If
            
            'Строка передачи сообщения
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
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
            Call frmDemo.SendMessage(strMessage)
        End If
        
            'Kорректная ситуация
        OutputEmploye = 0
        
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
        
            'Некорректная ситуация
    Else
        OutputEmploye = 1
    End If
    
End Function
            
            'Сжатие "Таблицы персон" для Служащих (Поиск Гостей и Установка Кода
            '  возврата:     "0" -  Сжатие выполнено успешно;
            '                "1" -  присутствуют Гости;
            '                "2" -  Сжатие невозможно.
Public Function AutoPresEmploye()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Установка Кода возврата функции
    AutoPresEmploye = 0
            ' "Таблица персон" не содержит нефиксированных строк
    If gTablePerson.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности получения информации
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            'Установка Кода возврата функции
        AutoPresEmploye = 2
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Служащего
            If Left(Trim(gTablePerson.Text), 2) = "00" Or _
            Left(Trim(gTablePerson.Text), 2) = "01" Then
            'Столбец - "Name"
                gTablePerson.Col = 0
            'Присутствуют Гости - Установка Кода возврата функции
                If Left(Trim(gTablePerson.Text), 1) = gVisitor Then AutoPresEmploye = 1
            End If
        Next
    End If

End Function
            
            'Физическое удаление логически удаленных строк 'Таблицы персон"
Public Function RealDelPerson()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            
            'Код возврата при обращении к функциям объекта "Mutex"
Dim lngRetCode As Long
            
            'Последняя строка не может быть удалена - выход из процедуры
    If gTablePerson.Rows = 2 Then Exit Function
    
            'Если "Таблицa персон" все еще доступна для вычеркивания строк
    If gTablePerson.Access < 1 Then
            'Ждать освобождения объекта "Mutex"
        lngRetCode = WaitForSingleObject(gMutex, 15000)
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = 1
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Строка логически удалена
            If Trim(gTablePerson.Text) = "Deleted" Then
            'Физическое удаление строки из "Таблицы персон"
                gTablePerson.RemoveItem gTablePerson.Row
            'Количество удалений/добавлений строк в "Таблице персон"
                gTablePerson.Tag = gTablePerson.Tag - 1
            Else
            'Текущая строка "Таблицы персон"
                If gTablePerson.Row < gTablePerson.Rows - 1 Then _
                gTablePerson.Row = gTablePerson.Row + 1
            End If
        Next
            'Освободить объект "Mutex"
        lngRetCode = ReleaseMutex(gMutex)
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
            'Снять запрос на реальное удаление строк из "Таблице персон"
        gRealDelPerson = False
    End If
    
End Function
            
            'Удаление строки из "Таблицы персон"
Private Sub cmdDelete_Click()
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Персональный код в "Таблице информации"
Dim strPersonCode As String
            'Признак успешного удаления
Dim intRealDel As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы персон"
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
            'Очистить текстовые поля
    txtName.Text = ""
    txtPersonCode.Text = ""
    txtReservation.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtType.Text = ""
            'Очистить список имен
    lstName.Clear
    
            'Признак неуспешного удаления
    intRealDel = 1
            
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя Персональный код
    strPersonCode = InputBox("PersonCode: 16 Characters !!!", "Delete ...")
    If Len(Trim(strPersonCode)) > 16 Then strPersonCode = _
    Left(Trim(strPersonCode), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strPersonCode)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strPersonCode = Left("0000000000000000", _
        16 - Len(Trim(strPersonCode))) + Trim(strPersonCode)
    End If
    
            'Персональный код не выбран
    If strPersonCode = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The PersonCode isn't selected"
            
            'Персональный код выбран
    Else
            'Столбец "PersonCode"
        gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
            If Trim(gTablePerson.Text) = strPersonCode Then
            'Поиск выполнен успешно
                intRealDel = 0
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
                intButtonsAndIcons = vbYesNo + vbQuestion
                strResponse = MsgBox("Deletion PersonCode ?", _
                intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
                If strResponse = vbYes Then
                    gTablePerson.Col = 0
                    gProtocol.strProtocName = Trim(gTablePerson.Text)
            'Удаление строки
                    gTablePerson.RemoveItem intRowNum
            'Количество удалений/добавлений строк в "Таблице персон"
                    gTablePerson.Tag = gTablePerson.Tag - 1
            'Установить признак внесенных изменений в "Таблицу персон"
                    gChangesTablePerson = True
            '  Протоколировать данное событие
                    gProtocol.strProtocPersonCode = strPersonCode
                    gProtocol.strProtocStatus = "04 - Manager"
            'Время
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                    gProtocol.strProtocReserve = "PersonCode Deletion"
            'Записать строку в файл "Таблицы протокола"
                    frmDemo.WriteProtocol
                End If
            'Досрочный выход из цикла
                Exit For
            End If
        Next
            
            'ПЕРСОНАЛЬНОГО КОДА нет в "Таблице Персон"
        If intRealDel = 1 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            MsgBox ("Unexistent PersonCode")
        End If
    End If
    
            'Установить фокус на кнопке "Delete"
    If frmTablePerson.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            'Сохранение "Таблицы персон" в файле по умолчанию
Public Function SaveTablePerson()
    Call cmdSave_Click
    SaveTablePerson = 0
    
End Function
            
            'Сохранение "Таблицы персон" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы персон"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы персон"
Dim intColNum As Integer
            
            'Вычислить длину записи (строки) "Таблицы персон"
    lngRecordLen = Len(gPerson)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TablePerson.dat"
    
            'Строк, удаленных из "Таблицы персон" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
    If gTablePerson.Tag < 0 Then
            'Удалить "старый" умалчиваемый файл
        Kill strPathFileName
    End If
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы персон" в файл
            Select Case intColNum
                Case 0
                gPerson.strName = gTablePerson.Text
                Case 1
            'Cохранить информацию о количестве свободных мест:
            '  - на Автостоянке
                If Trim(gPerson.strName) = "ParkFreePlaces" Then
                    gPerson.strPersonCode = Left("000000000000000", 16 - _
                    Len(CStr(gParkFreePlaces))) + CStr(gParkFreePlaces)
                    gTablePerson.Text = gPerson.strPersonCode
            '  - на Предприятии
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
            'Записать строку "Таблицы персон" в файл
        Put intFileNum, intRowNum, gPerson
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице персон"
    gTablePerson.Tag = 0
            'Сбросить признак внесенных изменений в "Таблицу персон"
    gChangesTablePerson = False
            'Установить фокус на кнопке "Cancel"
    If frmTablePerson.Visible = True Then cmdCancel.SetFocus
            
End Sub
            
            'Сохранение "Таблицы персон" в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы персон"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы персон"
Dim intColNum As Integer

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
        frmDemo.BeepSound
        MsgBox "The file isn't selected !"
            'Запись "Таблицы персон" в выбранный файл
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Таблицы персон"
        lngRecordLen = Len(gPerson)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Строк, удаленных из "Таблицы персон" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
        If gTablePerson.Tag < 0 Then
            'Удалить "старый" файл, если он существует
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            'Открыть выбранный файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'По всем столбцам "Таблицы персон"
            For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
                gTablePerson.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы персон" в файл
                Select Case intColNum
                    Case 0
                    gPerson.strName = gTablePerson.Text
                    Case 1
            'Сохранить информацию о количестве свободных мест:
            '  - на Автостоянке
                    If Trim(gPerson.strName) = "ParkFreePlaces" Then
                        gPerson.strPersonCode = Left("000000000000000", 16 - _
                        Len(CStr(gParkFreePlaces))) + CStr(gParkFreePlaces)
                        gTablePerson.Text = gPerson.strPersonCode
            '  - на Предприятии
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
            'Записать строку "Таблицы персон" в файл
            Put intFileNum, intRowNum, gPerson
        Next
            'Закрыть выбранный файл
        Close intFileNum
             'Количество удалений/добавлений строк в "Таблице персон"
        gTablePerson.Tag = 0
               'Сбросить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = False
    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Установить фокус на кнопке "Cancel"
    If frmTablePerson.Visible = True Then cmdCancel.SetFocus
    
End Sub

            'Загрузка формы "Таблица персон"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы персон"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы персон"
Dim intColNum As Integer

            'Установить ширину столбцов
    SetColWidth
            'Текущая строка = 0 (Заголовки столбцов)
    grdTablePerson.Row = 0
    grdTablePerson.Col = 0
    grdTablePerson.Text = "Name"
            'Записать в ячейку (строка 0, столбец 1)
    grdTablePerson.Col = 1
    grdTablePerson.Text = "PersonCode"
            'Записать в ячейку (строка 0, столбец 2)
    grdTablePerson.Col = 2
    grdTablePerson.Text = "Status"
            'Записать в ячейку (строка 0, столбец 3)
    grdTablePerson.Col = 3
    grdTablePerson.Text = "Time"
            'Записать в ячейку (строка 0, столбец 4)
    grdTablePerson.Col = 4
    grdTablePerson.Text = "Calendar"
            'Записать в ячейку (строка 0, столбец 5)
    grdTablePerson.Col = 5
    grdTablePerson.Text = "Reservation"
    
            
            'Заполнение "Таблицы персон" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы персон"
    lngRecordLen = Len(gPerson)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TablePerson.dat"
                
            'Файл отсутствует - ?
    On Error GoTo ErrorTablePerson
                'Количество строк "Таблицы персон" равно размеру файла по умолчанию +1
    grdTablePerson.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To grdTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        grdTablePerson.Row = intRowNum
            'Читать строку "Таблицы персон" из файла в буфер
        Get intFileNum, intRowNum, gPerson
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To grdTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            grdTablePerson.Col = intColNum
            'Заполнение текущей строки "Таблицы персон" из буфера
            Select Case intColNum
                Case 0
                grdTablePerson.Text = gPerson.strName
            'Признак необходимости сжатия "Таблицы персон":
            '   устанавливается всегда в "Host Computer'e" и в тех случаях
            '   в "Препроцессоре", когда последний использует свою
            '   собственную "Таблицу персон" - "ЗЕРКАЛЬНАЯ Таблицa персон" -
            '   обновить информацию о количестве свободных мест:
                If gCompresTablPers = 1 Then
                    If Trim(gPerson.strName) = "ParkFreePlaces" Then _
            'На Автостоянке
                        gParkFreePlaces = gPerson.strPersonCode
                    ElseIf Trim(gPerson.strName) = "AcceFreePlaces" Then _
            'На Предприятии
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
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице персон"
    grdTablePerson.Tag = 0
            'Сбросить признак внесенных изменений в "Таблицу персон"
    gChangesTablePerson = False
            'Снять запрос на физическое удаление строк из "Таблицы персон"
    gRealDelPerson = False
    
    Exit Sub
ErrorTablePerson:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TablePerson Error !")
            'Количество удалений/добавлений строк в "Таблице персон"
    grdTablePerson.Tag = 0
    
End Sub

            'Процедура анализа ячейки "Reserve" текущей строки "Таблицы персон"
            '   при въезде/выезде Постоянных и Бесплатных Клиентов Автостоянки
Public Function AnalysisParking(ByVal vntWork As Variant)
           'Статус
Dim strStatus As String
            'Упакованная ячейка "Reserve" или Подстрока "Контроль" поля
            '   "Информация" текущей строки "Таблицы персон"
Dim strChecking As String * 8
            'Распакованная ячейка "Reserve" или Подстрока "Контроль" поля
            '  "Информация" текущей строки "Таблицы персон"
Dim strCheckingUnPack As String
            'Рабочая переменная
Dim strDate As String
            'Время исключения Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Клиента
Dim strHour As String
Dim strMinute As String

            'Установить нулевой код возврата - доступ через терминал разрешен
    AnalysisParking = 0
            'Текущий столбец "Таблицы персон" = 2 (Статус)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
    gTablePerson.Col = 5
    strChecking = Trim(gTablePerson.Text)
            'Проверка корректного сочетания Номера считывателя и
            '   направления въезда/выезда Автомобиля
    If (vntWork = 0 And (Mid(Trim(strChecking), 7, 1) = "1" Or _
    (Mid(Trim(strChecking), 8, 1) <> "E" And Mid(Trim(strChecking), 7, 1) = "2"))) Or _
    (vntWork = 1 And (Mid(Trim(strChecking), 7, 1) = "0" Or _
    Mid(Trim(strChecking), 7, 1) = "2")) Then
            
            'Анализ статуса Клиента Автостоянки
            
            'Недопустимый для Автостоянки статус Клиента
        If Left(Trim(strStatus), 2) <> "05" And Left(Trim(strStatus), 2) <> "06" And _
        Left(Trim(strStatus), 2) <> "07" Then
            GoTo AnalysisError
        End If
            'Бесплатный Клиент - Доступ через терминал разрешен без оплаты
        If Left(Trim(strStatus), 2) = "07" Then Exit Function
            
            'Распаковка строки "Контроль"
        Call UnPack(strDate, strChecking)
            
            'Формирование распакованной строки "Контроль"
        strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/въезда/выезда Клиента и Резерв
        strCheckingUnPack = Trim(strCheckingUnPack) + Mid(Trim(strChecking), 7, 2)

            'Дата исключения Клиента
        strDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(strDate)
            
            'Последний оплаченный день парковки еще не наступил
        If ((CInt(Mid(strCheckingUnPack, 7, 4)) = CInt(Right(strDate, 4))) And _
        (CInt(Mid(strCheckingUnPack, 4, 2)) > CInt(Mid(strDate, 4, 2))) Or _
        (CInt(Mid(strCheckingUnPack, 4, 2)) = _
        CInt(Mid(strDate, 4, 2)) And CInt(Left(strCheckingUnPack, 2)) >= _
        CInt(Left(strDate, 2)))) Or _
        (CInt(Mid(strCheckingUnPack, 7, 4)) > CInt(Right(strDate, 4))) Then
            'Постоянный Клиент с оплаченным въездом/выездом
            If Left(Trim(strStatus), 2) = "05" Then
            'Автостоянка с ограничением времени непрерывного пребывания
            '  и АМ выезжает с автостоянки
                If gParkTimeLimit > 0 And vntWork = 1 Then
            'Текущий столбец "Таблицы персон" = 0 (Информация)
                    gTablePerson.Col = 0
                    strChecking = Left(Trim(gTablePerson.Text), 6)
            'Распакованная подстрока "Контроль" поля "Информация"
                    Call UnPack(strDate, strChecking)
            'Формирование распакованной подстроки "Контроль"
                    strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
                    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Вычислить время выезда Клиента
                    strDate = Format(Now, "h:mm:ss")
            'Часы
                    intHour = Hour(strDate)
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            'Минуты
                    intMinute = Minute(strDate)
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
            'Последние оплаченные час и минута еще не наступили
                    If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
                    (CInt(Mid(strCheckingUnPack, 12, 2)) = _
                    CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
                    CInt(strMinute)) Then
                        Exit Function
                    End If
            'Установить код возврата = 2 (нужна доплата)
                    AnalysisParking = 2
                End If
                Exit Function
            'Временный Клиент
            Else
                GoTo Continue
            End If
            'Последний оплаченный день истек
        Else
            GoTo AnalysisError
        End If
Continue:
            
            'Временный Клиент
        If Left(Trim(strStatus), 2) = "06" Then
            'Вычислить время исключения Клиента
            strDate = Format(Now, "h:mm:ss")
            'Часы
           intHour = Hour(strDate)
            If intHour < 10 Then
                strHour = "0" + Trim(Str(intHour))
            Else
                strHour = Trim(Str(intHour))
            End If
            'Минуты
            intMinute = Minute(strDate)
            If intMinute < 10 Then
                strMinute = "0" + Trim(Str(intMinute))
            Else
                strMinute = Trim(Str(intMinute))
            End If
            
            'Последние оплаченные час и минута еще не наступили
            If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
            (CInt(Mid(strCheckingUnPack, 12, 2)) = _
            CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
            CInt(strMinute)) Then
                Exit Function
            End If
        End If
            
    End If
    
            'Неверный Статус Клиента или некорректные направление
            '  либо время въезда/выезда
AnalysisError:
            'Установить не нулевой код возврата - доступ через терминал не разрешен
    AnalysisParking = 1

End Function

            'Процедура анализа ячейки "Reserve" текущей строки "Таблицы персон"
            '   при входе/выходе Постоянных и Бесплатных Клиентов
Public Function AnalysisAccess(ByVal vntWork As Variant)
           'Статус
Dim strStatus As String
            'Упакованная ячейка "Reserve" или Подстрока "Контроль" поля
            '   "Информация" текущей строки "Таблицы персон"
Dim strChecking As String * 8
            'Распакованная ячейка "Reserve" или Подстрока "Контроль" поля
            '  "Информация" текущей строки "Таблицы персон"
Dim strCheckingUnPack As String
            'Рабочая переменная
Dim strDate As String
            'Время исключения Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Клиента
Dim strHour As String
Dim strMinute As String

            'Установить нулевой код возврата - доступ через терминал разрешен
    AnalysisAccess = 0
            'Текущий столбец "Таблицы персон" = 2 (Статус)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 5 (Резерв)
    gTablePerson.Col = 5
    strChecking = Trim(gTablePerson.Text)
            'Проверка корректного сочетания Номера считывателя и
            '   направления входа/выхода Посетителя
    If (vntWork = 0 And (Mid(Trim(strChecking), 7, 1) = "1" Or _
    (Mid(Trim(strChecking), 8, 1) <> "E" And Mid(Trim(strChecking), 7, 1) = "2"))) Or _
    (vntWork = 1 And (Mid(Trim(strChecking), 7, 1) = "0" Or _
    Mid(Trim(strChecking), 7, 1) = "2")) Then
            
            'Анализ статуса Клиента
            
            'Недопустимый для Предприятия статус Клиента
        If Left(Trim(strStatus), 2) <> "08" And Left(Trim(strStatus), 2) <> "09" And _
        Left(Trim(strStatus), 2) <> "10" Then
            GoTo AnalysisError
        End If
            'Бесплатный Клиент - Доступ через терминал разрешен без оплаты
        If Left(Trim(strStatus), 2) = "10" Then Exit Function
            
            'Распаковка строки "Контроль"
        Call UnPack(strDate, strChecking)
        
            'Формирование распакованной строки "Контроль"
        strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/въезда/выезда Клиента и Резерв
        strCheckingUnPack = Trim(strCheckingUnPack) + Mid(Trim(strChecking), 7, 2)

            'Дата исключения Клиента
        strDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(strDate)
            
            'Последний оплаченный день парковки еще не наступил
        If ((CInt(Mid(strCheckingUnPack, 7, 4)) = CInt(Right(strDate, 4))) And _
        (CInt(Mid(strCheckingUnPack, 4, 2)) > CInt(Mid(strDate, 4, 2))) Or _
        (CInt(Mid(strCheckingUnPack, 4, 2)) = _
        CInt(Mid(strDate, 4, 2)) And CInt(Left(strCheckingUnPack, 2)) >= _
        CInt(Left(strDate, 2)))) Or _
        (CInt(Mid(strCheckingUnPack, 7, 4)) > CInt(Right(strDate, 4))) Then
            'Постоянный Клиент с оплаченным Входом-Выходом
            If Left(Trim(strStatus), 2) = "08" Then
            'Предприятие с ограничением времени непрерывного пребывания
            '  и Клиент выходит с предприятия
                If gAcceTimeLimit > 0 And vntWork = 1 Then
            'Текущий столбец "Таблицы персон" = 0 (Информация)
                    gTablePerson.Col = 0
                    strChecking = Left(Trim(gTablePerson.Text), 6)
            'Распакованная подстрока "Контроль" поля "Информация"
                    Call UnPack(strDate, strChecking)
            'Формирование распакованной подстроки "Контроль"
                    strCheckingUnPack = Left(Trim(strDate), 2) + "." + _
                    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Вычислить время выхода Клиента
                    strDate = Format(Now, "h:mm:ss")
            'Часы
                    intHour = Hour(strDate)
                    If intHour < 10 Then
                        strHour = "0" + Trim(Str(intHour))
                    Else
                        strHour = Trim(Str(intHour))
                    End If
            'Минуты
                    intMinute = Minute(strDate)
                    If intMinute < 10 Then
                        strMinute = "0" + Trim(Str(intMinute))
                    Else
                        strMinute = Trim(Str(intMinute))
                    End If
            'Последние оплаченные час и минута еще не наступили
                    If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
                    (CInt(Mid(strCheckingUnPack, 12, 2)) = _
                    CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
                    CInt(strMinute)) Then
                        Exit Function
                    End If
            'Установить код возврата = 2 (нужна доплата)
                    AnalysisAccess = 2
                End If
                Exit Function
            'Временный Клиент
            Else
                GoTo Continue
            End If
            'Последний оплаченный день истек
        Else
            GoTo AnalysisError
        End If
Continue:
            
            'Временный Клиент
        If Left(Trim(strStatus), 2) = "09" Then
            'Вычислить время исключения Клиента
            strDate = Format(Now, "h:mm:ss")
            'Часы
           intHour = Hour(strDate)
            If intHour < 10 Then
                strHour = "0" + Trim(Str(intHour))
            Else
                strHour = Trim(Str(intHour))
            End If
            'Минуты
            intMinute = Minute(strDate)
            If intMinute < 10 Then
                strMinute = "0" + Trim(Str(intMinute))
            Else
                strMinute = Trim(Str(intMinute))
            End If
            
            'Последние оплаченные час и минута еще не наступили
            If (CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
            (CInt(Mid(strCheckingUnPack, 12, 2)) = _
            CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
            CInt(strMinute)) Then
            'Номер считывателя="1" - Посетитель выходит с Предприятия
                If vntWork = 1 Then
            'Текущий столбец "Таблицы персон" = 0 (Информация)
                    gTablePerson.Col = 0
            'Инвентарь не брался напрокат или возвращен
                    If Left(Trim(gTablePerson.Text), 4) = "0000" Then Exit Function
            'Номер считывателя="0" - Посетитель входит на Предприятие
                Else
                    Exit Function
                End If
            End If
        End If
            
    End If
    
            'Неверный Статус Клиента или некорректные направление
            '  либо время входа/выхода, либо не возвращен прокатный инвентарь
AnalysisError:
            'Установить не нулевой код возврата - доступ через терминал не разрешен
    AnalysisAccess = 1

End Function

            'Процедура анализа ячейки "Name" текущей строки "Таблицы персон"
            '   при входе/выходе Служащих
Public Function AnalysisEmploye(ByVal vntWork As Variant)
Dim strStatus As String

            'Установить нулевой код возврата - доступ через терминал разрешен
    AnalysisEmploye = 0
            'Текущий столбец "Таблицы персон" = 2 (Статус)
    gTablePerson.Col = 2
    strStatus = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы персон" = 0 (Имя)
    gTablePerson.Col = 0
            'Проверка корректного сочетания Номера считывателя и
            '   направления входа/выхода Служащего
    If (vntWork = 0 And Right(Trim(gTablePerson.Text), 1) = "-") Or _
    (vntWork = 1 And Right(Trim(gTablePerson.Text), 1) = "+") Or _
    Len(Trim(gTablePerson.Text)) < 16 Then
            
            'Анализ статуса Служащего
            
            'Недопустимый статус
        If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            GoTo AnalysisError
        End If
        Exit Function
    End If
    
            'Неверный Статус Служащего или некорректное направление входа/выхода
AnalysisError:
            'Установить не нулевой код возврата - доступ через терминал не разрешен
    AnalysisEmploye = 1

End Function
            
            'Распаковка строки "Контроль"
Public Sub UnPack(ByRef strDate As String, ByVal strChecking As String)
             'Номер позиции спецсимвола в анализируемом поле
Dim intPosNum As Integer
            'Рабочий счетчик
Dim intCount As Integer
        
    strDate = ""
            'Поиск символa "z" (7AH) и замена его
            '  двойным нулем (Упаковыванное в Пусто - "")
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "z")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("00"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск символa "x" (78H) и замена его
            '  на "09" (Упаковыванное Пусто - "")
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "x")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("09"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск символa "y" (79H) и замена его
            '  на "10" (Упаковыванный Перевод Каретки)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "y")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("10"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск символa "w" (77H) и замена его
            '  на "13" (Упаковыванный Перевод Каретки)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "w")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("13"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск символa "r" (72H) и замена его
            '  на "32" (Упаковыванный Перевод Каретки)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, "r")
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + Chr(CByte(CInt("32"))) + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Дата и Время
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
            
            'Упаковка строки "Контроль"
Public Sub Pack(ByRef strChecking As String)
             'Номер позиции спецсимвола в анализируемом поле
Dim intPosNum As Integer
            'Рабочий счетчик
Dim intCount As Integer
            
            'Поиск двойного нуля (Упакованного в Пусто - "")
            '  и замена его символом "z" (7AH)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("00"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "z" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск "09" (Упакованного в Пусто - "")
            '  и замена его символом "x" (78H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("09"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "x" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск "10" (Упакованного в Перевод Каретки)
            '  и замена его символом "y" (79H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("10"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "y" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск "13" (Упакованного в Перевод Каретки)
            '  и замена его символом "w" (77H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("13"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "w" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next
            'Поиск "32" (Упакованного Пусто - "")
            '  и замена его символом "r" (72H)
    For intCount = 1 To 3
        intPosNum = InStr(1, strChecking, Chr(CByte(CInt("32"))))
        If intPosNum <> 0 Then
            strChecking = Left(strChecking, intPosNum - 1) + "r" + _
            Right(strChecking, 8 - intPosNum)
        End If
    Next

End Sub
            
            'Добавление строки с заданным полями в "Таблицу персон"
            '  по сообщению MSMQ, полученому по сети
Public Function MSMQReg(ByVal strMessage As String)
            'Номер текущей строки в "Таблице персон"
Dim intRowNum As Integer
            'Пустая строка "Таблицы персон"
Dim strPerson As String
            'Поле ячейки "Person or Terminal" в "Таблице персон"
Dim strName As String
            'Поле ячейки "PersonCode" в "Таблице персон"
Dim vntPersonCode As Variant
            'Поле ячейки "Status" в "Таблице персон"
Dim strStatus As String
            'Поле ячейки "Time" в "Таблице персон"
Dim strTime As String
            'Поле ячейки "Calendar" в "Таблице персон"
Dim strCalendar As String
            'Поле ячейки "Reserve" в "Таблице персон"
Dim strReserve As String
            'Номер анализируемой записи в строке сообщения
Dim intNumber As Integer
            
            'Признак необходимости сжатия "Таблицы персон" не установлен:
            '   "Препроцессор" использует "Таблицу персон" "Host Computer'а"
            '   - выйти из процедуры
    If gCompresTablPers = 0 Then Exit Function
    
            'Номер анализируемой записи в строке сообщения
    intNumber = 1
            'Искать в строке сообщения символы "07H" - разделители записей
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            'Подготовка ячейки "Person or Terminal" для "Таблице персон"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            'Подготовка ячейки "PersonCode" для "Таблице персон"
            vntPersonCode = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            'Подготовка ячейки "Status" для "Таблице персон"
            strStatus = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            'Подготовка ячейки "Time" для "Таблице персон"
            strTime = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            'Подготовка ячейки "Calendar" для "Таблице персон"
            strCalendar = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            'Подготовка ячейки "Reserve" для "Таблице персон"
            strReserve = strMessage
            'Принудительный выход, т.к. ячейка "Reserve" может включать "07H"
            Exit Do
        End If
    Loop
        
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице персон"
    If intRowNum = gTablePerson.Rows Then
            'Добавление строки в конец "Таблицы персон"
        gTablePerson.AddItem strPerson
        gTablePerson.Row = gTablePerson.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            'Изменение ячейки "PersonCode" в "Таблице персон"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            'Изменение ячейки "Status" в "Таблице персон"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            'Изменение ячейки "Time" в "Таблице персон"
        gTablePerson.Col = 3
        gTablePerson.Text = strTime
            'Изменение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Col = 4
        gTablePerson.Text = strCalendar
            'Изменение ячейки "Reserve" в "Таблице персон"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            
    
            'Количество удалений/добавлений строк в "Таблице персон"
        gTablePerson.Tag = gTablePerson.Tag + 1
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    
    End If
    
End Function
            
            'Удаление (ЛОГИЧЕСКОЕ) строки с заданным персональным кодом
            '  из "Таблицы персон" по сообщению MSMQ, полученому по сети
Public Function MSMQDel(ByVal vntPersonCode As Variant)
            'Текущий номер нефиксированной строки "Таблицы персон"
Dim intRowNum As Integer
            
            'Признак необходимости сжатия "Таблицы персон" не установлен:
            '   "Препроцессор" использует "Таблицу персон" "Host Computer'а"
            '   - выйти из процедуры
    If gCompresTablPers = 0 Then Exit Function
        
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
    
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    If intRowNum < gTablePerson.Rows Then
            'Логически удалить строку из "Таблицы персон"
        gTablePerson.Text = "Deleted"
            'Установить запрос на реальное удаление
        gRealDelPerson = True
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    End If
    
End Function
            
            'Коррекция заданных полей строки "Таблицы персон"
            '  по сообщению MSMQ, полученому по сети
Public Function MSMQCor(ByVal strMessage As String)
            'Номер текущей строки "Таблицы персон"
Dim intRowNum As Integer
            'Пустая строка "Таблицы персон"
Dim strPerson As String
            'Поле ячейки "Person or Terminal" в "Таблице персон"
Dim strName As String
            'Поле ячейки "PersonCode" в "Таблице персон"
Dim vntPersonCode As Variant
            'Поле ячейки "Status" в "Таблице персон"
Dim strStatus As String
            'Поле ячейки "Time" в "Таблице персон"
Dim strTime As String
            'Поле ячейки "Calendar" в "Таблице персон"
Dim strCalendar As String
            'Поле ячейки "Reserve" в "Таблице персон"
Dim strReserve As String
            'Номер анализируемой записи в строке сообщения
Dim intNumber As Integer
            
            'Признак необходимости сжатия "Таблицы персон" не установлен:
            '   "Препроцессор" использует "Таблицу персон" "Host Computer'а"
            '   - выйти из процедуры
    If gCompresTablPers = 0 Then Exit Function
    
            'Номер анализируемой записи в строке сообщения
    intNumber = 1
            'Искать в строке сообщения символы "07H" - разделители записей
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            'Подготовка ячейки "Person or Terminal" для "Таблице персон"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            'Подготовка ячейки "PersonCode" для "Таблице персон"
            vntPersonCode = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            'Подготовка ячейки "Status" для "Таблице персон"
            strStatus = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            'Подготовка ячейки "Time" для "Таблице персон"
            strTime = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            'Подготовка ячейки "Calendar" для "Таблице персон"
            strCalendar = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            'Подготовка ячейки "Reserve" для "Таблице персон"
            strReserve = strMessage
            'Принудительный выход, т.к. ячейка "Reserve" может включать "07H"
            Exit Do
        End If
    Loop
        
        'Текущий столбец "Таблицы персон" = 1 (Персональный код)
    gTablePerson.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
        If Trim(gTablePerson.Text) = vntPersonCode Then
            'Досрочный выход из цикла
            Exit For
        End If
    Next
    
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице персон"
    If intRowNum < gTablePerson.Rows Then
        gTablePerson.Row = intRowNum
            'Изменение ячейки "Person or Terminal" в "Таблице персон"
        gTablePerson.Col = 0
        gTablePerson.Text = strName
            'Изменение ячейки "PersonCode" в "Таблице персон"
        gTablePerson.Col = 1
        gTablePerson.Text = Trim(vntPersonCode)
            'Изменение ячейки "Status" в "Таблице персон"
        gTablePerson.Col = 2
        gTablePerson.Text = strStatus
            'Изменение ячейки "Time" в "Таблице персон"
        gTablePerson.Col = 3
        gTablePerson.Text = strTime
            'Изменение ячейки "Calendar" в "Таблице персон"
        gTablePerson.Col = 4
        gTablePerson.Text = strCalendar
            'Изменение ячейки "Reserve" в "Таблице персон"
        gTablePerson.Col = 5
        gTablePerson.Text = strReserve
            
            'Установить признак внесенных изменений в "Таблицу персон"
        gChangesTablePerson = True
    
    End If
    
End Function

            'Процедура установки ширины и выравнивания столбцов "Таблицы персон"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Цикл по всем столбцам
    For intColNumber = 0 To grdTablePerson.Cols - 1 Step 1
        grdTablePerson.ColWidth(intColNumber) = 1650
        grdTablePerson.ColAlignment(intColNumber) = 0
    Next
            'Увеличение размера 0-го и 1-го столбцов (Имя и Персональный код)
    intColNumber = 0
    grdTablePerson.ColWidth(intColNumber) = 2500
    intColNumber = 1
    grdTablePerson.ColWidth(intColNumber) = 2500
    
End Sub

