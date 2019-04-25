VERSION 5.00
Begin VB.Form frmDataAccessIn 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessInData"
   ClientHeight    =   3945
   ClientLeft      =   4485
   ClientTop       =   2925
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
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
   ScaleHeight     =   3945
   ScaleWidth      =   6990
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Frame fraPeople 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton optFamily 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optConvoy 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optBaby 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optHuman 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   19
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   18
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   17
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   99
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   320
      TabIndex        =   15
      Top             =   1920
      Width           =   1452
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraStatus 
      Caption         =   "????"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton optTime 
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
         Left            =   720
         TabIndex        =   12
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optCalendar 
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
         Left            =   720
         TabIndex        =   11
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optMoneyFree 
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
         Height          =   252
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Frame fraDayNight 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
         Begin VB.OptionButton optDay 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image imgCalendar 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image imgTime 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":24A4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
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
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   2040
      Picture         =   "frmDataAccessIn.frx":28FE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   720
      Picture         =   "frmDataAccessIn.frx":2F30
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   120
      Picture         =   "frmDataAccessIn.frx":356E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   1200
      Picture         =   "frmDataAccessIn.frx":3DA8
      Stretch         =   -1  'True
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line7 
      X1              =   2760
      X2              =   4080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblPersonCode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "#### "
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
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Info "
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
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblLat320 
      Alignment       =   2  'Center
      Caption         =   "320"
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
      Height          =   255
      Left            =   6360
      TabIndex        =   23
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label lblLat0 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label lblParole 
      Alignment       =   2  'Center
      Caption         =   "Parole"
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
      Left            =   4560
      TabIndex        =   21
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessIn.frx":4B2E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2280
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6840
      X2              =   4080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   720
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1680
      Y2              =   3720
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Image imgAccessIn 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessIn.frx":4F44
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Ls"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   1920
      Width           =   375
   End
End
Attribute VB_Name = "frmDataAccessIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Вносимая сумма оплаты в Сантимах
Dim intAccessMoney As Integer
            'Количество дней доступа
Dim intAccessDay As Integer
            'Тариф одного дня (Сутки)
Dim intAccessTariffFull As Integer
            'Тариф одного дня (День)
Dim intAccessTariffDay As Integer
            'Тариф одного дня (Ночь)
Dim intAccessTariffNight As Integer
            'Тариф (переменная для рассчетов)
Dim intAccessTariff As Integer
            'Признак времени допуска - (для Постоянных Посетителей)
Dim strTime As String
            'Текущая строка "Таблицы календаря"
Dim intRowNum As Integer
            'Текущая столбец "Таблицы календаря"
Dim intColNum As Integer
            'День, соответствующий Дате Регистрации
            '  Клиента Автостоянки (или последнему парковочному дню)
Dim intDayReg As Integer
            'Месяц, соответствующий Дате Регистрации
            '  Клиента Автостоянки (или последнему парковочному дню)
Dim intMonthReg As Integer
            'Год, соответствующий Дате Регистрации
            '  Клиента Автостоянки (или последнему парковочному дню)
Dim intYearReg As Integer
            'Номер позиции заданного символа в строке
Dim intPosNum As Integer
             'Введенный пароль
Dim strPassword As String

            'Перехват нажатия комбинаций клавиш "Alt"+ {"+" и "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Текущая форма доступна
    If Me.Enabled = True Then
            'Альтернатива "щелчку" мыши на кнопке "+"
        If KeyCode = 187 And Shift = 4 Then
            If cmdOK.Enabled = True Then
                Call cmdOK_Click
                Exit Sub
            End If
            'Альтернатива "щелчку" мыши на кнопке "Cancel"
        ElseIf KeyCode = 69 And Shift = 4 Then
            If cmdCancel.Enabled = True Then
                Call cmdCancel_Click
                Exit Sub
            End If
        End If
    End If
    
End Sub

            'Переключить признак печати документа - "Document"
Private Sub chkDocument_Click()
            'Сделать доступным нажатие на кнопку "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub

            'Возврат в вызвавшую процедуру (Кнопка "OK _ +")
Private Sub cmdOK_Click()
            'Статус
Dim strStatus As String
            'Строка "Контроль"
Dim strChecking As String * 8
            'Подстрока "Контроль" поля "txtInfo"
Dim strCheckingInfo As String * 8
            'Дата (и Время) регистрации Посетителя или
            '  дата последнего оплаченного дня
Dim strDate As String
            'Подстрока "Контрольные Дата и Время" поля "txtInfo"
Dim strDateInfo As String
            'Время регистрации Посетителя
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Посетителя
Dim strHour As String
Dim strMinute As String
            'Признак присутствия \ 0 - вошел \ 1 - вышел \ 2 - зарегистрирован
Dim strPersPresent As String * 1
            'Признак ("Е" - Окончательно вышел; "D" - Дневной тариф допуска;
            '  "N" - Ночной тариф допуска; "Другой символ"   - Суточный тариф
            '  допуска)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоРегистрации в "Таблице персон"
Dim intAutoRegistrCode  As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Стоимость ПРОКАТА ИНВЕНТАРЯ # 1, 2, 3 и 4
Dim intLease As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Количество ячеек времени, в течение которого разрешается Временному
            '  или Постоянному Посетителям непрерывно находиться на Предприятии
Dim intCellLimit As Integer
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer

            'Недоступное нажатие на кнопку "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            
            'Если Посетитель Временный
    If optTime.Value = True Then
    
            'Обнулить поле "Tag" формы "frmLease"
        frmLease.Tag = 0
            'Вывести на экран форму "frmLease" с уровнем модальности 1
        frmLease.Show 1
            
            'Сделать недоступными кнопки "OK _ +" и "Cancel _ Exit"
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
            
            'Коррекция в поле "txtInfo" информации о ПРОКАТЕ ИНВЕНТАРЯ
        If frmLease.Tag <> "Exit" Then
            'Стоимость ПРОКАТА ИНВЕНТАРЯ # 1, 2, 3 и 4
            intLease = 0
            If Mid(CStr(frmLease.Tag), 1, 1) = "+" Then intLease = gLease1
            If Mid(CStr(frmLease.Tag), 2, 1) = "+" Then intLease = intLease + gLease2
            If Mid(CStr(frmLease.Tag), 3, 1) = "+" Then intLease = intLease + gLease3
            If Mid(CStr(frmLease.Tag), 4, 1) = "+" Then intLease = intLease + gLease4
            intLease = intLease + CInt(Left(txtMoneyDate.Text, 3)) * 100 + _
            CInt(Mid(txtMoneyDate.Text, 5, 2))
            'Формирование шаблона в поле "ДеньгиДата"
            txtMoneyDate.Text = "000,00" + Mid(txtMoneyDate.Text, 7)
            'Изменение текстового поля "ДеньгиДаты"
            If Int(intLease / 100) < 10 Then
                txtMoneyDate.Text = "00" + Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            ElseIf Int(intLease / 100) < 100 Then
                txtMoneyDate.Text = "0" + Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            ElseIf Int(intLease / 100) > 99 Then
                txtMoneyDate.Text = Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            End If
            'Изменение текстового поля "ДеньгиДаты"
            If intLease - Int(intLease / 100) * 100 < 10 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
                Trim(Str(intLease - Int(intLease / 100) * 100)) + _
                Mid(txtMoneyDate.Text, 7)
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
                Trim(Str(intLease - Int(intLease / 100) * 100)) + _
                Mid(txtMoneyDate.Text, 7)
            End If
        End If
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'Окно собщения с повторным запросом регистрации
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Addition PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Papild. person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
            Me.Tag = 2
            'Сделать доступными кнопки "OK _ +" и "Cancel _ Exit"
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            'Восстановить стандартный курсор мыши
            Me.MousePointer = 0
            'Установить фокус на кнопке "Cancel _ Exit"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
    End If
            
            'Не нулевая сумма требуемой оплаты
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            'Обнулить поле "Tag" формы "frmMinus"
        frmMinus.Tag = 0
            'Вывести на экран форму "frmMinus" с уровнем модальности 1
        frmMinus.Show 1
            'Отказ от оплаты и от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
        If frmMinus.Tag = "Exit" Then
            'Возврат в вызвавшую процедуру
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            'Очистка строки и подстроки "Контроль"
    strChecking = ""
    strCheckingInfo = ""
            'Вычислить время регистрации Посетителя
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
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
            'Дата регистрации Посетителя
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
            'Признак регистрации Посетителя
    strPersPresent = "2"
            'Признак Посетителя
    strExpander = "P"
            'Анализ статуса Посетителя
    If optMoneyFree.Value = True Then
            'Бесплатный Посетитель
        strStatus = "10 - Access/Free"
    ElseIf optCalendar.Value = True Then
            'Постоянный Посетитель
        strStatus = "08 - Access/Calen."
            'Посетитель с Дневным тарифом допуска
        If optCalendar.Value = True And optDay.Value = True Then
            strExpander = "D"
            'Посетитель с Ночным тарифом допуска
        ElseIf optCalendar.Value = True And optNight.Value = True Then
            strExpander = "N"
        End If
            'Дата последнего оплаченного дня
        strDate = Mid(Trim(txtMoneyDate.Text), 11)
        If Len(Trim(strDate)) = 10 Then
            strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
            Right(Trim(strDate), 4)
        Else
            strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
            Right(Trim(strDate), 4)
        End If
    ElseIf optTime.Value = True Then
            'Временный Посетитель
        strStatus = "09 - Access/Time"
            'Количество ячеек времени, в течение которого разрешается Временному
            '  Посетителю находиться на Предприятии (предоплата входа)
        intCellLimit = gAcceInpCellNumb
    
            'Вычислить "сдвинутые" время и дату регистрации Посетителя,
            '  до которых ему будет разрешен Вход-Выход
            
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
                    strDate = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                Else
                    strDate = "0" + _
                    Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                End If
            'Изменение  Месяца и, возможно, Года
                If frmTableCalendar.comCalendar.Day = 1 Then
                    If frmTableCalendar.comCalendar.Month > 9 Then
                        strDate = "01" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    Else
                        strDate = "010" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    End If
                End If
            End If
            
            'Не требуется переход часа
        Else
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
    
    End If
            
            'Формирование упакованной строки "Контроль"
    For intCount = 1 To 7 Step 2
            'Дата
        strChecking = Trim(strChecking) + _
        Chr(CByte(CInt(Mid(strDate, intCount, 2))))
    Next
            'Часы
    strChecking = Trim(strChecking) + _
    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            'Минуты
    strChecking = Trim(strChecking) + _
    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            'Упаковка строки "Контроль"
    Call frmTablePerson.Pack(strChecking)
            
            'Признак регистрации Посетителя и Резерв для расширения
    strChecking = Left(strChecking, 6) + strPersPresent + strExpander
            
            'Если Посетитель Временный
    If optTime.Value = True Then
            'Коррекция в поле "txtInfo" информации о ПРОКАТЕ ИНВЕНТАРЯ
        If frmLease.Tag <> "Exit" Then _
        txtInfo.Text = Left(CStr(frmLease.Tag), 4) + Mid(txtInfo.Text, 5)
    End If
            'Коррекция в поле "txtInfo" информации о Клиенте (Взрослый)
    If optHuman.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "1" + Mid(txtInfo.Text, 6)
            'Коррекция в поле "txtInfo" информации о Клиенте (Дети)
    ElseIf optBaby.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "2" + Mid(txtInfo.Text, 6)
            'Коррекция в поле "txtInfo" информации о Клиенте (Конвой)
    ElseIf optConvoy.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "3" + Mid(txtInfo.Text, 6)
            'Коррекция в поле "txtInfo" информации о Клиенте (Семья)
    ElseIf optFamily.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "4" + Mid(txtInfo.Text, 6)
    End If
            
            'Постоянный Посетитель на Предприятии с ограничением времени
            '  непрерывного пребывания
    If gAcceTimeLimit > 0 And optCalendar.Value = True Then
            'Количество ячеек времени, в течение которого разрешается Постоянному
            '  Посетителю непрерывно находиться на Предприятии
        intCellLimit = gAccessCellLimit
            
            'Вычислить "сдвинутые" время и дату регистрации Постоянного
            '  Посетителя, до которых ему будет разрешен бесплатный выход
        
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
                    strDate = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                Else
                    strDate = "0" + _
                    Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                End If
            'Изменение  Месяца и, возможно, Года
                If frmTableCalendar.comCalendar.Day = 1 Then
                    If frmTableCalendar.comCalendar.Month > 9 Then
                        strDate = "01" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    Else
                        strDate = "010" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    End If
                End If
            End If
            
            'Не требуется переход часа
        Else
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
            
            'Коррекция поля "txtInfo"
        txtInfo = Left(strCheckingInfo, 6) + Trim(txtInfo)
        
    End If
            
            'Вызов процедуры-функции АвтоРегистрации
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoRegistrCode = frmTablePerson.AutoRegAccess(txtPersonCode.Text, _
    txtInfo.Text, strStatus, strChecking, strTime)
            
            '(Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА выполнена -
            '   протоколирование события
    If intAutoRegistrCode = 0 Then
            'Последняя (новая) строка "Таблицы персон"
        gTablePerson.Row = gTablePerson.Rows - 1
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
        gTablePerson.Col = 0
        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
        gProtocol.strProtocPersonCode = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
        gTablePerson.Col = 2
        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "AutoRegAcce " + Left(Trim(txtMoneyDate.Text), 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            'Признак (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
        Me.Tag = 1
            
            'Опция "Печать Документа" установлена
        If chkDocument.Value = 1 Then
            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
            Call frmDemo.PrintDocument(gProtocol.strProtocName, _
            gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
            gProtocol.strProtocTime, gProtocol.strProtocDate, _
            gProtocol.strProtocReserve, intError)
        End If
        
            'ОТКРЫТЬ ТЕРМИНАЛ ?
            
            'Установлена Опция разделения времени (параллельное выполнение
            '   процедур), режим Выполнение (при режиме Установок параллельное
            '   выполнение процедур невозможно), Временный Клиент Предприятия
            '   и установлен индекс входного терминала - открыть терминал
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        optTime.Value = True And gTermInp <> -1 Then
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  выполнена Регистрация Клиента Предприятия
            If frmDemo.cmdOpen(gTermInp).Tag = 0 And Me.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
                frmDemo.imgAccessInData(gTermInp).Enabled = False
                frmDemo.imgAccessOutData(gTermInp).Enabled = False
                frmDemo.imgAccessInfoData(gTermInp).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(gTermInp))))
                frmDemo.cmdOpen(gTermInp).Tag = vntAddr
                frmDemo.cmdOpen(gTermInp).Caption = "Addr=" + CStr(vntAddr)
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(gTermInp).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(gTermInp).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
                Call frmDemo.OpenBarrier(gTermInp)
            'Вызов процедуры-функции АвтоКоррекции для данного
            'ПЕРСОНАЛЬНОГО КОДА - Клиент вошел на Предприятие
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
            End If
        End If
            
            
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Отказ в АвтоРегистрация ПЕРСОНАЛЬНОГО КОДА -
            '   протоколирование события
    Else
            'Введенная ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'Введенный ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Выбранная опция (Статус)
        gProtocol.strProtocStatus = strStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid AutoRegAccess"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
        Me.Tag = 2
            
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
    End If
            
End Sub
            
            'Возврат в вызвавшую процедуру (Кнопка "Cancel _ Exit")
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            
            'Сделать доступными кнопки "OK _ +" и "Cancel _ Exit"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
            
            'Были не сохраненные изменения в текстовых полях текущей формы
    If Me.Tag = 1 And _
    ((txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
    (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1)) Then
            'Окно собщения с запросом изменения "Таблицы персон" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
            'Издать звуковой сигнал
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" + """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" + """, intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
            Me.Tag = 2
            'Установить фокус на кнопке "Cancel _ Exit"
            cmdCancel.SetFocus
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
    If Me.Tag = 0 Then Me.Tag = 2
            'Сделать невидимой текущую форму
    Me.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Активизация текущей формы
Private Sub Form_Activate()

            'Текущая форма видимая и установлен флаг завершения ее
            '  Активизации - выйти из процедуры (для блокирования возможной
            '  повторной Активизации, чистящей текстовые поля)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            'Увеличить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessPlus
            
            'Сделать доступным текстовое поле ПЕРСОНАЛЬНОГО КОДА
    txtPersonCode.Enabled = True
            'Выбрать опцию "Time"
    optTime.Value = True
            'Выбрать опцию "Human"
    optHuman.Value = True
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
    gAccessMoneyCell = gAccessMoneyCellHuman
            'Сделать доступными некоторые элементы управления формы
    fraPeople.Enabled = True
            'Сделать недоступными элементы управления формы "DataAccessIn"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    fraDayNight.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
            'Тариф одного дня (Сутки)
    intAccessTariff = intAccessTariffFull
            'Признак времени допуска - Суточный (для Постоянных Посетителей)
    strTime = "DayNight"
           'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop
            'Установить фокус на текстовом поле "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
             
             'Установить флаг завершения Активизации текущей формы
    Me.Tag = 1

End Sub

            'Деактивизация текущей формы
Private Sub Form_Deactivate()
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub
            
            'Загрузка текущей формы
Private Sub Form_Load()
            'Сделать недоступными элементы управления формы "DataAccessIn"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    optTime.Value = True
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Тариф одного дня (Сутки)
    intAccessTariffFull = gAccessDN
            'Тариф одного дня (День)
    intAccessTariffDay = gAccessD
            'Тариф одного дня (Ночь)
    intAccessTariffNight = gAccessN
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Процедура обработки "Щелчка мыши" на опции "Human"
Private Sub optHuman_Click()
            
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
    gAccessMoneyCell = gAccessMoneyCellHuman
            
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            
            'Установить фокус на кнопке "OK _ +"
    cmdOK.SetFocus

End Sub

            'Процедура обработки "Щелчка мыши" на опции "Baby"
Private Sub optBaby_Click()
            
            'Входной тариф Предприятия для Детей (для временных Клиентов)
    gAccessMoneyCell = gAccessMoneyCellBaby
            
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            
            'Установить фокус на кнопке "OK _ +"
    cmdOK.SetFocus

End Sub

            'Процедура обработки "Щелчка мыши" на опции "Convoy"
Private Sub optConvoy_Click()
            
            'Входной тариф Предприятия для Конвоя (для временных Клиентов)
    gAccessMoneyCell = gAccessMoneyCellConvoy
            
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            
            'Установить фокус на кнопке "OK _ +"
    cmdOK.SetFocus

End Sub

            'Процедура обработки "Щелчка мыши" на опции "Family"
Private Sub optFamily_Click()
            
            'Входной тариф Предприятия для Семьи (для временных Клиентов)
    gAccessMoneyCell = gAccessMoneyCellFamily
            
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            
            'Установить фокус на кнопке "OK _ +"
    cmdOK.SetFocus

End Sub

            'Процедура обработки "Щелчка мыши" на поле Персонального кода
Private Sub txtPersonCode_Click()
            'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
            'Сбросить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub

            'Процедура ввода и анализа "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtPersonCode.BackColor = vbCyan
            'Переход по ошибке преобразования данных
        On Error GoTo PersonCodeError
            'Персональный код в допустимом диапазоне
        If Len(Trim(txtPersonCode.Text)) > 0 And _
        Len(Trim(txtPersonCode.Text)) < 17 Then
            'Установлена Опция копирования "PersonCode"в поле "Info"
            '  при Регистрации Временного Посетителя
            If gAccessCodeInfo = 1 Then
            'Копирование "PersonCode"в поле "Info"
                txtInfo = Trim(txtPersonCode)
            'Голубой фон текстового поля
                txtInfo.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
                txtInfo.Tag = 1
            End If
            'Длина персонального кода меньше 16-и символов
            If Len(Trim(txtPersonCode.Text)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
                txtPersonCode.Text = Left("0000000000000000", _
                16 - Len(Trim(txtPersonCode.Text))) + Trim(txtPersonCode.Text)
            End If
            'Установить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 1
            'Установить фокус на текстовом поле "txtInfo"
            If txtInfo.Enabled = True Then txtInfo.SetFocus
            'Вся необходимая информация имеется
            If (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
            (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1) Then
            
            'Вычисление Даты  Регистрации Клиента
            'Вычисление Даты  Регистрации Клиента
            frmTableCalendar.comCalendar.Today
            intDayReg = frmTableCalendar.comCalendar.Day
            intMonthReg = frmTableCalendar.comCalendar.Month
            intYearReg = frmTableCalendar.comCalendar.Year
            
            'Сделать доступным нажатие на кнопку "OK _ +"
                cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                cmdOK.SetFocus
            End If
            
            Exit Sub
            'Персональный код в недопустимом диапазоне
PersonCodeError:
            'Издать звуковой сигнал
            frmDemo.BeepSound
            txtPersonCode.Text = "Error"
            'Сбросить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 0
            'Белый фон текстового поля
            txtPersonCode.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопку "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            txtPersonCode.Text = "Error"
            'Сбросить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 0
            'Белый фон текстового поля
            txtPersonCode.BackColor = vbWhite
            'Установить фокус на текстовом поле "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            'Сделать недоступным нажатие на кнопку "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub

            'Процедура анализа "PersonCode" при АвтоРегистрации Посетителя
            '  через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            
             'Ждать завершения Активизации текущей формы
    Do While Me.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Занести ПЕРСОНАЛЬНЫЙ КОД в соответствующее
            '  текстовое поле
    txtPersonCode.Text = Trim(vntPersonCode)
            'Сделать недоступным текстовое поле ПЕРСОНАЛЬНОГО КОДА
    txtPersonCode.Enabled = False
            'Голубой фон текстового поля
    txtPersonCode.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Установить фокус на текстовом поле "Info"
    If txtInfo.Enabled = True Then txtInfo.SetFocus
            'Установлена Опция копирования "PersonCode"в поле "Info"
            '  при Регистрации Временного Посетителя
    If gAccessCodeInfo = 1 Then
            'Копирование "PersonCode"в поле "Info"
        txtInfo = Trim(txtPersonCode)
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
        txtInfo.Tag = 1
    End If
            'Выбрать опцию "Calendar"
    optCalendar.Value = True
            
            'Вычисление Даты  Регистрации Клиента
    frmTableCalendar.comCalendar.Today
    intDayReg = frmTableCalendar.comCalendar.Day
    intMonthReg = frmTableCalendar.comCalendar.Month
    intYearReg = frmTableCalendar.comCalendar.Year
    
End Function

            'Процедура формирования "PersonCode", "Info" и Печать
            '  талона со штрих-кодом (+ чека) при АвтоРегистрации Клиента
            '  через специальный "Controller" с кнопкой "DALLAS"
Public Function DallasButton(ByVal strAddrPortType As String, intIndex As Integer)
            'Статус
Dim strStatus As String
            'Строка "Контроль" для Предприятия
Dim strChecking As String * 8
            'Дата (и Время) регистрации Клиента
Dim strDate As String
            'Время регистрации Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Клиента
Dim strHour As String
Dim strMinute As String
            'Признак присутствия \ 0 - вошел \ 1 - вышел \ 2 - зарегистрирован
Dim strPersPresent As String * 1
            'Признак ("Е" - Окончательно вышел; "D" - Дневной тариф допуска;
            '  "N" - Ночной тариф допуска; "Другой символ"   - Суточный тариф
            '  допуска)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоРегистрации в "Таблице персон"
Dim intAutoRegistrCode  As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Рабочий счетчик - номер регистрации Клиента
Static btCount As Byte
            'Нормализованный (две цифры) номер регистрации Клиента
Dim strCount As String
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  Временному Посетителю бесплатно находиться на Предприятии
Dim intCellLimit As Integer
            'Строка отсылаемого сообщения
Dim strMessage As String

    
            'Номер регистрации Клиента
    If btCount < gMaxCount And btCount > gMinCount - 1 Then
        btCount = btCount + CByte(1)
    Else
        btCount = CByte(gMinCount)
    End If
    strCount = Trim(Str(btCount))
    
            'Очистка строки "Контроль" для Предприятия
    strChecking = ""
                'Время регистрации Клиента
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
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
            'Дата регистрации Клиента
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = gProtocol.strProtocDate
    If Len(Trim(strDate)) = 10 Then
        strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
        Right(Trim(strDate), 4)
    Else
        strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
        Right(Trim(strDate), 4)
    End If
            
            'Сформировать и занести ПЕРСОНАЛЬНЫЙ КОД
            '  в соответствующее текстовое поле
    txtPersonCode.Text = "0000" + Trim(strCount) + Trim(strHour) + _
    Trim(strMinute) + Left(Trim(strDate), 4) + Right(Trim(strDate), 2)
    
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Копирование "PersonCode"в поле "Info"
    txtInfo = Trim(txtPersonCode)
            'Установить признак  изменений в текстовом поле "Info"
    txtInfo.Tag = 1
            'Выбрать опцию "Time"
    optTime.Value = True
                'Признак времени допуска на Предприятие -
            '  Суточный (для Временных Клиентов)
    strTime = "DayNight"
            'Признак регистрации Клиента
    strPersPresent = "2"
            'Признак Клиента
    strExpander = "P"
            'Временный Клиент
    strStatus = "09 - Access/Time"
            
            'Входной тариф Предприятия равен нулю или
            '  ?предоплата хода - для унификации программы?
    If gAccessMoneyCell = 0 Or gParkInpCellNumb > 0 Then
            'Количество ячеек времени, в течение которого разрешается Временному
            '  Посетителю находиться на Предприятии (предоплата входа
            '  или бесплатный вход/выход)
        intCellLimit = gAcceInpCellNumb
    
            'Вычислить "сдвинутые" время и дату регистрации Посетителя,
            '  до которых ему будет разрешен Вход-Выход
            
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
                    strDate = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                Else
                    strDate = "0" + _
                    Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                    Right(strDate, 6)
                End If
            'Изменение  Месяца и, возможно, Года
                If frmTableCalendar.comCalendar.Day = 1 Then
                    If frmTableCalendar.comCalendar.Month > 9 Then
                        strDate = "01" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    Else
                        strDate = "010" + _
                        Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                        Right(strDate, 4)
                    End If
                End If
            End If
            
            'Не требуется переход часа
        Else
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
    End If
            
            'Формирование упакованной строки "Контроль" для Предприятия
    For intCount = 1 To 7 Step 2
            'Дата
        strChecking = Trim(strChecking) + _
        Chr(CByte(CInt(Mid(strDate, intCount, 2))))
    Next
            'Часы
    strChecking = Trim(strChecking) + _
    Chr(CByte(CInt(Mid(strHour, 1, 2))))
            'Минуты
    strChecking = Trim(strChecking) + _
    Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            'Упаковка строки "Контроль"
    Call frmTablePerson.Pack(strChecking)
            
            'Признак регистрации Клиента и Резерв для расширения
    strChecking = Left(strChecking, 6) + strPersPresent + strExpander
            
            'Вызов процедуры-функции АвтоРегистрации
            'ПЕРСОНАЛЬНОГО КОДА для Предприятия
    intAutoRegistrCode = frmTablePerson.AutoRegAccess(txtPersonCode.Text, _
    txtInfo.Text, strStatus, strChecking, strTime)
            '(Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА выполнена -
            '   протоколирование события
    If intAutoRegistrCode = 0 Then
            'Последняя (новая) строка "Таблицы персон"
        gTablePerson.Row = gTablePerson.Rows - 1
            'Текущий столбец "Таблицы персон" = 0 (Персона или Терминал)
        gTablePerson.Col = 0
        gProtocol.strProtocName = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 1 (Персональный код)
        gTablePerson.Col = 1
        gProtocol.strProtocPersonCode = gTablePerson.Text
            'Текущий столбец "Таблицы персон" = 2 (Статус)
        gTablePerson.Col = 2
        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "AutoRegAcce " + Left(Trim(txtMoneyDate.Text), 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
        Me.Tag = 1
            
            'Опция "Печать Документа" установлена
        If chkDocument.Value = 1 Then
            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
            Call frmDemo.PrintDocument(gProtocol.strProtocName, _
            gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
            gProtocol.strProtocTime, gProtocol.strProtocDate, _
            gProtocol.strProtocReserve, intError)
        End If
        
    
        
            'ОТКРЫТЬ ТЕРМИНАЛ ?
            
            'Установлен режим Выполнение - открыть терминал
        If intError = 0 And frmDemo.chkSetup.Value = 1 Then
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  выполнена Регистрация Клиента Предприятия
            If frmDemo.cmdOpen(intIndex).Tag = 0 And Me.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
                frmDemo.imgAccessInData(intIndex).Enabled = False
                frmDemo.imgAccessOutData(intIndex).Enabled = False
                frmDemo.imgAccessInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
                Call frmDemo.OpenBarrier(intIndex)
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                strMessage = "AcceFreePlaces-1"
            'Вызов процедуры-функции АвтоКоррекции для данного
            'ПЕРСОНАЛЬНОГО КОДА - Клиент вошел на Предприятие
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
            'Отослать СООБЩЕНИЕ
                Call frmDemo.SendMessage(strMessage)
            End If
        End If
        
            'Отказ в АвтоРегистрация ПЕРСОНАЛЬНОГО КОДА -
            '   протоколирование события
    Else
            'Введенная ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'Введенный ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Выбранная опция (Статус)
        gProtocol.strProtocStatus = strStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid AutoRegAccess"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
    
    End If
    
End Function
            
            'Процедура обработки "Щелчка мыши" на поле Информации
Private Sub txtInfo_Click()
            'Белый фон текстового поля
    txtInfo.BackColor = vbWhite
            'Сбросить признак  изменений в текстовом поле "Info"
    txtInfo.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub
            'Процедура ввода и анализа текстового поля "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Информация в допустимом диапазоне
        If (Len(Trim(txtInfo.Text)) < 17 And Len(Trim(txtInfo.Text)) > 0 _
        And (gAcceTimeLimit = 0 Or _
        (gAcceTimeLimit > 0 And optCalendar.Value = False))) Or _
        (Len(Trim(txtInfo.Text)) < 11 And Len(Trim(txtInfo.Text)) > 0 _
        And gAcceTimeLimit > 0 And optCalendar.Value = True) Then
            'Установить признак  изменений в текстовом поле "Info"
            txtInfo.Tag = 1
            'Установить фокус на текстовом поле "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            'Вся необходимая информация имеется
            If (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
            (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1) Then
            'Сделать доступным нажатие на кнопку "OK _ +"
                 cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК _+"
                cmdOK.SetFocus
            End If
            Exit Sub
            'Имя в недопустимом диапазоне
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            txtInfo.Text = "Error"
            'Сбросить признак  изменений в текстовом поле "Info"
            txtInfo.Tag = 0
            'Белый фон текстового поля
            txtInfo.BackColor = vbWhite
            'Установить фокус на текстовом поле "Info"
            If txtInfo.Enabled = True Then txtInfo.SetFocus
            'Сделать недоступным нажатие на кнопку "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub
            
            'Выбрана опция - "Calendar"
Private Sub optCalendar_Click()
            'Сделать доступным элемент управления "fraDayNight"
    fraDayNight.Enabled = True
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать видимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = True
            'Выбрать опцию "Human"
    optHuman.Value = True
            'Сделать недоступными некоторые элементы управления формы
    fraPeople.Enabled = False
            'Сделать доступными некоторые элементы управления формы
    lblLat0.Enabled = True
    lblLat320.Enabled = True
    hsbLat.Enabled = True
    txtPersonCode.Enabled = True
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Day"
Private Sub optDay_Click()
            'Тариф одного дня (День)
    intAccessTariff = intAccessTariffDay
            'Признак времени допуска - Дневной (для Постоянных Посетителей)
    strTime = "Day"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "DayNight"
Private Sub optDayNight_Click()
            'Тариф одного дня (Сутки)
    intAccessTariff = intAccessTariffFull
            'Признак времени допуска - Суточный (для Постоянных Посетителей)
    strTime = "DayNight"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Night"
Private Sub optNight_Click()
            'Тариф одного дня (Ночь)
    intAccessTariff = intAccessTariffNight
            'Признак времени допуска - Ночной (для Постоянных Посетителей)
    strTime = "Night"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "MoneyFree"
Private Sub optMoneyFree_Click()
            'Сделать недоступным элемент управления "fraDayNight"
    fraDayNight.Enabled = False
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Выбрать опцию "Human"
    optHuman.Value = True
            'Сделать недоступными некоторые элементы управления формы
    fraPeople.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtPersonCode.Enabled = True
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать доступным нажатие на кнопку "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub
            
            'Выбрана опция - "Time"
Private Sub optTime_Click()
            'Сделать недоступным элемент управления "fraDayNight"
    fraDayNight.Enabled = False
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать доступными некоторые элементы управления формы
    fraPeople.Enabled = True
            'Выбрать опцию "Human"
    optHuman.Value = True
            'Сделать недоступными некоторые элементы управления формы
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
        cmdOK.MousePointer = 0
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If
            
End Sub

            'Процедура контроля времени при вводе пароля - событие "TimeOut"
Private Sub tmrParoleTimeOut_Timer()
            'Издать звуковой сигнал
    frmDemo.BeepSound
    
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
    frmDemo.WriteProtocol

            ' "Очистка" поля пароля пробелами
    txtParole.Text = ""
            ' "Погасить" этикетку "Пароль"
    lblParole.Enabled = False
            'Сбросить контроль времени ввода пароля
    tmrParoleTimeOut.Enabled = False
            'Белый фон текстового поля
    txtParole.BackColor = vbWhite
            'Сделать доступными кнопки "OK" и "Cancel"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    
End Sub

            'Процедура обработки "Щелчка мыши" на поле пароля
Private Sub txtParole_Click()
            'Белый фон текстового поля
    txtParole.BackColor = vbWhite
            'Сделать недоступными кнопки "OK" и "Cancel"
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
            ' "Проявить" этикетку "Пароль"
    lblParole.Enabled = True
            ' "Очистка" поля пароля пробелами
    strPassword = ""
            'Установить контроль времени ввода пароля
    tmrParoleTimeOut.Enabled = True
           'Удержание фокуса клавиатуры на поле пароля до его ввода
           '  или истечения контрольного времени
    Do While strPassword = "" And tmrParoleTimeOut.Enabled = True
        DoEvents
    Loop

End Sub

            'Процедура ввода и анализа пароля
Private Sub txtParole_KeyPress(KeyAscii As Integer)
            'Пароль ввведен и "Проявлена" этикетка "Пароль"
    If KeyAscii = vbKeyReturn And lblParole.Enabled = True Then
             'Голубой фон текстового поля
        txtParole.BackColor = vbCyan
           'Пароль
        strPassword = txtParole.Text
        
            'Протоколирование события - "Ввод пароля"
        gProtocol.strProtocName = "????????????????"
            'Системный пароль
        gProtocol.strProtocPersonCode = txtParole.Text
            'Статус
        gProtocol.strProtocStatus = "04 - Operator"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "PASSWORD Input"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
        
            'Анализ правильности текущего пароля - правильный
        If txtParole.Text = txtParole.Tag Then
            'Сделать доступными опции "MoneyFree" и "Document"
            imgDocument.Enabled = True
            chkDocument.Enabled = True
            imgMoneyFree.Enabled = True
            optMoneyFree.Enabled = True
            'Пароль неверный
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Сделать недоступными опции "MoneyFree" и "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
            imgMoneyFree.Enabled = False
            optMoneyFree.Enabled = False
             'Белый фон текстового поля
            txtParole.BackColor = vbWhite
            'Установить фокус на текстовом поле "Parole"
            If txtParole.Enabled = True Then txtParole.SetFocus
        End If
            'Сбросить контроль времени ввода пароля
        tmrParoleTimeOut.Enabled = False
            ' "Очистка" поля пароля пробелами
        txtParole.Text = ""
            ' "Погасить" этикетку "Пароль"
        lblParole.Enabled = False
            'Сделать доступными кнопки "OK" и "Cancel"
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If

End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Change()
            
            'Ползунок полосы прокрутки Латов "Уперся" справа
    If hsbLat.Value > hsbLat.Tag And (hsbLat.Tag * 100 + intAccessTariff) > 32000 Then
            'Восстановление предыдущего положения ползунков
        hsbSant.Value = hsbSant.Tag
        hsbLat.Value = hsbLat.Tag
    ElseIf hsbLat.Value = hsbLat.Tag Then
        Exit Sub
    End If
            'Вносимая сумма оплаты в Сантимах
    intAccessMoney = hsbLat.Value * 100 + hsbSant.Value
            'Ползунки полос прокрутки Латов и Сантимов в некорректном положении
            '  (вносимая сумма не оплачивает Целое число дней)
    If Int(intAccessMoney / intAccessTariff) * 100 <> intAccessMoney Or _
    hsbLat.Value * 100 > intAccessTariff Then
            'Ползунок двигался в сторону увеличения суммы
        If hsbLat.Value > hsbLat.Tag Then
            intAccessMoney = hsbLat.Tag * 100 + hsbSant.Tag + intAccessTariff
            'Ползунок двигался в сторону уменьшения суммы
        ElseIf hsbLat.Value < hsbLat.Tag Then
            intAccessMoney = hsbLat.Tag * 100 + hsbSant.Tag - intAccessTariff
        End If
            'Восстановление корректного положения ползунков
            hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
            hsbLat.Value = Int(intAccessMoney / 100)
            'Запомнить новое положение ползунков
        hsbSant.Tag = hsbSant.Value
        hsbLat.Tag = hsbLat.Value
    End If
            'Запомнить новое положение ползунков
    hsbSant.Tag = hsbSant.Value
    hsbLat.Tag = hsbLat.Value
            
            'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Изменение текстового поля "ДеньгиДаты"
    If hsbLat.Value < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    ElseIf hsbLat.Value < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    ElseIf hsbLat.Value > 99 Then
        txtMoneyDate.Text = Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If hsbSant.Value < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(hsbSant.Value)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(hsbSant.Value)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            'Не нулевое положение одного из ползунков полос прокрутки
    If hsbLat.Value > 0 Or hsbSant.Value > 0 Then
            'Установить признак  внесенной информации
        txtMoneyDate.Tag = 1
            'Количество дней
        intAccessDay = Int(intAccessMoney / intAccessTariff)
            
            'Установка "Календаря" на текущую дату
        frmTableCalendar.comCalendar.Today
            'Цикл по Дням "Календаря" (от Даты Регистрации Клиента)
        For intAccessDay = intAccessDay To 1 Step -1
            
            'Запись Числа, Месяца и Года в поле "ДеньгиДаты"
            If frmTableCalendar.comCalendar.Month > 9 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year))
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + ".0" + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year))
            End If
            'Продвижение "Календаря" на один день вперед
            frmTableCalendar.comCalendar.NextDay
            
        Next
            
            'Отмена внесенной информации
    Else
        txtMoneyDate.Tag = 0
            'Дата
        txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Сделать недоступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = vbNoDrop
    End If
EndCycle:
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Недостаточная оплата одного дня
    If Int(intAccessMoney / intAccessTariff) = 0 Then
           'Отмена внесенной информации
        txtMoneyDate.Tag = 0
           'Дата
        txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
           'Формирование шаблона в поле "ДеньгиДата"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = vbNoDrop
    End If
            'Переплата (возможна оплата только до конца года)
    If intAccessDay > 0 Then
            'Количество корректных (без переплаты) дней
        intAccessDay = Int(intAccessMoney / intAccessTariff) - intAccessDay
           'Восстановление корректной (без переплаты) суммы оплаты в Сантимах
        intAccessMoney = intAccessDay * intAccessTariff
            'Восстановление корректного положения ползунков
        hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
        hsbLat.Value = Int(intAccessMoney / 100)
        hsbLat_Change
    End If
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If
    
End Sub
