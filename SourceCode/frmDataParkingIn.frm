VERSION 5.00
Begin VB.Form frmDataParkingIn 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ParkingInData"
   ClientHeight    =   3960
   ClientLeft      =   4860
   ClientTop       =   2745
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   186
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
   ScaleHeight     =   3960
   ScaleWidth      =   6990
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.VScrollBar vsbDate 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4680
      Max             =   366
      TabIndex        =   31
      Top             =   1680
      Width           =   255
   End
   Begin VB.Frame fraMonth 
      Caption         =   " D   1/2M  1M   2M"
      Height          =   615
      Left            =   5040
      TabIndex        =   26
      Top             =   1680
      Width           =   1695
      Begin VB.OptionButton optNot 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optHalf 
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optOne 
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optTwo 
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
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
      TabIndex        =   17
      Top             =   3240
      Width           =   1212
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   14
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
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
      TabIndex        =   10
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
         Begin VB.OptionButton optNight 
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDay 
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
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
         TabIndex        =   13
         Top             =   600
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
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
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
         TabIndex        =   11
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingIn.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTime 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingIn.frx":045A
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image imgCalendar 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingIn.frx":20FC
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
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
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   320
      TabIndex        =   4
      Top             =   2400
      Width           =   1452
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   99
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   2
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   " ""-1"" D/M ""+1"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   32
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Ls"
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
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image imgParkingIn 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataParkingIn.frx":28FE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1560
      Y2              =   3120
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
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
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   720
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6840
      X2              =   4080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2280
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   1800
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataParkingIn.frx":2B10
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
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
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLat0 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      TabIndex        =   9
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label lblLat320 
      Alignment       =   2  'Center
      Caption         =   "320"
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
      TabIndex        =   8
      Top             =   2400
      Width           =   375
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
      TabIndex        =   7
      Top             =   1920
      Width           =   495
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
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDataParkingIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Вносимая сумма оплаты в Сантимах
Dim intParkingMoney As Integer
            'Количество парковочных дней
Dim intParkingDay As Integer
            'Тариф одного парковочного дня (Сутки)
Dim intParkingTariffFull As Integer
            'Тариф одного парковочного дня (День)
Dim intParkingTariffDay As Integer
            'Тариф одного парковочного дня (Ночь)
Dim intParkingTariffNight As Integer
            'Тариф Автостоянки (переменная для рассчетов)
Dim intParkingTariff As Integer
            'Признак времени допуска к Автостоянке -
            '   (для Постоянных Клиентов)
Dim strTime As String
            'Текущая строка "Таблицы календаря"
Dim intRowNum As Integer
            'Текущая столбец "Таблицы календаря"
Dim intColNum As Integer
            'Номер позиции заданного символа в строке
Dim intPosNum As Integer
             'Введенный пароль
Dim strPassword As String

            'Перехват нажатия комбинаций клавиш "Alt"+ {"+" и "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Форма "frmDataParkingIn" доступна
    If frmDataParkingIn.Enabled = True Then
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
            'Строка "Контроль" для Автостоянок
Dim strChecking As String * 8
            'Подстрока "Контроль" поля "txtInfo"
Dim strCheckingInfo As String * 8
            'Дата (и Время) регистрации Посетителя или
            '  дата последнего оплаченного дня
Dim strDate As String
            'Подстрока "Контрольные Дата и Время" поля "txtInfo"
Dim strDateInfo As String
            'Время регистрации Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Клиента
Dim strHour As String
Dim strMinute As String
            'Признак присутствия АМ \ 0 - въехал \ 1 - выехал \ 2 - зарегистрирован
Dim strCarPresent As String * 1
            'Признак АМ ("Е" - Окончательно выехал; "D" - Дневной тариф допуска
            '  на Автостоянку; "N" - Ночной тариф допуска на Автостоянку; "Другой
            '  символ"   - Суточный тариф допуска на Автостоянку)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоРегистрации в "Таблице персон"
Dim intAutoRegistrCode  As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
Dim intCellLimit As Integer
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer

            'Недоступное нажатие на кнопку "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            'Очистка строки и подстроки "Контроль"
    strChecking = ""
    strCheckingInfo = ""
            'Вычислить время регистрации Клиента
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
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
    strDateInfo = strDate
            
            'Признак регистрации АМ Клиента
    strCarPresent = "2"
            'Признак АМ Клиента
    strExpander = "P"
            'Анализ статуса Клиента Автостоянки
    If optMoneyFree.Value = True Then
            'Бесплатный Клиент
        strStatus = "07 - Parking/Free"
    ElseIf optCalendar.Value = True Then
            'Постоянный Клиент
        strStatus = "05 - Parking/Calen."
            'Клиент с Дневным тарифом допуска на Автостоянку
        If optCalendar.Value = True And optDay.Value = True Then
            strExpander = "D"
            'Клиент с Ночным тарифом допуска на Автостоянку
        ElseIf optCalendar.Value = True And optNight.Value = True Then
            strExpander = "N"
        End If
            'Дата последнего оплаченного парковочного дня
        strDate = Mid(Trim(txtMoneyDate.Text), 11)
        If Len(Trim(strDate)) = 10 Then
            strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
            Right(Trim(strDate), 4)
        Else
            strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
            Right(Trim(strDate), 4)
        End If
    ElseIf optTime.Value = True Then
            'Временный Клиент
        strStatus = "06 - Parking/Time"
    
            'Въездной тариф Автостоянки равен нулю или
            '  ?предоплата въезда - для унификации программы?
        If gParkingMoneyCell = 0 Or gParkInpCellNumb > 1 Then
            'Количество ячеек времени, в течение которого разрешается Временному
            '  Клиенту находиться на Автостоянке (предоплата въезда
            '  или бесплатный въезд/выезд)
            intCellLimit = gParkInpCellNumb
    
            'Вычислить "сдвинутые" время и дату регистрации Клиента,
            '  до которых ему будет разрешен Въезд-Выезд
            
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
            
            'Последний столбец "Таблицы календаря" (с текущим днем)
                    If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            'Последняя строка "Таблицы календаря" (с текущим днем)
                        If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            'Издать звуковой сигнал
                            frmDemo.BeepSound
                            If frmDemo.optEnglish = True Then
                                MsgBox ("New Year TableCalendar Error")
                            Else
                                MsgBox ("Jauna gada kalend. nesask.")
                            End If
            'Оставить последние возможные строку и столбец "Таблицы календаря"
            
            'Не последняя строка "Таблицы календаря" (с текущим днем)
                        Else
            'Текущая строка "Таблицы календаря"+1 (следующий день)
                            frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            'Текущий столбец "Таблицы календаря" =1 (следующий день)
                            frmTableCalendar.grdTableCalendar.Col = 1
                        End If
            'Не последний столбец "Таблицы календаря" (с текущим днем)
                    Else
            'Текущая строка "Таблицы календаря" (текущий день)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum
            'Текущий столбец "Таблицы календаря" +1 (следующий день)
                        frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                    End If
                
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
                    intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                    If intPosNum <> 0 Then
            'Изменение  Числа
                        If intPosNum = 3 Then
                            strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                            Right(strDate, 6)
                        Else
                            strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                            Right(strDate, 6)
                        End If
                    Else
            'Изменение  Числа
                        If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                            strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                            Right(strDate, 6)
                        Else
                            strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                            Right(strDate, 6)
                        End If
                    End If
            'Переход месяца
                    If Left(strDate, 2) = "01" Then
            'Изменение  Месяца
                        If CInt(Mid(strDate, 3, 2)) + 1 > 9 Then
                            strDate = "01" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                            Right(strDate, 4)
                        Else
                            strDateInfo = "01" + "0" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                            Right(strDate, 4)
                        End If
            'Переход года
                        If Mid(strDate, 3, 2) = "13" Then
            'Изменение  Месяца и Года
                            strDate = "01" + "01" + Trim(Str(CInt(Right(strDateInfo, 4)) + 1))
                        End If
                    End If
                End If
            'Не требуется переход часа
            Else
                intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            
            'Признак регистрации АМ Клиента и Резерв для расширения
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            'Постоянный Клиент на Автостоянке с ограничением времени
            '  непрерывного пребывания АМ на Автостоянке
    If gParkTimeLimit > 0 And optCalendar.Value = True Then
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
        intCellLimit = gParkingCellLimit
            
            'Вычислить "сдвинутые" время и дату регистрации Постоянного
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
            
            'Последний столбец "Таблицы календаря" (с текущим днем)
                If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            'Последняя строка "Таблицы календаря" (с текущим днем)
                    If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            'Издать звуковой сигнал
                        frmDemo.BeepSound
                        If frmDemo.optEnglish = True Then
                            MsgBox ("New Year TableCalendar Error")
                        Else
                            MsgBox ("Jauna gada kalend. nesask.")
                        End If
            'Оставить последние возможные строку и столбец "Таблицы календаря"
            
            'Не последняя строка "Таблицы календаря" (с текущим днем)
                    Else
            'Текущая строка "Таблицы календаря"+1 (следующий день)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            'Текущий столбец "Таблицы календаря" =1 (следующий день)
                        frmTableCalendar.grdTableCalendar.Col = 1
                    End If
            'Не последний столбец "Таблицы календаря" (с текущим днем)
                Else
            'Текущая строка "Таблицы календаря" (текущий день)
                    frmTableCalendar.grdTableCalendar.Row = gRowNum
            'Текущий столбец "Таблицы календаря" +1 (следующий день)
                    frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                End If
                
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            'Изменение  Числа
                    If intPosNum = 3 Then
                        strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDateInfo, 6)
                    Else
                        strDateInfo = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDateInfo, 6)
                    End If
                Else
            'Изменение  Числа
                    If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                        strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDateInfo, 6)
                    Else
                        strDateInfo = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDateInfo, 6)
                    End If
                End If
            'Переход месяца
                If Left(strDateInfo, 2) = "01" Then
            'Изменение  Месяца
                    If CInt(Mid(strDateInfo, 3, 2)) + 1 > 9 Then
                        strDateInfo = "01" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                        Right(strDateInfo, 4)
                    Else
                        strDateInfo = "01" + "0" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                        Right(strDateInfo, 4)
                    End If
            'Переход года
                    If Mid(strDateInfo, 3, 2) = "13" Then
            'Изменение  Месяца и Года
                        strDateInfo = "01" + "01" + Trim(Str(CInt(Right(strDateInfo, 4)) + 1))
                    End If
                End If
            End If
            'Не требуется переход часа
        Else
            intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            'ПЕРСОНАЛЬНОГО КОДА для Автостоянки
    intAutoRegistrCode = frmTablePerson.AutoRegParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoRegPark " + Left(Trim(txtMoneyDate.Text), 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            'Признак (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingIn.Tag = 1
            
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
            '   выполнение процедур невозможно), Временный Клиент Автостоянки
            '   и установлен индекс въездного терминала - открыть терминал
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        optTime.Value = True And gTermInp <> -1 Then
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  выполнена Регистрация Клиента Автостоянки
            If frmDemo.cmdOpen(gTermInp).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
                frmDemo.imgParkingInData(gTermInp).Enabled = False
                frmDemo.imgParkingOutData(gTermInp).Enabled = False
                frmDemo.imgParkingInfoData(gTermInp).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gParkAddrTerm(gTermInp))))
                frmDemo.cmdOpen(gTermInp).Tag = vntAddr
                frmDemo.cmdOpen(gTermInp).Caption = "Addr=" + CStr(vntAddr)
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(gTermInp).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(gTermInp).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
                Call frmDemo.OpenBarrier(gTermInp)
            'Вызов процедуры-функции АвтоКоррекции для данного
            'ПЕРСОНАЛЬНОГО КОДА - Автомобиль въехал на Автостоянку
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "Invalid AutoRegParking"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingIn.Tag = 2
    
    
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
            
            'Были не сохраненные изменения в текстовых полях текущей формы
    If frmDataParkingIn.Tag = 1 And _
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
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
    If frmDataParkingIn.Tag = 0 Then frmDataParkingIn.Tag = 2
            'Сделать невидимой текущую форму
    frmDataParkingIn.Visible = False
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
            'Выбрать опцию "Not"
    optNot.Value = True
            'Выбрать опцию "Time"
    optTime.Value = True
            'Сделать недоступными элементы управления формы "DataParkingIn"
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    fraDayNight.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    txtMoneyDate.Enabled = False
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
           'Дата
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Тариф одного парковочного дня (Сутки)
    intParkingTariff = intParkingTariffFull
            'Признак времени допуска к Автостоянке -
            '  Суточный (для Постоянных Клиентов)
    strTime = "DayNight"
           'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop
            'Установить фокус на текстовом поле "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
             'Установить флаг завершения Активизации текущей формы
    frmDataParkingIn.Tag = 1

End Sub

            'Деактивизация текущей формы
Private Sub Form_Deactivate()
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub
            
            'Загрузка текущей формы
Private Sub Form_Load()
            'Сделать недоступными элементы управления формы "DataParkingIn"
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    optNot.Value = True
    optTime.Value = True
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Тариф одного парковочного дня (Сутки)
    intParkingTariffFull = gParkingDN
            'Тариф одного парковочного дня (День)
    intParkingTariffDay = gParkingD
            'Тариф одного парковочного дня (Ночь)
    intParkingTariffNight = gParkingN
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
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
            If gParkingCodeInfo = 1 Then
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
            'Установить фокус на текстовом поле "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
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

            'Процедура анализа "PersonCode" при АвтоРегистрации Клиента
            '  Автостоянки через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
             'Ждать завершения Активизации текущей формы
    Do While frmDataParkingIn.Tag = 0
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
    If gParkingCodeInfo = 1 Then
            'Копирование "PersonCode"в поле "Info"
        txtInfo = Trim(txtPersonCode)
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
        txtInfo.Tag = 1
    End If
            'Выбрать опцию "Calendar"
    optCalendar.Value = True
    
End Function

            'Процедура формирования "PersonCode", "Info" и Печать
            '  талона со штрих-кодом (+ чека) при АвтоРегистрации Клиента
            '  через специальный "Controller" с кнопкой "DALLAS"
Public Function DallasButton(ByVal strAddrPortType As String, intIndex As Integer)
            'Статус
Dim strStatus As String
            'Строка "Контроль" для Автостоянок
Dim strChecking As String * 8
            'Дата (и Время) регистрации Клиента
Dim strDate As String
            'Время регистрации Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Клиента
Dim strHour As String
Dim strMinute As String
            'Признак присутствия АМ \ 0 - въехал \ 1 - выехал \ 2 - зарегистрирован
Dim strCarPresent As String * 1
            'Признак АМ ("Е" - Окончательно выехал; "D" - Дневной тариф допуска
            '  на Автостоянку; "N" - Ночной тариф допуска на Автостоянку; "Другой
            '  символ"   - Суточный тариф допуска на Автостоянку)
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
            '  АМ Временного Клиента бесплатно находиться на Автостоянке
Dim intCellLimit As Integer
            'Строка отсылаемого сообщения
Dim strMessage As String
    
            'Номер регистрации Клиента
    If btCount < 99 And btCount > 9 Then
        btCount = btCount + CByte(1)
    Else
        btCount = CByte(10)
    End If
    strCount = Trim(Str(btCount))
    
            'Очистка строки "Контроль" для Автостоянок
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
    txtPersonCode.Text = "000000" + Trim(strCount) + Trim(strHour) + _
    Trim(strMinute) + Left(Trim(strDate), 4)
    
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Копирование "PersonCode"в поле "Info"
    txtInfo = Trim(txtPersonCode)
            'Установить признак  изменений в текстовом поле "Info"
    txtInfo.Tag = 1
            'Выбрать опцию "Time"
    optTime.Value = True
                'Признак времени допуска к Автостоянке -
            '  Суточный (для Временных Клиентов)
    strTime = "DayNight"
            'Признак регистрации АМ Клиента
    strCarPresent = "2"
            'Признак АМ Клиента
    strExpander = "P"
            'Временный Клиент
    strStatus = "06 - Parking/Time"
    
            'Въездной тариф Автостоянки равен нулю или
            '  ?предоплата въезда - для унификации программы?
    If gParkingMoneyCell = 0 Or gParkInpCellNumb > 0 Then
            'Количество ячеек времени, в течение которого разрешается Временному
            '  Клиенту находиться на Автостоянке (предоплата въезда
            '  или бесплатный въезд/выезд)
        intCellLimit = gParkInpCellNumb
    
            'Вычислить "сдвинутые" время и дату регистрации Клиента,
            '  до которых ему будет разрешен Въезд-Выезд
            
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
            
            'Последний столбец "Таблицы календаря" (с текущим днем)
                If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            'Последняя строка "Таблицы календаря" (с текущим днем)
                    If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            'Издать звуковой сигнал
                        frmDemo.BeepSound
                        If frmDemo.optEnglish = True Then
                            MsgBox ("New Year TableCalendar Error")
                        Else
                            MsgBox ("Jauna gada kalend. nesask.")
                        End If
            'Оставить последние возможные строку и столбец "Таблицы календаря"
                
            'Не последняя строка "Таблицы календаря" (с текущим днем)
                    Else
            'Текущая строка "Таблицы календаря"+1 (следующий день)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            'Текущий столбец "Таблицы календаря" =1 (следующий день)
                        frmTableCalendar.grdTableCalendar.Col = 1
                    End If
            'Не последний столбец "Таблицы календаря" (с текущим днем)
                Else
            'Текущая строка "Таблицы календаря" (текущий день)
                    frmTableCalendar.grdTableCalendar.Row = gRowNum
            'Текущий столбец "Таблицы календаря" +1 (следующий день)
                    frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                End If
                
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            'Изменение  Числа
                    If intPosNum = 3 Then
                        strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDate, 6)
                    Else
                        strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDate, 6)
                    End If
                Else
            'Изменение  Числа
                    If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                        strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDate, 6)
                    Else
                        strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDate, 6)
                    End If
                End If
            'Переход месяца
                If Left(strDate, 2) = "01" Then
            'Изменение  Месяца
                    If CInt(Mid(strDate, 3, 2)) + 1 > 9 Then
                        strDate = "01" + Trim(Str(CInt(Mid(strDate, 3, 2)) + 1)) + _
                        Right(strDate, 4)
                    Else
                        strDate = "01" + "0" + Trim(Str(CInt(Mid(strDate, 3, 2)) + 1)) + _
                        Right(strDate, 4)
                    End If
            'Переход года
                    If Mid(strDate, 3, 2) = "13" Then
            'Изменение  Месяца и Года
                        strDate = "01" + "01" + Trim(Str(CInt(Right(strDate, 4)) + 1))
                    End If
                End If
            End If
            'Не требуется переход часа
        Else
            intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            
            'Формирование упакованной строки "Контроль" для Автостоянок
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
            
            'Признак регистрации АМ Клиента и Резерв для расширения
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            'Вызов процедуры-функции АвтоРегистрации
            'ПЕРСОНАЛЬНОГО КОДА для Автостоянки
    intAutoRegistrCode = frmTablePerson.AutoRegParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoRegPark " + Left(Trim(txtMoneyDate.Text), 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingIn.Tag = 1
            
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
            '  выполнена Регистрация Клиента Автостоянки
            If frmDemo.cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
                frmDemo.imgParkingInData(intIndex).Enabled = False
                frmDemo.imgParkingOutData(intIndex).Enabled = False
                frmDemo.imgParkingInfoData(intIndex).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(intIndex).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(intIndex).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
                Call frmDemo.OpenBarrier(intIndex)
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                strMessage = "ParkFreePlaces-1"
            'Отослать СООБЩЕНИЕ
                Call frmDemo.SendMessage(strMessage)
            'Вызов процедуры-функции АвтоКоррекции для данного
            'ПЕРСОНАЛЬНОГО КОДА - Автомобиль въехал на Автостоянку
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorParking(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
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
        gProtocol.strProtocReserve = "Invalid AutoRegParking"
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
        And (gParkTimeLimit = 0 Or _
        (gParkTimeLimit > 0 And optCalendar.Value = False))) Or _
        (Len(Trim(txtInfo.Text)) < 11 And Len(Trim(txtInfo.Text)) > 0 _
        And gParkTimeLimit > 0 And optCalendar.Value = True) Then
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
            'Информация в недопустимом диапазоне
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
    lblDate.Visible = True
            'Сделать доступными некоторые элементы управления формы
    hsbLat.Enabled = True
    vsbDate.Enabled = True
    fraMonth.Enabled = True
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Day"
Private Sub optDay_Click()
            'Тариф одного парковочного дня (День)
    intParkingTariff = intParkingTariffDay
            'Признак времени допуска к Автостоянке -
            '  Дневной (для Постоянных Клиентов)
    strTime = "Day"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "DayNight"
Private Sub optDayNight_Click()
            'Тариф одного парковочного дня (Сутки)
    intParkingTariff = intParkingTariffFull
            'Признак времени допуска к Автостоянке -
            '  Суточный (для Постоянных Клиентов)
    strTime = "DayNight"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Night"
Private Sub optNight_Click()
            'Тариф одного парковочного дня (Ночь)
    intParkingTariff = intParkingTariffNight
            'Признак времени допуска к Автостоянке -
            '  Ночной (для Постоянных Клиентов)
    strTime = "Night"
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Not"
Private Sub optNot_Click()
            'Сделать видимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = True
    lblDate.Visible = True
            'Сделать доступными некоторые элементы управления формы
    hsbLat.Enabled = True
    vsbDate.Enabled = True
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub
            
            'Выбрана опция - "Half"
Private Sub optHalf_Click()
            'Количество дней до конца месяца
    Dim intToMonthEnd As Integer
            'Конечное число действия абонимента
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            'Стоимость абонимента на 1/2 месяцa
    Dim intLat As Integer
    Dim intSant As Integer
            
            'Тариф одного парковочного дня (Сутки)
    intParkingTariff = intParkingTariffFull
            'Признак времени допуска к Автостоянке -
            '  Суточный (для Постоянных Клиентов)
    strTime = "DayNight"
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

            'Количество дней до конца месяца
    intToMonthEnd = -1
            'Цикл по строкам "Таблицы календаря" (с текущего дня)
    For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = gRowNum Then
            'По столбцам "Таблицы календаря" (с текущего дня)
            intColNum = gColNum
        Else
            'По всем столбцам "Таблицы календаря"
            intColNum = 1
        End If
            'По всем столбцам "Таблицы календаря"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Количество дней до конца месяца
            intToMonthEnd = intToMonthEnd + 1
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              'Переход месяца
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            Else
              'Переход месяца
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            'Вычисление конечного числа действия абонимента
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Переход года
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - на Февраль
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And 0 = gYear Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And Not (0 = gYear Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            'Переход месяца - до Июля
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
            'Переход месяца - на Август
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 7 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - после Августа
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
    End If
    
              'Переход года и месяца
    If Mid(txtMoneyDate.Text, 4, 2) = 12 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              'Последний месяц этого же года
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 16)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              'Этот же месяц - Февраль
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 2 And intFinishDay <= 15 Then
        If 0 = gYear Mod 4 Then
            txtMoneyDate.Text = Trim(Str(intFinishDay + 14)) + _
            Trim(Mid(txtMoneyDate.Text, 3))
        Else
            txtMoneyDate.Text = Trim(Str(intFinishDay + 13)) + _
            Trim(Mid(txtMoneyDate.Text, 3))
        End If
              'Этот же месяц - не Февраль
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 15)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              'Этот же месяц
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 15)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
    End If
            
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Стоимость суточного абонимента на 1/2 Месяцa
    intLat = intParkingTariff * 15
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            'Изменение текстового поля "ДеньгиДаты"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub
            
            'Выбрана опция - "One"
Private Sub optOne_Click()
            'Количество дней до конца месяца
    Dim intToMonthEnd As Integer
            'Конечное число действия абонимента
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            'Стоимость суточного абонимента на 1 месяц
    Dim intLat As Integer
    Dim intSant As Integer
            
            'Тариф одного парковочного дня (Сутки)
    intParkingTariff = intParkingTariffFull
            'Признак времени допуска к Автостоянке -
            '  Суточный (для Постоянных Клиентов)
    strTime = "DayNight"
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop

            'Количество дней до конца месяца
    intToMonthEnd = -1
            'Цикл по строкам "Таблицы календаря" (с текущего дня)
    For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = gRowNum Then
            'По столбцам "Таблицы календаря" (с текущего дня)
            intColNum = gColNum
        Else
            'По всем столбцам "Таблицы календаря"
            intColNum = 1
        End If
            'По всем столбцам "Таблицы календаря"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Количество дней до конца месяца
            intToMonthEnd = intToMonthEnd + 1
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              'Переход месяца
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            Else
              'Переход месяца
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            'Вычисление конечного числа действия абонимента
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Переход года
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - на Февраль
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And 0 = gYear Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And Not (0 = gYear Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            'Переход месяца - до Июля
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
            'Переход месяца - на Август
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 7 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - после Августа
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
    End If
    strFinishDay = Trim(Str(intFinishDay))
    
              'Переход года и месяца
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        txtMoneyDate.Text = strFinishDay + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 Then
        txtMoneyDate.Text = strFinishDay + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 Then
        txtMoneyDate.Text = strFinishDay + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
    End If
            
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Стоимость суточного абонимента на 1 Месяц
    intLat = intParkingTariff * 30
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            'Изменение текстового поля "ДеньгиДаты"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub
            
            'Выбрана опция - "Two"
Private Sub optTwo_Click()
            'Количество дней до конца месяца
    Dim intToMonthEnd As Integer
            'Конечное число действия абонимента
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            'Стоимость суточного абонимента на 2 месяцa
    Dim intLat As Integer
    Dim intSant As Integer
            
            'Тариф одного парковочного дня (Сутки)
    intParkingTariff = intParkingTariffFull
            'Признак времени допуска к Автостоянке -
            '  Суточный (для Постоянных Клиентов)
    strTime = "DayNight"
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK _ +"
     cmdOK.MousePointer = vbNoDrop

            'Количество дней до конца месяца
    intToMonthEnd = -1
            'Цикл по строкам "Таблицы календаря" (с текущего дня)
    For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = gRowNum Then
            'По столбцам "Таблицы календаря" (с текущего дня)
            intColNum = gColNum
        Else
            'По всем столбцам "Таблицы календаря"
            intColNum = 1
        End If
            'По всем столбцам "Таблицы календаря"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Количество дней до конца месяца
            intToMonthEnd = intToMonthEnd + 1
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              'Переход месяца
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            Else
              'Переход месяца
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Количество дней до конца месяца (без текущего дня)
                    intToMonthEnd = intToMonthEnd - 1
            'Количество дней до конца месяца исчерпано
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            'Вычисление конечного числа действия абонимента
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            'Переход года
    If Mid(txtMoneyDate.Text, 4, 2) = 11 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход  года & месяца - на Февраль
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And 0 = (gYear + 1) Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And Not (0 = (gYear + 1) Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            'Переход месяца - до Июля
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 6 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 6 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - на Август
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 6 Then
        intFinishDay = 31 - intToMonthEnd
            'Переход месяца - после Августа
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 6 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 6 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
    End If
    strFinishDay = Trim(Str(intFinishDay))
    
              'Переход года и месяца
    If Mid(txtMoneyDate.Text, 4, 2) = 11 Then
        txtMoneyDate.Text = strFinishDay + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        txtMoneyDate.Text = strFinishDay + ".02." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 8 Then
        txtMoneyDate.Text = strFinishDay + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 2)) + Trim(Mid(txtMoneyDate.Text, 6))
              'Переход месяца
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 8 Then
        txtMoneyDate.Text = strFinishDay + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 2)) + Trim(Mid(txtMoneyDate.Text, 6))
    End If
            
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Стоимость суточного абонимента на 2 Месяцa
    intLat = intParkingTariff * 60
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            'Изменение текстового поля "ДеньгиДаты"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            'Установить признак  внесенной информации
    txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub
            
            'Выбрана опция - "MoneyFree"
Private Sub optMoneyFree_Click()
            'Сделать недоступным элемент управления "fraDayNight"
    fraDayNight.Enabled = False
            'Выбрать опцию "optDayNight"
    optDayNight.Value = True
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
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
    lblDate.Visible = False
            'Сделать недоступными некоторые элементы управления формы
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            'Очистить текстовое поле
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            'Сбросить признак изменений в текстовом поле
    txtMoneyDate.Tag = 0
            'Сделать доступным нажатие на кнопку "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        If cmdOK.Visible = True Then cmdOK.SetFocus
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
            'Очистить текстовoе полe
    txtParole.Text = ""
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
            
            'Обработка события "Scroll" - прокрутка для ползунка "Date"
Private Sub vsbDate_Scroll()
            'Установлен признак  внесенной информации
    If txtMoneyDate.Tag = 1 Then vsbDate_Change
    
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Date"
Private Sub vsbDate_Change()
            
            'Не установлен признак  внесенной информации
    If txtMoneyDate.Tag = 0 Then
            'Восстановление предыдущего положения ползункa
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            
            'Ползунок полосы прокрутки Даты "Уперся" сверху
    If vsbDate.Value >= vsbDate.Max And vsbDate.Tag = vsbDate.Max Then
            'Восстановление предыдущего положения ползункa
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            'Ползунoк полосы прокрутки Даты в некорректном положении
            '  (выход за пределы текущего календаря)
    If vsbDate.Value > ((frmTableCalendar.grdTableCalendar.Rows - 1) * 7 - _
    ((gRowNum - 1) * 7 + gColNum)) Then
            'Восстановление предыдущего положения ползункa
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            
            'Запомнить новое положение ползункa
    vsbDate.Tag = vsbDate.Value
            
            'Позиция знака "=" в поле "ДеньгиДата"
    intPosNum = InStr(1, Trim(txtMoneyDate.Text), "=")
            'Формирование шаблона в поле "ДеньгиДата"
    txtMoneyDate.Text = Left(Trim(txtMoneyDate.Text), intPosNum) + _
    Format(Now, "dd/mm/yyyy")
            
            'Не нулевое положение ползункa
    If vsbDate.Value > 0 Then
            'Количество парковочных дней
        intParkingDay = vsbDate.Value
            'Цикл по строкам "Таблицы календаря" (с текущего дня)
        For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Row = intRowNum
            If intRowNum = gRowNum Then
            'По столбцам "Таблицы календаря" (с текущего дня)
                intColNum = gColNum
            Else
            'По всем столбцам "Таблицы календаря"
                intColNum = 1
            End If
            'По всем столбцам "Таблицы календаря"
            For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
                frmTableCalendar.grdTableCalendar.Col = intColNum
            'Количество парковочных дней исчерпано
                If intParkingDay < 0 Then GoTo EndCycle
            'Количество парковочных дней
                intParkingDay = intParkingDay - 1
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            'Изменение  Числа и Месяца в поле "ДеньгиДаты"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) + _
                    Right(txtMoneyDate.Text, 8)
            'Переход месяца
                    If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                    And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Изменение  Месяца в поле "ДеньгиДаты"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
            'Переход года
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            'Изменение  Месяца и Года в поле "ДеньгиДаты"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                Else
            'Изменение  Числа и Месяца в поле "ДеньгиДаты"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Trim(frmTableCalendar.grdTableCalendar.Text) + _
                    Right(txtMoneyDate.Text, 8)
            'Переход месяца
                    If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                    (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Изменение  Месяца в поле "ДеньгиДаты"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
            'Переход года
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            'Изменение  Месяца и Года в поле "ДеньгиДаты"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                End If
            Next
        Next
    End If
EndCycle:
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If
    
End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Change()
            
            'Ползунок полосы прокрутки Латов "Уперся" справа
    If hsbLat.Value > hsbLat.Tag And (hsbLat.Tag * 100 + intParkingTariff) > 32000 Then
            'Восстановление предыдущего положения ползунков
        hsbSant.Value = hsbSant.Tag
        hsbLat.Value = hsbLat.Tag
    ElseIf hsbLat.Value = hsbLat.Tag Then
        Exit Sub
    End If
            'Вносимая сумма оплаты в Сантимах
    intParkingMoney = hsbLat.Value * 100 + hsbSant.Value
            'Ползунки полос прокрутки Латов и Сантимов в некорректном положении
            '  (вносимая сумма не оплачивает Целое число парковочных дней)
    If Int(intParkingMoney / intParkingTariff) * 100 <> intParkingMoney Or _
    hsbLat.Value * 100 > intParkingTariff Then
            'Ползунок двигался в сторону увеличения суммы
        If hsbLat.Value > hsbLat.Tag Then
            intParkingMoney = hsbLat.Tag * 100 + hsbSant.Tag + intParkingTariff
            'Ползунок двигался в сторону уменьшения суммы
        ElseIf hsbLat.Value < hsbLat.Tag Then
            intParkingMoney = hsbLat.Tag * 100 + hsbSant.Tag - intParkingTariff
        End If
            'Восстановление корректного положения ползунков
            hsbSant.Value = intParkingMoney - Int(intParkingMoney / 100) * 100
            hsbLat.Value = Int(intParkingMoney / 100)
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
            'Количество парковочных дней
        intParkingDay = Int(intParkingMoney / intParkingTariff)
            'Цикл по строкам "Таблицы календаря" (с текущего дня)
        For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Row = intRowNum
            If intRowNum = gRowNum Then
            'По столбцам "Таблицы календаря" (с текущего дня)
                intColNum = gColNum
            Else
            'По всем столбцам "Таблицы календаря"
                intColNum = 1
            End If
            'По всем столбцам "Таблицы календаря"
            For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
                frmTableCalendar.grdTableCalendar.Col = intColNum
            'Количество парковочных дней исчерпано
                If intParkingDay < 1 Then GoTo EndCycle
            'Количество парковочных дней
                intParkingDay = intParkingDay - 1
            'Позиция признака Опции в текущей ячейке "Таблицы календаря"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            'Изменение  Числа и Месяцав поле "ДеньгиДаты"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) + _
                    Right(txtMoneyDate.Text, 8)
              'Переход месяца
                    If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                    And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Изменение  Месяца в поле "ДеньгиДаты"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
              'Переход года
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            'Изменение  Месяца и Года в поле "ДеньгиДаты"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                Else
            'Изменение  Числа и Месяца в поле "ДеньгиДаты"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Trim(frmTableCalendar.grdTableCalendar.Text) + _
                    Right(txtMoneyDate.Text, 8)
              'Переход месяца
                    If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                    (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            'Изменение  Месяца в поле "ДеньгиДаты"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
            'Переход года
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            'Изменение  Месяца и Года в поле "ДеньгиДаты"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                End If
            Next
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
    If Int(intParkingMoney / intParkingTariff) = 0 Then
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
    If intParkingDay > 0 Then
            'Количество корректных (без переплаты) парковочных дней
        intParkingDay = Int(intParkingMoney / intParkingTariff) - intParkingDay
           'Восстановление корректной (без переплаты) суммы оплаты в Сантимах
        intParkingMoney = intParkingDay * intParkingTariff
            'Восстановление корректного положения ползунков
        hsbSant.Value = intParkingMoney - Int(intParkingMoney / 100) * 100
        hsbLat.Value = Int(intParkingMoney / 100)
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

