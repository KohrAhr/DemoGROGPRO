VERSION 5.00
Begin VB.Form frmDataAccessServ 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessServData"
   ClientHeight    =   3105
   ClientLeft      =   4665
   ClientTop       =   2745
   ClientWidth     =   7080
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
   ScaleHeight     =   3105
   ScaleWidth      =   7080
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Frame fraStatus 
      Caption         =   "????"
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
      Height          =   2055
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optNight 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDay 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessServ.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   99
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   320
      TabIndex        =   6
      Top             =   960
      Width           =   1452
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkDocument 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4200
      TabIndex        =   2
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Left            =   4200
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6960
      Y1              =   2160
      Y2              =   2160
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
      Left            =   6480
      TabIndex        =   19
      Top             =   960
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
      Left            =   4680
      TabIndex        =   18
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Image imgAccessServ 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessServ.frx":0802
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image imgDocument 
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessServ.frx":0BC4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   1440
      Y2              =   2400
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1440
      Y2              =   1440
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
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
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
      TabIndex        =   17
      Top             =   1080
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
      TabIndex        =   16
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDataAccessServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Строка "Контроль" для Предприятия
Dim strChecking As String * 8
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Вносимая сумма оплаты в Сантимах
Dim intAccessMoney As Integer
            'Количество дней действия пропуска
Dim intAccessDay As Integer
            'Тариф одного дня (Сутки)
Dim intAccessTariffFull As Integer
            'Тариф одного дня (День)
Dim intAccessTariffDay As Integer
            'Тариф одного дня (Ночь)
Dim intAccessTariffNight As Integer
            'Тариф Предприятия (переменная для рассчетов)
Dim intAccessTariff As Integer
            'Текущая строка "Таблицы календаря"
Dim intRowNum As Integer
            'Текущая столбец "Таблицы календаря"
Dim intColNum As Integer
            'День, соответствующий Дате Регистрации
            '  Клиента Предприятия (или последнему парковочному дню)
Dim intDayReg As Integer
            'Месяц, соответствующий Дате Регистрации
            '  Клиента Предприятия (или последнему парковочному дню)
Dim intMonthReg As Integer
            'Год, соответствующий Дате Регистрации
            '  Клиента Предприятия (или последнему парковочному дню)
Dim intYearReg As Integer
            'Номер позиции заданного символа в строке
Dim intPosNum As Integer

            'Перехват нажатия комбинаций клавиш "Alt"+ {"OK" и "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Текущая форма доступна
    If Me.Enabled = True Then
            'Альтернатива "щелчку" мыши на кнопке "OK"
        If KeyCode = 79 And Shift = 4 Then
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
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Голубой фон текстового поля
        txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопку "OK_+"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub
            'Возврат в вызвавшую процедуру (Кнопка "OK")
Private Sub cmdOK_Click()
            'Статус
Dim strStatus As String
            'Дата последнего дня (и Время)
            '  коррекции информации о Посетителе
Dim strDate As String
            'Время регистрации Посетителя
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время регистрации Посетителя
Dim strHour As String
Dim strMinute As String
            'Признак присутствия Посетителя \ 0-вошел \1-вышел \ 2-регистрация
Dim strCarPresent As String * 1
            'Признак Посетителя ("Е" - Окончательно вышел; "D" - Дневной тариф;
            '  "N" - Ночной тариф; "Другой символ - Суточный тариф)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоКоррекции в "Таблице персон"
Dim intAutoCorrectionCode  As Integer
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer

            'Недоступное нажатие на кнопку "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            'Не нулевая сумма требуемой оплаты
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            'Обнулить поле "Tag" формы "frmMinus"
        frmMinus.Tag = 0
            'Вывести на экран форму "frmMinus" с уровнем модальности 1
        frmMinus.Show 1
            'Отказ от оплаты и от (Авто)Коррекции ПЕРСОНАЛЬНОГО КОДА
        If frmMinus.Tag = "Exit" Then
            'Возврат в вызвавшую процедуру
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            'Очистка строки "Контроль"
    strChecking = ""
            'Вычислить время коррекции информации о Посетителе
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
            
            'Признак перерегистрации Посетителя
    strCarPresent = "2"
            'Признак Посетителя
    strExpander = "P"
            'Посетитель с Дневным тарифом допуска
    If optDay.Value = True Then
        strExpander = "D"
            'Посетитель с Ночным тарифом допуска
    ElseIf optNight.Value = True Then
        strExpander = "N"
    End If
    
            'Дата последнего оплаченного дня
    strDate = Mid(Trim(txtMoneyDate.Text), 11)
    If Mid(Trim(strDate), 3, 1) = "." Then
        strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
        Mid(Trim(strDate), 7, 4)
    Else
        strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
        Mid(Trim(strDate), 6, 4)
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
            
            'Признаки Посетителя и его присутствия на Предприятии
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            'Вызов процедуры-функции АвтоКоррекции для данного
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoCorrectionCode = frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
    txtInfo.Text, strChecking, strStatus)
            '(Авто)Коррекция для даннного ПЕРСОНАЛЬНОГО КОДА выполнена -
            '   протоколирование события
    If intAutoCorrectionCode = 0 Then
            'Введенная ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'Введенный ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Статус
        gProtocol.strProtocStatus = strStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "AutoCorAcce " + Left(Trim(txtMoneyDate.Text), 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            'Признак (Авто)Коррекции для данного ПЕРСОНАЛЬНОГО КОДА
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
        
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Отказ от АвтоКоррекции для данного ПЕРСОНАЛЬНОГО КОДА -
            '   протоколирование события
    Else
            'Введенная ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'Введенный ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Статус
        gProtocol.strProtocStatus = strStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid AutoCorAccess"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtInfo.BackColor = vbWhite
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Коррекции для данного ПЕРСОНАЛЬНОГО КОДА
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
            
            'Были не сохраненные изменения в текстовых полях текущей формы
    If txtPersonCode.Tag = 1 And _
    (txtInfo.Tag = 1 Or txtMoneyDate.Tag = 1) Then
            'Окно собщения с запросом изменения "Таблицы персон" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
            'Издать звуковой сигнал
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" OK """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" OK """, intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
        'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
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
            
            'Сделать недоступными элементы управления формы "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtMoneyDate.BackColor = vbWhite
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Установить фокус на текстовом поле "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
            'Сделать недоступным нажатие на кнопку "OK"
    cmdOK.MousePointer = vbNoDrop
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
            'Сделать недоступными элементы управления формы "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми элементы управления формы "DataAccessServ"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Тариф одного дня (Сутки)
    intAccessTariffFull = gAccessDN
            'Тариф одного дня (День)
    intAccessTariffDay = gAccessD
            'Тариф одного дня (Ночь)
    intAccessTariffNight = gAccessN
            'Сделать недоступным нажатие на кнопку "OK"
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
    txtInfo.BackColor = vbWhite
    txtMoneyDate.BackColor = vbWhite
            'Очистить текстовое поле "Информация"
    txtInfo.Text = ""
            'Очистить текстовое поле "ДеньгиДата"
    txtMoneyDate.Text = ""
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            'Сделать недоступными элементы управления формы "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сделать недоступным нажатие на кнопку "OK"
    cmdOK.MousePointer = vbNoDrop

End Sub

            'Процедура ввода и анализа "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Строка "Информация"
Dim strInfo As String
            'Статус
Dim strStatus As String
            'Время коррекции информации о Посетителе
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время коррекции
            '  информации о Посетителе
Dim strHour As String
Dim strMinute As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            'Текущий номер строки "Таблицы протокола"
Dim intRowNum As Integer
            'Текущая сумма "Z_Отчета"
Dim lngZ_Report As Long
            'Текущий указатель "Z_Отчета"
Dim strZ_Report As String
            'Номер текущей строки в "Системной таблице"
Dim intRowNumSys As Integer
            'Код возврата при сохранении "Системной таблицы"
Dim intSaveTableSystem As Integer
            
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtPersonCode.BackColor = vbCyan
            'Переход по ошибке преобразования данных
        On Error GoTo PersonCodeError
            'Персональный код в допустимом диапазоне
        If Len(Trim(txtPersonCode.Text)) > 0 And _
        Len(Trim(txtPersonCode.Text)) < 17 Then
            'Длина персонального кода меньше 16-и символов
            If Len(Trim(txtPersonCode.Text)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
                txtPersonCode.Text = Left("0000000000000000", _
                16 - Len(Trim(txtPersonCode.Text))) + Trim(txtPersonCode.Text)
            End If
            'Установить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 1
            'Очистить текстовое поле "Информация"
            txtInfo.Text = ""
            'Очистить текстовое поле "ДеньгиДата"
            txtMoneyDate.Text = ""
            'Сделать невидимой метку текстового поля "txtMoneyDate"
            lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
            imgCalendar.Visible = False
            fraDayNight.Visible = False
            'Сделать недоступными элементы управления формы "DataAccessServ"
            lblInfo.Enabled = False
            txtInfo.Enabled = False
            lblLat0.Enabled = False
            lblLat320.Enabled = False
            hsbLat.Enabled = False
            
            'Запрошен "Z_Отчет"
            If Right(txtPersonCode.Text, 8) = "Z_Report" Then
            'Обнулить текущую сумму "Z_Отчета"
                lngZ_Report = 0
            'Цикл по всем строкам "Таблицы протокола"
                For intRowNum = 1 To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
                    Get gProtocFileNum, intRowNum, gProtocol
            'Событие протокола:
            '   - Регистрация платного Клиента Автостоянки или
            '   - Исключение платного Клиента Автостоянки
            '   - Коррекция Постоянного Клиента Автостоянки
            '   - Регистрация платного Посетителя Предприятия или
            '   - Исключение платного Посетителя Предприятия
            '   - Коррекция Постоянного Посетителя Предприятия
                    If ((Left(Trim(gProtocol.strProtocStatus), 2) = "05" Or _
                    Left(Trim(gProtocol.strProtocStatus), 2) = "06") And _
                    (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Or _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "05" And _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoCorPark") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "08" Or _
                    Left(Trim(gProtocol.strProtocStatus), 2) = "09") And _
                    (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Or _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "08" And _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoCorAcce")) And _
                    Left(gProtocol.strProtocName, 1) <> "@" Then
            'До этой точки включительно "Z_Отчет" уже выводился
                        If gZ_Report = Trim(gProtocol.strProtocTime) + _
                        Left(Trim(gProtocol.strProtocDate), 6) Then
            'Обнулить текущие сумму "Z_Отчета"
                            lngZ_Report = 0
            'Очистить указатель текущей точки "Z_Отчета"
                            strZ_Report = ""
                        Else
            'Корректировать текущие сумму и указатель точки "Z_Отчета"
                            If Mid(gProtocol.strProtocReserve, 13, 3) <> "   " Then
                                lngZ_Report = lngZ_Report + _
                                Mid(Trim(gProtocol.strProtocReserve), 13, 3) * 100 + _
                                Mid(Trim(gProtocol.strProtocReserve), 17, 2)
                                strZ_Report = Trim(gProtocol.strProtocTime) + _
                                Left(Trim(gProtocol.strProtocDate), 6)
                            End If
                        End If
                    End If
                Next
            'Изменение текстового поля "ДеньгиДаты"
                txtMoneyDate.Text = Trim(Str(Int(lngZ_Report / 100))) + " Ls " + _
                Trim(Str(lngZ_Report - Int(lngZ_Report / 100) * 100)) + " s"
               
            'Опция "Печать Документа" установлена
        If chkDocument.Value = 1 Then
            'Сделать недоступной кнопку "Exit_Cancel"
                cmdCancel.Enabled = False
            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
            Call frmDemo.PrintZReport(gProtocol.strProtocName, _
            gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
            gProtocol.strProtocTime, gProtocol.strProtocDate, _
            gProtocol.strProtocReserve, txtMoneyDate.Text, strZ_Report)
            'Сделать доступной кнопку "Exit_Cancel"
                cmdCancel.Enabled = True
        End If
            
            'Запомнить новую точку "Z_Отчета"
                If strZ_Report <> "" Then
                    gZ_Report = strZ_Report
            'Фиксированный столбец "Системной таблицы" (Объект)
                    frmTableSystem.grdTableSystem.Col = 0
            'Цикл по всем нефиксированным строкам "Системной таблицы"
                    For intRowNumSys = 1 To _
                    frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка  "Системной таблицы"
                        frmTableSystem.grdTableSystem.Row = intRowNumSys
            'Строка "Системной таблицы" с указателем точки "Z_Отчета"
                        If Trim(frmTableSystem.grdTableSystem.Text) = "Z_Report" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
                            frmTableSystem.grdTableSystem.Col = 1
            'Дата и время формирования последнего "Z_Отчета"
                            frmTableSystem.grdTableSystem.Text = gZ_Report
                            Exit For
                        End If
                    Next
            'Сохранить новую 'Системную таблицу"
                    intSaveTableSystem = frmTableSystem.SaveTableSystem()
                End If
                
            'Установить фокус на кнопке "Exit_Cancel"
                cmdCancel.SetFocus
                Exit Sub
            
            'Запрошен "Персональный Код"
            Else
            'Вызов процедуры-функции АвтоПоиска ПЕРСОНАЛЬНОГО КОДА
                intAutoFindCode = _
                frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
                strInfo, strStatus, strChecking)
            End If
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
            If intAutoFindCode = 0 Then
            
            'Анализ статуса Посетителя Предприятия
            
            'Недопустимый для Предприятия статус Посетителя
                If Left(Trim(strStatus), 2) <> "08" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                    frmDemo.BeepSound
                    If frmDemo.optEnglish = True Then
                        MsgBox "Status Error", intButtonsAndIcons, "Error"
                    Else
                        MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                    End If
            'Сбросить признак  изменений в текстовом поле "PersonCode"
                    txtPersonCode.Tag = 0
                    GoTo PersonCodeError
                End If
            
            'Заполнить текстовое поле "Информация"
                txtInfo.Text = strInfo
            
            'Распаковка строки "Контроль"
                Call frmTablePerson.UnPack(strDate, strChecking)
            
            'Отображение распакованной строки "Контроль"
                txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
                Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/входа/выхода Посетителя
                If Mid(Trim(strChecking), 7, 1) = "0" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
                ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
                ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
                txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
                End If

            'Вычислить время коррекции информации о Посетителе
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
            'Дата коррекции информации о Посетителе
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            'Сделать видимой соответствующую "Иконку"
                imgCalendar.Visible = True
            'Установить и сделать видимым соответствующий
            '  Признак Посетителя
                If Right(Trim(strChecking), 1) = "D" Then
            'Дневной тариф допуска
                    optDay.Value = True
                ElseIf Right(Trim(strChecking), 1) = "N" Then
            'Ночной тариф допуска
                    optNight.Value = True
                ElseIf Right(Trim(strChecking), 1) <> "D" And _
                Right(Trim(strChecking), 1) <> "N" Then
            'Суточный тариф допуска
                    optDayNight.Value = True
                End If
            'Посетитель не вышел Окончательно с Предприятия
                If Right(Trim(strChecking), 1) <> "E" Then
                    fraDayNight.Visible = True
            'Посетитель вышел Окончательно с Предприятия
                Else
                    Exit Sub
                End If
            'Формирование шаблона в поле "ДеньгиДаты"
                txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Определение тарифа Предприятия (переменной для рассчетов)
            'Дневной тариф допуска
                If optDay.Value = True Then
                    intAccessTariff = intAccessTariffDay
            'Ночной тариф допуска
                ElseIf optNight.Value = True Then
                    intAccessTariff = intAccessTariffNight
            'Суточный тариф допуска
                ElseIf optDayNight.Value = True Then
                    intAccessTariff = intAccessTariffFull
                End If
            'Сделать видимой метку текстового поля "txtMoneyDate"
                lblMoneyDate.Visible = True
            'Сделать доступными элементы управления формы "DataAccessServ"
                lblInfo.Enabled = True
                txtInfo.Enabled = True
                lblLat0.Enabled = True
                lblLat320.Enabled = True
                hsbLat.Enabled = True
    
            'Дата последнего дня действия пропуска
                strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            'Вызов процедуры-функции вычисления Даты
            '  последнего дня действия пропуска
                intDayReg = Left(strDate, 2)
                intMonthReg = Mid(strDate, 4, 2)
                intYearReg = Right(strDate, 4)
            'Вычисление Даты , соответствующей дню,
            '  следующему за последним днем действия пропуска
                frmTableCalendar.comCalendar.Day = intDayReg
                frmTableCalendar.comCalendar.Month = intMonthReg
                frmTableCalendar.comCalendar.Year = intYearReg
                frmTableCalendar.comCalendar.NextDay
                intDayReg = frmTableCalendar.comCalendar.Day
                intMonthReg = frmTableCalendar.comCalendar.Month
                intYearReg = frmTableCalendar.comCalendar.Year
            'Белый фон текстового поля
                txtInfo.BackColor = vbWhite
                txtMoneyDate.BackColor = vbWhite
            'Установить фокус на текстовом поле "txtPersonCode"
                If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
                Exit Sub
            End If
            
            'Персональный код в недопустимом диапазоне или другая ошибка
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
            'Сделать недоступным нажатие на кнопку "OK_-"
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
            'Сделать недоступным нажатие на кнопку "OK_-"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub
            'Процедура обработки "Щелчка мыши" на поле Информации
Private Sub txtInfo_Click()
            'Белый фон текстового поля
    txtInfo.BackColor = vbWhite
            'Сбросить признак  изменений в текстовом поле "Info"
    txtInfo.Tag = 0
            'Отсутствует необходимая информация
    If txtMoneyDate.Tag = 0 Then
            'Сделать недоступным нажатие на кнопку "OK"
         cmdOK.MousePointer = vbNoDrop
    End If

End Sub
            
            'Процедура ввода и анализа текстового поля "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Информация в допустимом диапазоне
        If Len(Trim(txtInfo.Text)) < 17 And Len(Trim(txtInfo.Text)) > 0 Then
            'Установить признак  изменений в текстовом поле "Info"
            txtInfo.Tag = 1
            'Сделать доступным нажатие на кнопку "OK"
            cmdOK.MousePointer = 0
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
            'Отсутствует необходимая информация
            If txtMoneyDate.Tag = 0 Then
            'Сделать недоступным нажатие на кнопку "OK"
                cmdOK.MousePointer = vbNoDrop
            End If
        End If
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
            'Ползунки полос прокрутки Латов и Сантимов в некорректном
            '  положении (вносимая сумма не оплачивает Целое число дней)
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
    txtMoneyDate.Text = strDate
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
            'Количество дней действия пропуска
        intAccessDay = Int(intAccessMoney / intAccessTariff)
            'Восстановление ИСХОДНОГО сотояния "Календаря"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            'Цикл по Дням "Календаря" (от последнего
            '  дня действия пропуска +1)
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
        txtMoneyDate.Text = strDate
            'Формирование шаблона в поле "ДеньгиДата"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопку "OK"
        If txtInfo.Tag = 0 Then
            cmdOK.MousePointer = vbNoDrop
        End If
    End If
EndCycle:
            'Голубой фон текстового поля
    txtMoneyDate.BackColor = vbCyan
            'Недостаточная оплата одного дня
    If Int(intAccessMoney / intAccessTariff) = 0 Then
           'Отмена внесенной информации
       txtMoneyDate.Tag = 0
           'Дата
       txtMoneyDate.Text = strDate
           'Формирование шаблона в поле "ДеньгиДата"
       txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопку "OK"
        If txtInfo.Tag = 0 Then
            cmdOK.MousePointer = vbNoDrop
        End If
    End If
            'Переплата (возможна оплата только до конца года)
    If intAccessDay > 0 Then
            'Количество корректных (без переплаты) дней действия пропуска
        intAccessDay = Int(intAccessMoney / intAccessTariff) - intAccessDay
           'Восстановление корректной (без переплаты) суммы оплаты в Сантимах
        intAccessMoney = intAccessDay * intAccessTariff
            'Восстановление корректного положения ползунков
        hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
        hsbLat.Value = Int(intAccessMoney / 100)
        hsbLat_Change
    End If
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK"
        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If
    
End Sub


