VERSION 5.00
Begin VB.Form frmDataAccessInfo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessInfoData"
   ClientHeight    =   4305
   ClientLeft      =   4665
   ClientTop       =   2925
   ClientWidth     =   6960
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
   ScaleHeight     =   4305
   ScaleWidth      =   6960
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.CommandButton cmdCleaning 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Outlook"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   21
      Top             =   3600
      Width           =   615
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
      Top             =   3600
      Width           =   1212
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   17
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
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
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4680
      TabIndex        =   15
      Tag             =   "0"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Pressing"
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
      TabIndex        =   13
      Top             =   2880
      Width           =   1212
   End
   Begin VB.ListBox lstPersonCode 
      Height          =   1110
      ItemData        =   "frmDataAccessInfo.frx":0000
      Left            =   720
      List            =   "frmDataAccessInfo.frx":0002
      TabIndex        =   12
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox lstInfo 
      Height          =   1110
      ItemData        =   "frmDataAccessInfo.frx":0004
      Left            =   720
      List            =   "frmDataAccessInfo.frx":0006
      TabIndex        =   11
      ToolTipText     =   "Information"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtAccessIn 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3120
      TabIndex        =   10
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtAccessOut 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   9
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtAccessReg 
      Enabled         =   0   'False
      Height          =   288
      Left            =   6000
      TabIndex        =   8
      Tag             =   "0"
      Top             =   3840
      Width           =   735
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   0
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
         Picture         =   "frmDataAccessInfo.frx":0008
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessInfo.frx":0462
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessInfo.frx":2104
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   6360
      Picture         =   "frmDataAccessInfo.frx":2906
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   5520
      Picture         =   "frmDataAccessInfo.frx":2F38
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   4440
      Picture         =   "frmDataAccessInfo.frx":3CBE
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   5040
      Picture         =   "frmDataAccessInfo.frx":44F8
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgAccessInfo 
      Height          =   615
      Left            =   1680
      Picture         =   "frmDataAccessInfo.frx":4B36
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
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
      X2              =   6720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   1440
      Y2              =   720
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
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessInfo.frx":4F10
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4320
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4320
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
      TabIndex        =   19
      Top             =   1560
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
      TabIndex        =   18
      Top             =   240
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Image imgAccessIn 
      Height          =   375
      Left            =   2640
      Picture         =   "frmDataAccessInfo.frx":5326
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgAccessOut 
      Height          =   375
      Left            =   4080
      Picture         =   "frmDataAccessInfo.frx":576C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgAccessReg 
      Height          =   375
      Left            =   5520
      Picture         =   "frmDataAccessInfo.frx":5BB2
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
End
Attribute VB_Name = "frmDataAccessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Строка "Контроль"
Dim strChecking As String * 8
             'Введенный пароль
Dim strPassword As String
            'Текущий номер столбца "Таблицы персон"
Dim intColNumCorr As Integer

            'Сжатие данных в "Таблице персон" (Кнопка "Pressing")
Private Sub cmdOK_Click()
            'Код возврата при Сжатии данных в "Таблице персон"
Dim intAutoPressingCode  As Integer

            'Вызов процедуры-функции Сжатия данных
            '  в "Таблице персон" для Предприятия
    intAutoPressingCode = frmTablePerson.AutoPresAccess()
            
            'Ha Предприятии присутствуют Посетители , которые должны
            '   были Окончательно выйдти после оплаты
    If intAutoPressingCode = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения
        If frmDemo.optEnglish = True Then
            MsgBox ("The Persons for Exit  are Present")
        Else
            MsgBox ("Ir Persona izejai")
        End If
            
            'ИНФОРМАЦИЯ - отсутствует
        gProtocol.strProtocName = "PRESSING TabPers"
            'ПЕРСОНАЛЬНЫЙ КОД - отсутствует
        gProtocol.strProtocPersonCode = "PRESSING TabPers"
            'Статус
        gProtocol.strProtocStatus = "04 - Operator"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Pressing AccessInfo "
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак Сжатия данных в "Таблице персон"
        Me.Tag = 1
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            
            'Сжатие невозможно - протоколирование события
    ElseIf intAutoPressingCode = 2 Then
            'ИНФОРМАЦИЯ - отсутствует
        gProtocol.strProtocName = "PRESSING TabPers"
            'ПЕРСОНАЛЬНЫЙ КОД - отсутствует
        gProtocol.strProtocPersonCode = "PRESSING TabPers"
            'Статус
        gProtocol.strProtocStatus = "04 - Operator"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid Press. Access"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак отказа от Сжатия данных в "Таблице персон"
        Me.Tag = 2
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
    End If
            
End Sub
            
            'Возврат в вызвавшую процедуру (Кнопка "Cancel _ Exit")
Private Sub cmdCancel_Click()
        'Признак отказа от Сжатия данных в "Таблицe персон"
    If Me.Tag = 0 Then Me.Tag = 2
            'Сделать невидимой текущую форму
    Me.Visible = False
            
            'Очистить списки "lstInfo" и "lstPersonCode"
    lstInfo.Clear
    lstPersonCode.Clear
    
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub

            'Очистка "Таблицы персон" от Временных Клиентов (Кнопка "Cleaning")
Private Sub cmdCleaning_Click()
            'Код возврата при Очистке "Таблицы персон"
Dim intCleaningCode  As Integer

            'Вызов процедуры-функции Очистки "Таблицы персон"
    intCleaningCode = frmTablePerson.CleaningAccess()
            'Очистка выполнена - протоколирование события
    If intCleaningCode = 0 Then
            'ИНФОРМАЦИЯ - отсутствует
        gProtocol.strProtocName = "CLEANING TabPers"
            'ПЕРСОНАЛЬНЫЙ КОД - отсутствует
        gProtocol.strProtocPersonCode = "CLEANING TabPers"
            'Статус
        gProtocol.strProtocStatus = "04 - Operator"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Cleaning AccessInfo "
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак Очистки "Таблицы персон"
        Me.Tag = 1
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Очиска невозможна - протоколирование события
    ElseIf intCleaningCode = 1 Then
            'ИНФОРМАЦИЯ - отсутствует
        gProtocol.strProtocName = "CLEANING TabPers"
            'ПЕРСОНАЛЬНЫЙ КОД - отсутствует
        gProtocol.strProtocPersonCode = "CLEANING TabPers"
            'Статус
        gProtocol.strProtocStatus = "04 - Operator"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid Clean. Access"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак отказа от Очистки "Таблицы персон"
        Me.Tag = 2
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
    End If
            
End Sub
            
            'Активизация текущей формы
Private Sub Form_Activate()
            'Текущий номер строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            
            'Текущая форма видимая и установлен флаг завершения ее
            '  Активизации - выйти из процедуры (для блокирования возможной
            '  повторной Активизации, чистящей текстовые поля)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            'Увеличить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessPlus
            
            'Сделать недоступными элементы управления формы "DataAccessInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Очистить текстовые поля
    txtAccessIn.Text = "0"
    txtAccessOut.Text = "0"
    txtAccessReg.Text = "0"
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             'Белый фон текстового поля
    txtParole.BackColor = vbWhite
    
            'Очистить списки "lstInfo" и "lstPersonCode"
    lstInfo.Clear
    lstPersonCode.Clear
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
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Посетителя
            If Left(Trim(gTablePerson.Text), 2) = "10" Or _
            Left(Trim(gTablePerson.Text), 2) = "08" Or _
            Left(Trim(gTablePerson.Text), 2) = "09" Then
            'Столбец - "Person or Terminal"
                gTablePerson.Col = 0
            'Заполнение списка "lstInfo" записями из "Таблицы персон"
                lstInfo.AddItem gTablePerson.Text
            'Столбец - "PersonCode"
                gTablePerson.Col = 1
            'Заполнение списка "lstPersonCode" записями из "Таблицы персон"
                lstPersonCode.AddItem gTablePerson.Text
            'Столбец - "Reserve"
                gTablePerson.Col = 5
            'Заполнение полей ИНФОРМАЦИИ о Предприятии
                If Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
                    txtAccessIn.Text = Str(CInt(txtAccessIn.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
                    txtAccessOut.Text = Str(CInt(txtAccessOut.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
                    txtAccessReg.Text = Str(CInt(txtAccessReg.Text) + 1)
                End If
            End If
        Next
            'Список пустой
        If lstInfo.ListCount = 0 Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Вывод сообщения о пустом списке
            If frmDemo.optEnglish = True Then
                MsgBox ("The ClientList is Empty")
            Else
                MsgBox ("Klientu saraksts ir neaizpild.")
            End If
            Exit Sub
        End If
            'Выбрать  элементы списков
        lstInfo.ListIndex = 0
        lstPersonCode.ListIndex = 0
    End If
             
            
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
    If intAutoFindCode = 0 Then
            
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
            'Признак Посетителя
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            'Анализ статуса Посетителя
            
            'Недопустимый для Предприятия статус Посетителя
        If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
        Left(Trim(strStatus), 2) <> "09" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
            intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
            frmDemo.BeepSound
            If frmDemo.optEnglish = True Then
                MsgBox "Status Error", intButtonsAndIcons, "Error"
            Else
                MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
            End If
            GoTo UnknownError
        End If
            'Бесплатный Посетитель
        If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
            imgMoneyFree.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            'Постоянный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
            imgCalendar.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
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
            End If
            'Временный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Сделать видимой соответствующую "Иконку"
            imgTime.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            'Взрослый
            If Mid(lstInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            'Дети
            ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            'Конвой
            ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            'Семья
            ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
        End If
             'Установить флаг завершения Активизации текущей формы
        Me.Tag = 1
        Exit Sub
    End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
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
            'Сделать недоступными элементы управления формы "DataAccessInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Очистить текстовые поля
    txtAccessIn.Text = ""
    txtAccessOut.Text = ""
    txtAccessReg.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             'Белый фон текстового поля
    txtParole.BackColor = vbWhite

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            'Перехват нажатия комбинаций клавиш "Alt"+ {"P", "E", "^" и "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            
            'Текущая форма доступна
    If Me.Enabled = True Then
            'Альтернатива "щелчку" мыши на кнопке "P"
        If KeyCode = 80 And Shift = 4 Then
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
            
            'Список пустой
    If lstInfo.ListCount = 0 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о пустом списке
        If frmDemo.optEnglish = True Then
            MsgBox ("The ClientList is Empty")
        Else
            MsgBox ("Klientu saraksts ir neaizpild.")
        End If
    Else
            'Альтернатива "щелчку" мыши на предыдущем элементе списка
        If KeyCode = 38 And Shift = 4 And lstInfo.ListIndex <> 0 Then
            'Выбрать  элементы списков
            lstInfo.ListIndex = lstInfo.ListIndex - 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            'Альтернатива "щелчку" мыши на следующем элементе списка
        ElseIf KeyCode = 40 And Shift = 4 And _
        lstInfo.ListIndex <> lstInfo.ListCount - 1 Then
            'Выбрать  элементы списков
            lstInfo.ListIndex = lstInfo.ListIndex + 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            'Альтернатива "щелчку" мыши на первом элементе списка
        ElseIf KeyCode = 33 And Shift = 4 And lstInfo.ListIndex <> 0 Then
            'Выбрать  элементы списков
            lstInfo.ListIndex = 0
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            'Альтернатива "щелчку" мыши на последнем элементе списка
        ElseIf KeyCode = 34 And Shift = 4 And _
        lstInfo.ListIndex <> lstInfo.ListCount - 1 Then
            'Выбрать  элементы списков
            lstInfo.ListIndex = lstInfo.ListCount - 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            'Альтернатива "щелчку" мыши на текущем элементе списка
        ElseIf (KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or _
        KeyCode = 34) And Shift = 4 Then
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
        End If
        
    End If
    Exit Sub
            
DataCorrect:
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Очистить текстовое поле "ДеньгиДата" для Посетителей
    txtMoneyDate.Text = ""
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
    If intAutoFindCode = 0 Then
            
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
            'Признак Посетителя
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            'Анализ статуса Посетителя
            
            'Недопустимый для Предприятия статус Посетителя
        If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
        Left(Trim(strStatus), 2) <> "09" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
            intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
            frmDemo.BeepSound
            If frmDemo.optEnglish = True Then
                MsgBox "Status Error", intButtonsAndIcons, "Error"
            Else
                MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
            End If
            GoTo UnknownError
        End If
            'Бесплатный Посетитель
        If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
            imgMoneyFree.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            'Постоянный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
            imgCalendar.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
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
            End If
            'Временный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Сделать видимой соответствующую "Иконку"
            imgTime.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            'Взрослый
            If Mid(lstInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            'Дети
            ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            'Конвой
            ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            'Семья
            ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
        End If
        Exit Sub
    End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"

End Sub

            'Выбор строки "Таблицы персон" при "щелчке" на списке "PersonCode"
Private Sub lstPersonCode_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            'Рабочая строка
Dim strWork As String
            'Рабочая переменная
Dim intWork As Integer
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер строки "Таблицы персон"
        lstInfo.ListIndex = lstPersonCode.ListIndex
            'Сделать невидимыми "Иконки"
        imgMoneyFree.Visible = False
        imgCalendar.Visible = False
        fraDayNight.Visible = False
        imgTime.Visible = False
            'Очистить текстовое поле "ДеньгиДата"
        txtMoneyDate.Text = ""
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
        intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
        If intAutoFindCode = 0 Then
            
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
            'Признак Посетителя
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
            'Анализ статуса Посетителя
            
            'Недопустимый для Предприятия статус Посетителя
            If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
            Left(Trim(strStatus), 2) <> "09" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
                intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                frmDemo.BeepSound
                If frmDemo.optEnglish = True Then
                    MsgBox "Status Error", intButtonsAndIcons, "Error"
                Else
                    MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                End If
                GoTo UnknownError
            End If
            'Бесплатный Посетитель
            If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
                imgMoneyFree.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            'Постоянный Посетитель
            ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
                imgCalendar.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
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
                End If
            'Временный Посетитель
            ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Сделать видимой соответствующую "Иконку"
                imgTime.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            'Взрослый
                If Mid(lstInfo.Text, 5, 1) = "1" Then
                    imgHuman.Visible = True
            'Дети
                ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                    imgBaby.Visible = True
            'Конвой
                ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                    imgConvoy.Visible = True
            'Семья
                ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                    imgFamily.Visible = True
                End If
            End If
            
            'Опция "Печать Документа" установлена
            If chkDocument.Value = 1 Then
            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
                Call frmDemo.PrintDocument(gProtocol.strProtocName, _
                gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
                gProtocol.strProtocTime, gProtocol.strProtocDate, _
                gProtocol.strProtocReserve, intError)
            End If
            
            Exit Sub
        End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
    End If

End Sub
            
            'Выбор строки "Таблицы персон" при "щелчке" на списке "Info"
Private Sub lstInfo_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            'Рабочая строка
Dim strWork As String
            'Рабочая переменная
Dim intWork As Integer
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер строки "Таблицы персон"
        lstPersonCode.ListIndex = lstInfo.ListIndex
            'Сделать невидимыми "Иконки"
        imgMoneyFree.Visible = False
        imgCalendar.Visible = False
        fraDayNight.Visible = False
        imgTime.Visible = False
            'Очистить текстовое поле "ДеньгиДата"
        txtMoneyDate.Text = ""
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
        intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
        If intAutoFindCode = 0 Then
            
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
            'Признак Окончательного выхода Посетителя
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
            'Анализ статуса Посетителя Предприятия
            
            'Недопустимый для Предприятия статус Посетителя
            If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
            Left(Trim(strStatus), 2) <> "09" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
                intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                frmDemo.BeepSound
                If frmDemo.optEnglish = True Then
                    MsgBox "Status Error", intButtonsAndIcons, "Error"
                Else
                    MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                End If
                GoTo UnknownError
            End If
            'Бесплатный Посетитель
            If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
                imgMoneyFree.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            'Постоянный Посетитель
            ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
                imgCalendar.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
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
                End If
            'Временный Посетитель
            ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Сделать видимой соответствующую "Иконку"
                imgTime.Visible = True
            'Сделать невидимыми соответствующие "Иконки"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            'Взрослый
                If Mid(lstInfo.Text, 5, 1) = "1" Then
                    imgHuman.Visible = True
            'Дети
                ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                    imgBaby.Visible = True
            'Конвой
                ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                    imgConvoy.Visible = True
            'Семья
                ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                    imgFamily.Visible = True
                End If
            End If
            
            'Опция "Печать Документа" установлена
            If chkDocument.Value = 1 Then
            'Печать Документа (Пропуска со Штрих-кодом, Простого
            '  Чека и/или Кассового Чека)
                Call frmDemo.PrintDocument(gProtocol.strProtocName, _
                gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
                gProtocol.strProtocTime, gProtocol.strProtocDate, _
                gProtocol.strProtocReserve, intError)
            End If
            
            Exit Sub
        End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
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
            'Сделать доступной опцию "Document"
            imgDocument.Enabled = True
            chkDocument.Enabled = True
            'Пароль неверный
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Сделать недоступной опцию "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
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


