VERSION 5.00
Begin VB.Form frmDataEmployeOut 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmployeOutData"
   ClientHeight    =   3720
   ClientLeft      =   9120
   ClientTop       =   3120
   ClientWidth     =   2595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   2595
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   480
      Top             =   600
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   5
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   240
      Width           =   972
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
      TabIndex        =   2
      Top             =   3000
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "--"
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
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
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image imgEmployeOut 
      Height          =   615
      Left            =   1800
      Picture         =   "frmDataEmployeOut.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   615
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
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "frmDataEmployeOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             'Введенный пароль
Dim strPassword As String

            'Возврат в вызвавшую процедуру (Кнопка "OK_-")
Private Sub cmdOK_Click()
            'Статус
Dim strStatus As String
            'Код возврата при АвтоУдалении в "Таблице персон"
Dim intAutoDeletionCode  As Integer

            'Недоступное нажатие на кнопку "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            
            'Вызов процедуры-функции АвтоУдаления
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoDeletionCode = frmTablePerson.AutoDelEmploye(txtPersonCode.Text, _
    strStatus)
            '(Авто)Удаление ПЕРСОНАЛЬНОГО КОДА выполненo -
            '   протоколирование события
    If intAutoDeletionCode = 0 Then
            'Найденная ИНФОРМАЦИЯ
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
        gProtocol.strProtocReserve = "AutoDelete"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
            'Признак (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        Me.Tag = 1
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Отказ в АвтоУдалении ПЕРСОНАЛЬНОГО КОДА -
            '   протоколирование события
    Else
            'Найденная ИНФОРМАЦИЯ
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
        gProtocol.strProtocReserve = "Invalid AutoDelEmploye"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        frmDataAccessOut.Tag = 2
            
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
    If Me.Tag = 1 And txtPersonCode.Tag = 1 Then
            'Окно собщения с запросом изменения "Таблицы персон" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
            'Издать звуковой сигнал
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" -- """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" -- """, intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Нет"
        If strResponse = vbNo Then
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
        'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    If frmDataAccessOut.Tag = 0 Then Me.Tag = 2
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
            
            'Сделать недоступным текстовое поле "PersonCode"
    txtPersonCode.Enabled = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            'Сделать недоступным текстовoе полe "txtPersonCode"
    txtPersonCode.Enabled = False
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
            
            'Установить фокус на текстовом поле "txtParole"
    If txtParole.Enabled = True Then txtParole.SetFocus
            'Сделать недоступным нажатие на кнопку "OK _ -"
    cmdOK.MousePointer = vbNoDrop
             'Установить флаг завершения Активизации текущей формы
    Me.Tag = 1
            'Установить контроль времени ввода пароля
    tmrParoleTimeOut.Enabled = True
            'Имитировать щелчок мышью на текстовом поле "txtParole"
    txtParole_Click

End Sub

            'Деактивизация текущей формы
Private Sub Form_Deactivate()
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub
            
            'Загрузка текущей формы
Private Sub Form_Load()
            
            'Сделать недоступным текстовое поле "PersonCode"
    txtPersonCode.Enabled = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
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
    txtInfo.BackColor = vbWhite
            'Сбросить признаки изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
    cmdOK.MousePointer = vbNoDrop

End Sub

            'Процедура ввода и анализа "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            
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
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА или ИМЕНИ Служащего
            intAutoFindCode = frmTablePerson.AutoFindEmploye(txtPersonCode.Text, _
            txtInfo.Text, strStatus)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
            If intAutoFindCode = 0 Then
            'Голубой фон текстового поля
                txtInfo.BackColor = vbCyan
            'Вычислить время исключения Посетителя
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата исключения Посетителя
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            'Анализ статуса Служащего
            
            'Недопустимый статус Служащего
                If Left(Trim(strStatus), 2) <> "00" And _
                Left(Trim(strStatus), 2) <> "01" Then
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
            'Сделать недоступным нажатие на кнопку "OK_-"
                    cmdOK.MousePointer = vbNoDrop
                    GoTo PersonCodeError
            'Допустимый статус Служащего
                Else
            'Установить фокус на кнопке "ОК_-"
                    If cmdOK.Visible = True Then cmdOK.SetFocus
            'Сделать доступным нажатие на кнопку "OK_-"
                    cmdOK.MousePointer = 0
                End If
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
            txtInfo.BackColor = vbWhite
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
            txtInfo.BackColor = vbWhite
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
    txtPersonCode.BackColor = vbWhite
            'Сбросить признаки изменений в текстовом поле "Info"
    txtInfo.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
    cmdOK.MousePointer = vbNoDrop

End Sub

            'Процедура ввода и анализа "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Переход по ошибке преобразования данных
        On Error GoTo InfoError
            'Информация в допустимом диапазоне
        If Len(Trim(txtInfo.Text)) > 0 And _
        Len(Trim(txtInfo.Text)) < 17 Then
            'Установить признак  изменений в текстовом поле "Info"
            txtInfo.Tag = 1
            'Очистить текстовое поле "PersonCode"
            txtPersonCode.Text = ""
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА или ИМЕНИ Служащего
            intAutoFindCode = frmTablePerson.AutoFindEmploye(txtPersonCode.Text, _
            txtInfo.Text, strStatus)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
            If intAutoFindCode = 0 Then
            'Голубой фон текстового поля
                txtPersonCode.BackColor = vbCyan
            'Вычислить время исключения Посетителя
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата исключения Посетителя
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            'Анализ статуса Служащего
            
            'Недопустимый статус Служащего
                If Left(Trim(strStatus), 2) <> "00" And _
                Left(Trim(strStatus), 2) <> "01" Then
            'Окно собщения о неверном  статусе Посетителя - на экран
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            'Издать звуковой сигнал
                    frmDemo.BeepSound
                    If frmDemo.optEnglish = True Then
                        MsgBox "Status Error", intButtonsAndIcons, "Error"
                    Else
                        MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                    End If
            'Сбросить признак  изменений в текстовом поле "Info"
                    txtInfo.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
                    cmdOK.MousePointer = vbNoDrop
                    GoTo InfoError
            'Допустимый статус Служащего
                Else
            'Установить фокус на кнопке "ОК_-"
                    If cmdOK.Visible = True Then cmdOK.SetFocus
            'Сделать доступным нажатие на кнопку "OK_-"
                    cmdOK.MousePointer = 0
                End If
                Exit Sub
            End If
            'Персональный код в недопустимом диапазоне
InfoError:
            'Издать звуковой сигнал
            frmDemo.BeepSound
            txtInfo.Text = "Error"
            'Белый фон текстового поля
            txtInfo.BackColor = vbWhite
            txtPersonCode.BackColor = vbWhite
            'Установить фокус на текстовом поле "Info"
            If txtInfo.Enabled = True Then txtInfo.SetFocus
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            txtInfo.Text = "Error"
            'Белый фон текстового поля
            txtInfo.BackColor = vbWhite
            txtPersonCode.BackColor = vbWhite
            'Установить фокус на текстовом поле "Info"
            If txtInfo.Enabled = True Then txtInfo.SetFocus
        End If
    End If

End Sub

            'Процедура анализа "PersonCode" при АвтоУдалении
            '  Посетителя через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
             'Ждать завершения Активизации текущей формы
    Do While Me.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Занести ПЕРСОНАЛЬНЫЙ КОД в соответствующее
            '  текстовое поле
    txtPersonCode.Text = Trim(vntPersonCode)
            'Сделать недоступным текстовое поле "PersonCode"
    txtPersonCode.Enabled = False
            'Голубой фон текстового поля
    txtPersonCode.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
        
        
            'Очистить текстовое поле "Информация"
    txtInfo.Text = ""
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА или ИМЕНИ Служащего
    intAutoFindCode = frmTablePerson.AutoFindEmploye(txtPersonCode.Text, _
    txtInfo.Text, strStatus)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
    If intAutoFindCode = 0 Then
            'Вычислить время исключения Посетителя
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата исключения Посетителя
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            'Анализ статуса Посетителя Предприятия
            
            'Недопустимый статус Служащего
        If Left(Trim(strStatus), 2) <> "00" And _
        Left(Trim(strStatus), 2) <> "01" Then
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
            'Сделать недоступным нажатие на кнопку "OK_-"
             cmdOK.MousePointer = vbNoDrop
             GoTo PersonCodeError
            'Допустимый статус Служащего
        Else
            'Сделать доступным нажатие на кнопку "OK_-"
            cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_-"
            If cmdOK.Visible = True Then cmdOK.SetFocus
        End If
        Exit Function
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

End Function

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
            'Сбросить контроль времени ввода пароля
    tmrParoleTimeOut.Enabled = False
            'Белый фон текстового поля
    txtParole.BackColor = vbWhite
            'В (Авто)Удаленииии отказано
    Me.Tag = 2
            'Возврат в вызвавшую процедуру
    cmdCancel_Click
    
End Sub

            'Процедура обработки "Щелчка мыши" на поле пароля
Private Sub txtParole_Click()
            
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            'Сделать недоступными текстовое поле "PersonCode"
    txtPersonCode.Enabled = False
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
            'Пароль введен
    If KeyAscii = vbKeyReturn Then
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
            'Сделать доступным текстовoе полe "PersonCode"
            txtPersonCode.Enabled = True
            'Установить фокус на текстовом поле "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            'Пароль неверный
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
             'Белый фон текстового поля
            txtParole.BackColor = vbWhite
            'Установить фокус на текстовом поле "Parole"
            If txtParole.Enabled = True Then txtParole.SetFocus
        End If
            'Сбросить контроль времени ввода пароля
        tmrParoleTimeOut.Enabled = False
            ' "Очистка" поля пароля пробелами
        txtParole.Text = ""
    End If

End Sub
