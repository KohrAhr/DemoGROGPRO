VERSION 5.00
Begin VB.Form frmDataEmployeIn 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmployeInData"
   ClientHeight    =   3735
   ClientLeft      =   9120
   ClientTop       =   2925
   ClientWidth     =   2580
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
   ScaleHeight     =   3735
   ScaleWidth      =   2580
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   480
      Top             =   600
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
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
      TabIndex        =   3
      Top             =   3000
      Width           =   1212
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblParole 
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
   Begin VB.Image imgEmployeIn 
      Height          =   615
      Left            =   1800
      Picture         =   "frmDataEmployeIn.frx":0000
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
      TabIndex        =   5
      Top             =   1680
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
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmDataEmployeIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             'Введенный пароль
Dim strPassword As String

            'Возврат в вызвавшую процедуру (Кнопка "OK _ +")
Private Sub cmdOK_Click()
            'Код возврата при АвтоРегистрации в "Таблице персон"
Dim intAutoRegistrCode  As Integer

            'Недоступное нажатие на кнопку "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            'Вызов процедуры-функции АвтоРегистрации
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoRegistrCode = frmTablePerson.AutoRegEmploye(txtPersonCode.Text, _
    txtInfo.Text)
            
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
        gProtocol.strProtocReserve = "AutoRegistration"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
            'Признак (Авто)Регистрация ПЕРСОНАЛЬНОГО КОДА
        frmDataEmployeIn.Tag = 1
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Отказ в АвтоРегистрация ПЕРСОНАЛЬНОГО КОДА -
            '   протоколирование события
    Else
            'Введенная ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'Введенный ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Умалчиваемая опция (Статус)
        gProtocol.strProtocStatus = gDefaultStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Invalid AutoRegEmploye"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
        frmDataEmployeIn.Tag = 2
    
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
    If frmDataEmployeIn.Tag = 1 And txtPersonCode.Tag = 1 _
    And txtInfo.Tag = 1 Then
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
    
            'Признак отказа от (Авто)Регистрации ПЕРСОНАЛЬНОГО КОДА
    If frmDataEmployeIn.Tag = 0 Then frmDataEmployeIn.Tag = 2
            'Сделать невидимой текущую форму
    frmDataEmployeIn.Visible = False
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
            
            'Сделать недоступным текстовое поле "Info"
    txtInfo.Enabled = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            'Сделать доступным текстовое поле "PersonCode"
    txtPersonCode.Enabled = True
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
           
            'Установить фокус на текстовом поле "PersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
           'Сделать недоступным нажатие на кнопку "OK _ +"
    cmdOK.MousePointer = vbNoDrop
             'Установить флаг завершения Активизации текущей формы
    frmDataEmployeIn.Tag = 1

End Sub

            'Деактивизация текущей формы
Private Sub Form_Deactivate()
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub
            
            'Загрузка текущей формы
Private Sub Form_Load()
            
            'Сделать недоступными текстовые поля "PersonCode" и "Info"
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
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
            'Длина персонального кода меньше 16-и символов
            If Len(Trim(txtPersonCode.Text)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
                txtPersonCode.Text = Left("0000000000000000", _
                16 - Len(Trim(txtPersonCode.Text))) + Trim(txtPersonCode.Text)
            End If
            'Установить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 1
            'Установить фокус на текстовом поле "txtInfo"
            If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
            txtInfo.SetFocus
            'Копирование "PersonCode"в поле "Info" с коррекцией
            txtInfo = "*" + Right(Trim(txtPersonCode), 14) + " "
            'Голубой фон текстового поля
            txtInfo.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
            txtInfo.Tag = 1
            'Вся необходимая информация имеется
            If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
                 cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                If cmdOK.Visible = True Then cmdOK.SetFocus
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

            'Процедура анализа "PersonCode" при АвтоРегистрации Служащего
            '  через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
             'Ждать завершения Активизации текущей формы
    Do While frmDataEmployeIn.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Занести ПЕРСОНАЛЬНЫЙ КОД в соответствующее
            '  текстовое поле
    txtPersonCode.Text = Trim(vntPersonCode)
            'Сделать недоступным текстовое поле "PersonCode"
    txtPersonCode.Enabled = False
            'Сделать доступным текстовое поле "Info"
    txtInfo.Enabled = True
            'Голубой фон текстового поля "PersonCode"
    txtPersonCode.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Установить фокус на текстовом поле "Info"
    If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
    txtInfo.SetFocus
    
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
        If Len(Trim(txtInfo.Text)) < 16 And Len(Trim(txtInfo.Text)) > 0 Then
            'Установить признак  изменений в текстовом поле "Info"
            txtInfo.Tag = 1
            'Вся необходимая информация имеется
            If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK _ +"
                 cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК _+"
                If cmdOK.Visible = True Then cmdOK.SetFocus
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
            If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
            txtInfo.SetFocus
            'Сделать недоступным нажатие на кнопку "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
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
            'Сбросить контроль времени ввода пароля
    tmrParoleTimeOut.Enabled = False
            'Белый фон текстового поля
    txtParole.BackColor = vbWhite
            'В (Авто)Регистрации отказано
    frmDataEmployeIn.Tag = 2
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
            'Сделать недоступными текстовые поля "PersonCode" и "Info"
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
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
            'Сделать доступными текстовые поля "PersonCode" и "Info"
            txtPersonCode.Enabled = True
            txtInfo.Enabled = True
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
