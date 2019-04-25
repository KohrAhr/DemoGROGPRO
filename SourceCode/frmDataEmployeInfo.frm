VERSION 5.00
Begin VB.Form frmDataEmployeInfo 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmployeInfoData"
   ClientHeight    =   3975
   ClientLeft      =   7440
   ClientTop       =   2925
   ClientWidth     =   4260
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
   ScaleHeight     =   3975
   ScaleWidth      =   4260
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   12
      Tag             =   "0"
      ToolTipText     =   "Info"
      Top             =   2160
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   2280
      Top             =   240
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   10
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtParole 
      Enabled         =   0   'False
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   8
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox txtEmployeReg 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtEmployeOut 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtEmployeIn 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox lstInfo 
      Height          =   690
      ItemData        =   "frmDataEmployeInfo.frx":0000
      Left            =   720
      List            =   "frmDataEmployeInfo.frx":0002
      TabIndex        =   2
      ToolTipText     =   "Information"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox lstPersonCode 
      Height          =   690
      ItemData        =   "frmDataEmployeInfo.frx":0004
      Left            =   720
      List            =   "frmDataEmployeInfo.frx":0006
      TabIndex        =   1
      ToolTipText     =   "PersonCode"
      Top             =   1200
      Width           =   1935
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
      TabIndex        =   0
      Top             =   3360
      Width           =   1212
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
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgEmployeReg 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":0008
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgEmployeOut 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":045A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgEmployeIn 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":08A0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
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
      TabIndex        =   5
      Top             =   840
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
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgEmployeInfo 
      Height          =   615
      Left            =   3360
      Picture         =   "frmDataEmployeInfo.frx":0CE6
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "frmDataEmployeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             'Введенный пароль
Dim strPassword As String

            'Сжатие данных в "Таблице персон" (Кнопка "Pressing")
Private Sub cmdOK_Click()
            'Код возврата при Сжатии данных в "Таблице персон"
Dim intAutoPressingCode  As Integer

            'Вызов процедуры-функции Сжатия данных
            '  в "Таблице персон"
    intAutoPressingCode = frmTablePerson.AutoPresEmploye()
            
            'Присутствуют Гости
    If intAutoPressingCode = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения
        If frmDemo.optEnglish = True Then
            MsgBox ("The Visitors are Present")
        Else
            MsgBox ("Viesi ir")
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
        gProtocol.strProtocReserve = "Pressing EmployeInfo"
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
        gProtocol.strProtocReserve = "Invalid Press. Employe"
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
            
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Активизация текущей формы
Private Sub Form_Activate()
            'Текущий номер строки "Таблицы персон"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            
            'Текущая форма видимая и установлен флаг завершения ее
            '  Активизации - выйти из процедуры (для блокирования возможной
            '  повторной Активизации, чистящей текстовые поля)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            'Увеличить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessPlus
            
            'Очистить текстовые поля
    txtEmployeIn.Text = "0"
    txtEmployeOut.Text = "0"
    txtEmployeReg.Text = "0"
    txtParole.Text = ""
    txtPersonCode.Text = ""
    txtInfo.Text = ""
             'Белый фон текстового поля
    txtParole.BackColor = vbWhite
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
            'Сделать недоступными текстовые поля "Parole" и "PersonCode"
    txtParole.Enabled = False
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
            'Сделать недоступными списки "lstInfo" и "lstPersonCode"
    lstInfo.Enabled = False
    lstPersonCode.Enabled = False
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
            'Установить фокус на кнопке "Exit_Cancel"
        cmdCancel.SetFocus
        Exit Sub
    Else
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = intRowNum
            'Столбец - "Status"
            gTablePerson.Col = 2
            'Анализ статуса Посетителя
            If Left(Trim(gTablePerson.Text), 2) = "00" Or _
            Left(Trim(gTablePerson.Text), 2) = "01" Then
            'Столбец - "PersonCode"
                gTablePerson.Col = 1
            'Если строка "Таблицы персон" не удалена логически
                If gTablePerson.Text <> "Deleted" Then
            'Заполнение списка "lstPersonCode" записями из "Таблицы персон"
                    lstPersonCode.AddItem gTablePerson.Text
            'Столбец - "Person or Terminal"
                    gTablePerson.Col = 0
            'Заполнение списка "lstInfo" записями из "Таблицы персон"
                    lstInfo.AddItem gTablePerson.Text
            'Заполнение полей ИНФОРМАЦИИ о Предприятии
                    If Len(Trim(gTablePerson.Text)) < 16 Then
                        txtEmployeReg.Text = Str(CInt(txtEmployeReg.Text) + 1)
                    ElseIf Right(Trim(gTablePerson.Text), 1) = "+" Then
                        txtEmployeIn.Text = Str(CInt(txtEmployeIn.Text) + 1)
                    ElseIf Right(Trim(gTablePerson.Text), 1) = "-" Then
                        txtEmployeOut.Text = Str(CInt(txtEmployeOut.Text) + 1)
                    End If
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
            'Установить фокус на кнопке "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
            'Выбрать  элементы списков
        lstInfo.ListIndex = 0
        lstPersonCode.ListIndex = 0
            'Сделать доступным текстовое поле "txtParole"
        txtParole.Enabled = True
            
            'Установить фокус на текстовом поле "txtParole"
        If txtParole.Enabled = True Then txtParole.SetFocus
             'Установить флаг завершения Активизации текущей формы
        Me.Tag = 1
            'Установить контроль времени ввода пароля
        tmrParoleTimeOut.Enabled = True
            'Имитировать щелчок мышью на текстовом поле "txtParole"
        txtParole_Click
        Exit Sub
    End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    txtPersonCode.Text = "Unknown Error"
            'Установить фокус на кнопке "Exit_Cancel"
    cmdCancel.SetFocus
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
            'Очистить текстовые поля
    txtEmployeIn.Text = ""
    txtEmployeOut.Text = ""
    txtEmployeReg.Text = ""
    txtParole.Text = ""
    txtPersonCode.Text = ""
    txtInfo.Text = ""
             'Белый фон текстового поля
    txtParole.BackColor = vbWhite
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            'Перехват нажатия комбинаций клавиш "Alt"+ {"^" и "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Список пустой
    If lstInfo.ListCount = 0 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о пустом списке
        If frmDemo.optEnglish = True Then
            MsgBox ("The List is Empty")
        Else
            MsgBox ("Saraksts ir neaizpild.")
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
            'Очистить текстовoe поле
    txtPersonCode.Text = ""
    txtInfo.Text = ""
            'Заполнить текстовое поле "PersonCode"
    txtPersonCode.Text = lstPersonCode.Text
    txtInfo.Text = lstInfo.Text
            'Голубой фон текстового поля
    txtPersonCode.BackColor = vbCyan
    txtInfo.BackColor = vbCyan
            'Установить фокус на кнопке "Exit_Cancel"
    cmdCancel.SetFocus

End Sub

            'Выбор строки "Таблицы персон" при "щелчке" на списке "PersonCode"
Private Sub lstPersonCode_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Очистить текстовые поля
        txtPersonCode.Text = ""
        txtInfo.Text = ""
             'Белый фон текстового поля
        txtPersonCode.BackColor = vbWhite
        txtInfo.BackColor = vbWhite
            'Номер строки "Таблицы персон"
        lstInfo.ListIndex = lstPersonCode.ListIndex
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА или ИМЕНИ Служащего
        intAutoFindCode = frmTablePerson.AutoFindEmploye(lstPersonCode.Text, _
        txtInfo.Text, strStatus)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
        If intAutoFindCode = 0 Then
    
            'Анализ статуса Служащего
            
            'Недопустимый статус
            If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            'Окно собщения о неверном  статусе - на экран
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
            'Заполнить текстовые поля
            txtPersonCode.Text = lstPersonCode.Text
            txtInfo.Text = lstInfo.Text
            'Голубой фон текстового поля
            txtPersonCode.BackColor = vbCyan
            txtInfo.BackColor = vbCyan
            'Установить фокус на кнопке "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
        frmDemo.BeepSound
        txtPersonCode.Text = "Unknown Error"
            'Установить фокус на кнопке "Exit_Cancel"
        cmdCancel.SetFocus
    End If

End Sub

            'Выбор строки "Таблицы персон" при "щелчке" на списке "Info"
Private Sub lstInfo_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Статус
Dim strStatus As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Очистить текстовые поля
        txtPersonCode.Text = ""
        txtInfo.Text = ""
             'Белый фон текстового поля
        txtPersonCode.BackColor = vbWhite
        txtInfo.BackColor = vbWhite
            'Номер строки "Таблицы персон"
        lstPersonCode.ListIndex = lstInfo.ListIndex
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА или ИМЕНИ Служащего
        intAutoFindCode = frmTablePerson.AutoFindEmploye(lstPersonCode.Text, _
        txtInfo.Text, strStatus)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
        If intAutoFindCode = 0 Then
    
            'Анализ статуса Служащего
            
            'Недопустимый статус
            If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            'Окно собщения о неверном  статусе - на экран
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
            'Заполнить текстовые поля
            txtPersonCode.Text = lstPersonCode.Text
            txtInfo.Text = lstInfo.Text
            'Голубой фон текстового поля
            txtPersonCode.BackColor = vbCyan
            txtInfo.BackColor = vbCyan
            'Установить фокус на кнопке "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
            'Неизвестная ошибка
UnknownError:
            'Издать звуковой сигнал
        frmDemo.BeepSound
        txtInfo.Text = "Unknown Error"
            'Установить фокус на кнопке "Exit_Cancel"
        cmdCancel.SetFocus
    End If

End Sub

            'Процедура обработки "Щелчка мыши" на поле Персонального кода
Private Sub txtPersonCode_Click()
            
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite

End Sub

            'Процедура ввода и анализа "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            'Текущий номер строки списка "lstPersonCode"
Dim intRowNum As Integer
            
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
            
            'Цикл по всем строкам списка "lstPersonCode"
            For intRowNum = 0 To lstPersonCode.ListCount - 1 Step 1
            'Текущая строка списка
                lstPersonCode.ListIndex = intRowNum
            'Требуемая строка списка "lstPersonCode" найдена
                If Trim(lstPersonCode.Text) = Trim(txtPersonCode.Text) Then
            'Номер строки "Таблицы персон"
                    lstInfo.ListIndex = lstPersonCode.ListIndex
            'Установить фокус на кнопке "Exit_Cancel"
                    cmdCancel.SetFocus
                    Exit Sub
                End If
            Next
        
        End If
        
            'Персональный код в недопустимом диапазоне
PersonCodeError:
            'Издать звуковой сигнал
        frmDemo.BeepSound
        txtPersonCode.Text = "Error"
            'Белый фон текстового поля
        txtPersonCode.BackColor = vbWhite
            'Установить фокус на текстовом поле "PersonCode"
        If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
    End If

End Sub

            'Процедура обработки "Щелчка мыши" на поле Информации
Private Sub txtInfo_Click()
            
             'Белый фон текстового поля
    txtInfo.BackColor = vbWhite

End Sub

            'Процедура ввода и анализа "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            'Текущий номер строки списка "lstPersonCode"
Dim intRowNum As Integer
            
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Голубой фон текстового поля
        txtInfo.BackColor = vbCyan
            'Переход по ошибке преобразования данных
        On Error GoTo InfoError
            'Информация в допустимом диапазоне
        If Len(Trim(txtInfo.Text)) > 0 And _
        Len(Trim(txtInfo.Text)) < 17 Then
            
            'Цикл по всем строкам списка "lstInfo"
            For intRowNum = 0 To lstInfo.ListCount - 1 Step 1
            'Текущая строка списка
                lstInfo.ListIndex = intRowNum
            'Требуемая строка списка "lstInfo" найдена
                If InStr(1, Trim(lstInfo.Text), Trim(txtInfo.Text)) <> 0 Then
            'Номер строки "Таблицы персон"
                    lstPersonCode.ListIndex = lstInfo.ListIndex
            'Установить фокус на кнопке "Exit_Cancel"
                    cmdCancel.SetFocus
                    Exit Sub
                End If
            Next
        
        End If
        
            'Персональный код в недопустимом диапазоне
InfoError:
            'Издать звуковой сигнал
        frmDemo.BeepSound
        txtInfo.Text = "Error"
            'Белый фон текстового поля
        txtInfo.BackColor = vbWhite
            'Установить фокус на текстовом поле "Info"
        If txtInfo.Enabled = True Then txtInfo.SetFocus
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
            'В (Авто)Сжатии отказано
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
            'Сделать недоступным текстовые поля
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
            'Сделать недоступными списки "lstInfo" и "lstPersonCode"
    lstInfo.Enabled = False
    lstPersonCode.Enabled = False
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
            'Сделать доступными списки "lstInfo" и "lstPersonCode"
            lstInfo.Enabled = True
            lstPersonCode.Enabled = True
            'Сделать доступным текстовые поля
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
