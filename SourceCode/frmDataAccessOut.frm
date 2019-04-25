VERSION 5.00
Begin VB.Form frmDataAccessOut 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessOutData"
   ClientHeight    =   3960
   ClientLeft      =   3000
   ClientTop       =   2745
   ClientWidth     =   8715
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
   ScaleHeight     =   3960
   ScaleWidth      =   8715
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.CommandButton cmdOutConst 
      BackColor       =   &H00FF0000&
      Caption         =   "Sant=""50"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOutFree 
      BackColor       =   &H00FF0000&
      Caption         =   "Ls=""000,00"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   16
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   15
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4200
      TabIndex        =   14
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   6480
      Top             =   120
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   10
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   840
      Width           =   972
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
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
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
      TabIndex        =   12
      Top             =   3240
      Width           =   1212
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      LargeChange     =   320
      Left            =   4920
      Max             =   320
      SmallChange     =   320
      TabIndex        =   9
      Top             =   2280
      Width           =   1452
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   6960
      Max             =   99
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
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
      Height          =   3375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
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
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
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
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":24A4
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   0
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   2040
      Picture         =   "frmDataAccessOut.frx":28FE
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   1200
      Picture         =   "frmDataAccessOut.frx":2F30
      Stretch         =   -1  'True
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   120
      Picture         =   "frmDataAccessOut.frx":3CB6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   720
      Picture         =   "frmDataAccessOut.frx":44F0
      Stretch         =   -1  'True
      Top             =   840
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
      TabIndex        =   22
      Top             =   240
      Width           =   495
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
      TabIndex        =   21
      Top             =   1920
      Width           =   495
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
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   7080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   8520
      X2              =   4080
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
      Left            =   5400
      Picture         =   "frmDataAccessOut.frx":4B2E
      Stretch         =   -1  'True
      Top             =   240
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
      Left            =   4920
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   1440
      Y2              =   600
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   6240
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Image imgAccessOut 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessOut.frx":4F44
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   8520
      X2              =   8520
      Y1              =   1560
      Y2              =   3720
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
      TabIndex        =   19
      Top             =   2280
      Width           =   135
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
      TabIndex        =   18
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   8520
      Y1              =   3720
      Y2              =   3720
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
      Left            =   4200
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5280
      X2              =   6240
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmDataAccessOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Строка "Контроль" для Предприятия
Dim strChecking As String * 8
            'Подстрока "Контроль" для Предприятия
Dim strCheckingInfo As String * 8
            'Вносимая сумма оплаты в Сантимах
Dim lngAccessMoney As Long
            'Количество дней посещения
Dim intAccessDay As Integer
            'Тариф одного дня посещения (Сутки)
Dim intAccessTariffFull As Integer
            'Тариф одного дня посещения (День)
Dim intAccessTariffDay As Integer
            'Тариф одного дня посещения (Ночь)
Dim intAccessTariffNight As Integer
            'Тариф одного часа посещения (переменная для рассчетов)
Dim intAccessTariffHour As Integer
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
             'Введенный пароль
Dim strPassword As String

            'Перехват нажатия комбинаций клавиш "Alt"+ {"--", "E" , "L" и "S"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Текущая форма доступна
    If Me.Enabled = True Then
            'Альтернатива "щелчку" мыши на кнопке "--"
        If KeyCode = 189 And Shift = 4 Then
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
            'Альтернатива "щелчку" мыши на кнопке "0 Ls"
        ElseIf KeyCode = 76 And Shift = 4 Then
            If cmdOutFree.Visible = True Then
                Call cmdOutFree_Click
                Exit Sub
            End If
            'Альтернатива "щелчку" мыши на кнопке "XX San"
        ElseIf KeyCode = 83 And Shift = 4 Then
            If cmdOutConst.Visible = True Then
                Call cmdOutConst_Click
                Exit Sub
            End If
        End If
    End If
    
End Sub

            'Возврат в вызвавшую процедуру (Кнопка "0" Ls)
Private Sub cmdOutFree_Click()
            
            'Недоступное нажатие на кнопку "0 Ls"
    If cmdOutFree.MousePointer = vbNoDrop Then Exit Sub

            'Изменение текстового поля "ДеньгиДаты"
    txtMoneyDate.Text = "000,00 Ls=" + Mid(txtMoneyDate.Text, 11)

End Sub

            'Возврат в вызвавшую процедуру (Кнопка "XX" San)
Private Sub cmdOutConst_Click()
            
            'Недоступное нажатие на кнопку "XX San"
    If cmdOutConst.MousePointer = vbNoDrop Then Exit Sub

            'Изменение текстового поля "ДеньгиДаты"
    If Int(gTariffConst / 100) < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gTariffConst / 100) < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gTariffConst / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            'Изменение текстового поля "ДеньгиДаты"
    If gTariffConst - Int(gTariffConst / 100) * 100 < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(Str(gTariffConst - Int(gTariffConst / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(Str(gTariffConst - Int(gTariffConst / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If

End Sub

            'Переключить признак печати документа - "Document"
Private Sub chkDocument_Click()
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            'Голубой фон текстового поля
        txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
        cmdOK.MousePointer = 0
        cmdOutFree.MousePointer = 0
        cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If

End Sub

            'Возврат в вызвавшую процедуру (Кнопка "OK_-")
Private Sub cmdOK_Click()
            'Статус
Dim strStatus As String
            'Код возврата при АвтоУдалении в "Таблице персон"
Dim intAutoDeletionCode  As Integer
            'Номер элемента в массиве "Таблицы терминалов", хранящего Запрос
Dim intRequest As Integer
            'Адрес контроллера
Dim vntAddr As Variant
            'Признак ошибки при печати штрих-кода и др.
Dim intError As Integer

            'Недоступное нажатие на кнопку "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            'Если Посетитель Временный
    If imgTime.Visible = True Then
            'Обнулить поле "Tag" формы "frmLease"
        frmLease.Tag = 0
            'Вывести на экран форму "frmLease" с уровнем модальности 1
        frmLease.Show 1
            'Коррекция в поле "txtInfo" информации о ПРОКАТЕ ИНВЕНТАРЯ
        If frmLease.Tag <> "Exit" Then _
        txtInfo.Text = Left(CStr(frmLease.Tag), 4) + Mid(txtInfo.Text, 5)
    End If

            'Не нулевая сумма требуемой оплаты
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            'Обнулить поле "Tag" формы "frmMinus"
        frmMinus.Tag = 0
            'Вывести на экран форму "frmMinus" с уровнем модальности 1
        frmMinus.Show 1
            'Отказ от оплаты и от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        If frmMinus.Tag = "Exit" Then
            'Возврат в вызвавшую процедуру
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            'Сделать недоступными кнопки "OK _ +" и "Cancel _ Exit"
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
            
            'Взрослый
    If imgHuman.Visible = True Then
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
        gAccessMoneyCell = gAccessMoneyCellHuman
            'Дети
    ElseIf imgBaby.Visible = True Then
            'Входной тариф Предприятия для Детей (для временных Клиентов)
        gAccessMoneyCell = gAccessMoneyCellBaby
            'Конвой
    ElseIf imgConvoy.Visible = True Then
            'Входной тариф Предприятия для Конвоя (для временных Клиентов)
        gAccessMoneyCell = gAccessMoneyCellConvoy
            'Семья
    ElseIf imgFamily.Visible = True Then
            'Входной тариф Предприятия для Семьи (для временных Клиентов)
        gAccessMoneyCell = gAccessMoneyCellFamily
            'Не Временный Посетитель
    ElseIf imgTime.Visible = False Then
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
        gAccessMoneyCell = gAccessMoneyCellHuman
    End If
            'Временный Посетитель - необходимо продлить время и дату
            '  удаления Посетителя, до которых ему будет разрешен Выход
    If imgTime.Visible = True Then Call Prolong(strStatus)
            
            'Вызов процедуры-функции АвтоУдаления ПЕРСОНАЛЬНОГО КОДА
    intAutoDeletionCode = frmTablePerson.AutoDelAccess(txtPersonCode.Text, strStatus)
                                                                                             
            '(Авто)Удаление ПЕРСОНАЛЬНОГО КОДА выполненo -
            '   протоколирование события
    If intAutoDeletionCode = 0 Then
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
        If gAccessDeletion = 1 Then
            'ФИЗИЧЕСКОЕ Удаление
            gProtocol.strProtocReserve = "AutoDelAcce " + Left(Trim(txtMoneyDate.Text), 9)
        Else
            'ЛОГИЧЕСКОЕ Удаление
            gProtocol.strProtocReserve = "LogDelAcce " + Left(Trim(txtMoneyDate.Text), 9)
        End If
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtMoneyDate.Tag = 0
            'Признак (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
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
            '   выполнение процедур невозможно), Временный Клиент Предприятия,
            '   установлен признак НЕМЕДЛЕННОЕ Удаление и установлен индекс
            '   выходного терминала - открыть терминал
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        imgTime.Visible = True And gAccessDeletion = 1 And gTermOut <> -1 Then
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  выполнено Исключение Клиента Предприятия
            '  и установлена Опция "Физическое удаление"
            If frmDemo.cmdOpen(gTermOut).Tag = 0 And Me.Tag = 1 And _
            gAccessDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Предприятия
                frmDemo.imgAccessInData(gTermOut).Enabled = False
                frmDemo.imgAccessOutData(gTermOut).Enabled = False
                frmDemo.imgAccessInfoData(gTermOut).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(gTermOut))))
                frmDemo.cmdOpen(gTermOut).Tag = vntAddr
                frmDemo.cmdOpen(gTermOut).Caption = "Addr=" + CStr(vntAddr)
            'Метка "N_?" - (зеленый фон)
                frmDemo.lblInform(gTermOut).BackColor = vbGreen
            'Включить контроль "TimeOut" электронной "Кнопки"
                frmDemo.tmrButton(gTermOut).Enabled = True
            'Имитировать нажатие электронной "Кнопки"
                Call frmDemo.OpenBarrier(gTermOut)
            End If
        End If
        
        
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
            'Отказ в АвтоУдалении ПЕРСОНАЛЬНОГО КОДА -
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
        gProtocol.strProtocReserve = "Invalid AutoDelAccess"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
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
    ((txtPersonCode.Tag = 1 And imgMoneyFree.Visible = True) Or _
    (txtPersonCode.Tag = 1 And imgTime.Visible = True) Or _
    (txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1)) Then
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
            'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
            Me.Tag = 2
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
                'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
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
            
            'Сделать доступным элемент управления формы "DataAccessOut"
    txtPersonCode.Enabled = True
            'Сделать недоступными элементы управления формы "DataAccessOut"
    lblParole.Enabled = False
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимыми кнопки разрешения выхода Специальных Клиентов
    cmdOutFree.Visible = False
    cmdOutConst.Visible = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Сделать невидимыми соответствующие "Иконки"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Установить фокус на текстовом поле "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
            'Сделать недоступным нажатие на кнопку "OK _ -"
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
            'Сделать недоступными элементы управления формы "DataAccessOut"
    lblParole.Enabled = False
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми элементы управления формы "DataAccessOut"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Очистить текстовые поля
    txtPersonCode.Text = ""
    txtParole.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Тариф одного дня посещения (Сутки)
    intAccessTariffFull = gAccessDN
            'Тариф одного дня посещения (День)
    intAccessTariffDay = gAccessD
            'Тариф одного дня посещения (Ночь)
    intAccessTariffNight = gAccessN
            'Сделать недоступным нажатие на кнопку "OK _ -"
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
            'Очистить текстовое поле "Информация"
    txtInfo.Text = ""
            'Очистить текстовое поле "ДеньгиДата"
    txtMoneyDate.Text = ""
            'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            'Сделать невидимыми соответствующие "Иконки"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            'Сделать недоступными элементы управления формы "DataAccessOut"
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            'Сбросить признаки изменений в текстовых полях
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            'Запомнить начальное положение ползунков для полос прокрутки
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            'Сбросить полосы прокрутки
    hsbSant.Value = 0
    hsbLat.Value = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
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
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Время исключения Посетителя
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Посетителя
Dim strHour As String
Dim strMinute As String
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
            'Очистить текстовое поле "ДеньгиДата"
            txtMoneyDate.Text = ""
            'Сделать невидимой метку текстового поля "txtMoneyDate"
            lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
            imgMoneyFree.Visible = False
            imgCalendar.Visible = False
            fraDayNight.Visible = False
            imgTime.Visible = False
            'Сделать невидимыми соответствующие "Иконки"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            'Сделать недоступными элементы управления формы "DataAccessOut"
            lblLat0.Enabled = False
            lblLat320.Enabled = False
            hsbLat.Enabled = False
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
            intAutoFindCode = frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
            strInfo, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
            If intAutoFindCode = 0 Then
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

            'Вычислить время исключения Посетителя
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
            'Дата исключения Посетителя
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
                strDate = Trim(gProtocol.strProtocDate)
            
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
            'Сбросить признак  изменений в текстовом поле "PersonCode"
                    txtPersonCode.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
                     cmdOK.MousePointer = vbNoDrop
                     GoTo PersonCodeError
                End If
            'Бесплатный Посетитель
                If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
                    imgMoneyFree.Visible = True
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
                    gAccessMoneyCell = gAccessMoneyCellHuman
            'Формирование шаблона в поле "ДеньгиДаты"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Вся необходимая информация имеется
                    If txtPersonCode.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK_-"
                        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                        cmdOK.SetFocus
                    End If
            'Постоянный Посетитель
                ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
                    imgCalendar.Visible = True
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
                    gAccessMoneyCell = gAccessMoneyCellHuman
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
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Последний оплаченный день посещения еще не наступил
                    If CInt(Mid(txtMoneyDate.Text, 17, 4)) = CInt(Mid(txtMoneyDate.Text, 42, 4)) And _
                    ((CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
                    (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
                    CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
                    CInt(Mid(txtMoneyDate.Text, 36, 2)))) Or _
                    CInt(Mid(txtMoneyDate.Text, 17, 4)) > CInt(Mid(txtMoneyDate.Text, 42, 4)) Then
            
            'Вся необходимая информация имеется
                        txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
                        txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопку "OK_-"
                        cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                        cmdOK.SetFocus
                        Exit Sub
                    End If
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
            'Сделать доступными элементы управления формы "DataAccessOut"
                    lblLat0.Enabled = True
                    lblLat320.Enabled = True
                    hsbLat.Enabled = True
            'Временный Посетитель
                ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Тариф Предприятия (переменная для рассчетов) = Средняя
            '  стоимость полных суток (по тарифу одного часа)
                    intAccessTariff = (gAccessHourD + gAccessHourN) / 2 * 24
            'Сделать видимой метку текстового поля "txtMoneyDate"
                    lblMoneyDate.Visible = True
            'Сделать видимой соответствующую "Иконку"
                    imgTime.Visible = True
            'Взрослый
                    If Mid(txtInfo.Text, 5, 1) = "1" Then
                        imgHuman.Visible = True
            'Входной тариф Предприятия для Взрослых (для временных Клиентов)
                        gAccessMoneyCell = gAccessMoneyCellHuman
            'Дети
                    ElseIf Mid(txtInfo.Text, 5, 1) = "2" Then
                        imgBaby.Visible = True
            'Входной тариф Предприятия для Детей (для временных Клиентов)
                        gAccessMoneyCell = gAccessMoneyCellBaby
            'Конвой
                    ElseIf Mid(txtInfo.Text, 5, 1) = "3" Then
                        imgConvoy.Visible = True
            'Входной тариф Предприятия для Конвоя (для временных Клиентов)
                        gAccessMoneyCell = gAccessMoneyCellConvoy
            'Семья
                    ElseIf Mid(txtInfo.Text, 5, 1) = "4" Then
                        imgFamily.Visible = True
            'Входной тариф Предприятия для Семьи (для временных Клиентов)
                        gAccessMoneyCell = gAccessMoneyCellFamily
                    End If
            'Сделать доступными элементы управления формы "DataAccessOut"
                    lblLat0.Enabled = True
                    lblLat320.Enabled = True
                    hsbLat.Enabled = True
            'Формирование шаблона в поле "ДеньгиДаты"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
                End If
    
            'Не Бесплатный Посетитель
                If Left(Trim(strStatus), 2) <> "10" Then
            'Дата Регистрации Посетителя (или последнего дня посещения)
                    strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            'Вычисление Даты  Регистрации Клиента Предприятия
            '  (или последнего дня действия пропуска)
                    intDayReg = Left(strDate, 2)
                    intMonthReg = Mid(strDate, 4, 2)
                    intYearReg = Right(strDate, 4)
            'Белый фон текстового поля
                    txtMoneyDate.BackColor = vbWhite
            'Временный Посетитель
                    If Left(Trim(strStatus), 2) = "09" Then
            'Имитировать событие "Scroll" - прокрутка для ползунка "Lat"
                        hsbLat.Value = hsbLat.Max
            'Восстановлениe строки "Контроль"
'                        strChecking = strCheckingSafe
                    End If
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

            'Процедура анализа "PersonCode" при АвтоУдалении
            '  Посетителя через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Строка "Информация"
Dim strInfo As String
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Время исключения Посетителя
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Посетителя
Dim strHour As String
Dim strMinute As String
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
            'Сделать недоступным текстовое поле ПЕРСОНАЛЬНОГО
            '  КОДА вызываемой формы "frmDataAccessOut"
    txtPersonCode.Enabled = False
            'Голубой фон текстового поля
    txtPersonCode.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Сделать невидимыми соответствующие "Иконки"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА
    intAutoFindCode = frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
    strInfo, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
    If intAutoFindCode = 0 Then
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

            'Вычислить время исключения Посетителя
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
            'Дата исключения Посетителя
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(gProtocol.strProtocDate)
            
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
            'Сбросить признак  изменений в текстовом поле "PersonCode"
            txtPersonCode.Tag = 0
            'Сделать недоступным нажатие на кнопку "OK_-"
            cmdOK.MousePointer = vbNoDrop
            GoTo PersonCodeError
        End If
            'Бесплатный Посетитель
        If Left(Trim(strStatus), 2) = "10" Then
            'Сделать видимой соответствующую "Иконку"
            imgMoneyFree.Visible = True
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Вся необходимая информация имеется
            If txtPersonCode.Tag = 1 Then
            'Сделать доступным нажатие на кнопку "OK_-"
                cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                cmdOK.SetFocus
            End If
            'Постоянный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            'Сделать видимой соответствующую "Иконку"
            imgCalendar.Visible = True
            'Установить и сделать видимым соответствующий
            '  Признак Посетителя
            If Right(Trim(strChecking), 1) = "D" Then
            'Дневной тариф допуска
                optDay.Value = True
            ElseIf Right(Trim(strChecking), 1) = "N" Then
            'Ночной тариф допуска к
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
                Exit Function
            End If
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Последний оплаченный день посещения еще не наступил
            If (CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
            (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
            CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
            CInt(Mid(txtMoneyDate.Text, 36, 2))) Then
            'Вся необходимая информация имеется
                txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
                txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопку "OK_-"
                cmdOK.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                cmdOK.SetFocus
                Exit Function
            End If
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
            'Сделать доступными элементы управления формы "DataAccessOut"
            lblLat0.Enabled = True
            lblLat320.Enabled = True
            hsbLat.Enabled = True
            'Временный Посетитель
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            'Тариф Предприятия (переменная для рассчетов) = Средняя
            '  стоимость полных суток (по тарифу одного часа)
            intAccessTariff = (gAccessHourD + gAccessHourN) / 2 * 24
            'Сделать видимой метку текстового поля "txtMoneyDate"
            lblMoneyDate.Visible = True
            'Сделать видимой соответствующую "Иконку"
            imgTime.Visible = True
            'Взрослый
            If Mid(txtInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            'Дети
            ElseIf Mid(txtInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            'Конвой
            ElseIf Mid(txtInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            'Семья
            ElseIf Mid(txtInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
            'Сделать доступными элементы управления формы "DataAccessOut"
            lblLat0.Enabled = True
            lblLat320.Enabled = True
            hsbLat.Enabled = True
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
        End If
    
            'Не Бесплатный Посетитель
        If Left(Trim(strStatus), 2) <> "10" Then
            'Дата Регистрации Посетителя (или последнего дня посещения)
            strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            'Вычисление Даты  Регистрации Клиента Предприятия
            '  (или последнего дня действия пропуска)
            intDayReg = Left(strDate, 2)
            intMonthReg = Mid(strDate, 4, 2)
            intYearReg = Right(strDate, 4)
            'Белый фон текстового поля
            txtMoneyDate.BackColor = vbWhite
            'Имитировать событие "Scroll" - прокрутка для ползунка "Lat"
            hsbLat.Value = hsbLat.Max
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
            'Сделать недоступным нажатие на кнопку "OK_-"
    cmdOK.MousePointer = vbNoDrop
            'Установить фокус на текстовом поле "PersonCode"
    If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus

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
            'Сделать видимыми кнопки разрешения выхода Специальных Клиентов
            cmdOutFree.Visible = True
            cmdOutConst.Visible = True
            'Пароль неверный
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Сделать недоступной опцию "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
            'Сделать невидимыми кнопки разрешения выхода Специальных Клиентов
            cmdOutFree.Visible = False
            cmdOutConst.Visible = False
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
            'Время Регистрации Посетителя
Dim strHourReg As String
Dim strMinuteReg As String
Dim lngTimeReg As Long
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Время исключения Посетителя
Dim intHour As Integer
Dim intMinute As Integer
Dim lngTimeDel As Long
            'Нормализованное (по две цифры) время исключения Посетителя
Dim strHour As String
Dim strMinute As String
            'Нормализованная информация (две цифры числа) из поля "ДеньгиДаты"
Dim strMoneyDate As String
            'Дневное Начальное Время допуска на Предприятие (в Минутах)
Dim lngAccessTimeD As Long
            'Ночное Начальное Время допуска на Предприятие (в Минутах)
Dim lngAccessTimeN As Long
            
            'Событие - сброс ползунков
    If hsbLat.Tag <> 0 And hsbLat.Value = 0 And _
    hsbSant.Tag <> 0 Then hsbSant.Value = 0
            'Запомнить текущее положение ползунков
    hsbLat.Tag = hsbLat.Value
    hsbSant.Tag = hsbSant.Value
            'Не введен персональный код - выход из процедуры
    If txtPersonCode.Tag = 0 Then Exit Sub
            
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
            'Вычислить время исключения Посетителя
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
            'Дата исключения Посетителя
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Trim(gProtocol.strProtocDate)
            'Формирование шаблона в поле "ДеньгиДаты"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            
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
            'Сбросить признак  внесенной информации
    txtMoneyDate.Tag = 0
           'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопку "OK_-"
    cmdOK.MousePointer = vbNoDrop
            'Не нулевое положение одного из ползунков полос прокрутки
    If hsbLat.Value > 0 Or hsbSant.Value > 0 Then
            'Установить признак  внесенной информации
        txtMoneyDate.Tag = 1
            'Вносимая сумма Доплаты в Сантимах
        lngAccessMoney = hsbLat.Value * 100 + hsbSant.Value
            'Количество Доплачиваемых дней посещения
        intAccessDay = Int(lngAccessMoney / intAccessTariff)
            'Восстановление ИСХОДНОГО сотояния "Календаря"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            'Цикл по Дням "Календаря" (от последнего
            '  дня посещения или Даты Регистрации Клиента)
        For intAccessDay = intAccessDay To 1 Step -1
            'Количество Доплачиваемых дней посещения исчерпано
            If frmTableCalendar.comCalendar.Day = _
            Left(strDate, 2) And _
            frmTableCalendar.comCalendar.Month = _
            Mid(strDate, 4, 2) And _
            frmTableCalendar.comCalendar.Year = _
            Right(strDate, 4) Then GoTo EndCycle
            
            'Запись Числа, Месяца и Года в поле "ДеньгиДаты"
            If frmTableCalendar.comCalendar.Month > 9 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year)) + _
                Right(txtMoneyDate.Text, 31)
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + ".0" + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year)) + _
                Right(txtMoneyDate.Text, 31)
            End If
            'Продвижение "Календаря" на один день вперед
            frmTableCalendar.comCalendar.NextDay
            
        Next
    End If
EndCycle:
           'Недостаточная оплата
    If (frmTableCalendar.comCalendar.Day <> Left(strDate, 2) Or _
    frmTableCalendar.comCalendar.Month <> Mid(strDate, 4, 2) Or _
    frmTableCalendar.comCalendar.Year <> Right(strDate, 4)) And _
       (imgCalendar.Visible = True And imgTime.Visible = False Or _
        hsbLat.Value = 320) Or _
           (hsbLat.Value = 0 And hsbSant.Value = 0) Then
          'Отмена внесенной информации
        txtMoneyDate.Tag = 0
           'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
        cmdOK.MousePointer = vbNoDrop
        cmdOutFree.MousePointer = vbNoDrop
        cmdOutConst.MousePointer = vbNoDrop
    End If
            'Переплата (возможна Доплата только до текущего дня)
    If intAccessDay > 0 And (hsbLat.Value <> 0 Or hsbSant.Value <> 0) Then
            'Количество корректных (без переплаты) дней посещения
        intAccessDay = Int(lngAccessMoney / intAccessTariff) - intAccessDay
           'Восстановление корректной (без переплаты) суммы Доплаты в Сантимах
        lngAccessMoney = intAccessDay * intAccessTariff
            
            'Временный Посетитель - необходима коррекция Суммы в поле "ДеньгиДаты"
        If imgTime.Visible = True Then
            'Время Исключения Посетителя (в Минутах)
            lngTimeDel = intHour * 60 + intMinute
            'Нормализованная информация (две цифры числа) из поля "ДеньгиДаты"
            If Mid(Trim(txtMoneyDate.Text), 12, 1) = "." Then
                strMoneyDate = Left(Trim(txtMoneyDate.Text), 10) + "0" + _
                Trim(Mid(Trim(txtMoneyDate.Text), 11))
            Else
                strMoneyDate = Trim(txtMoneyDate.Text)
            End If
            'Время Регистрации Посетителя (в Минутах)
            strHourReg = Mid(Trim(strMoneyDate), 22, 2)
            strMinuteReg = Mid(Trim(strMoneyDate), 25, 2)
            lngTimeReg = CInt(strHourReg) * 60 + CInt(strMinuteReg)
           
           'Коррекция суммы Доплаты в Сантимах
            
            
            'Посещение менее "?"-и Минут - Входной тариф = 0
            '   (Уплачено при входе и текущая дата или предыдущая дата)
            If (lngTimeDel - lngTimeReg) <= 0 And _
            intAccessDay = 0 And (Left(strDate, 2) = Mid(Trim(strMoneyDate), 11, 2) Or _
            CInt(Left(strDate, 2)) < CInt(Mid(Trim(strMoneyDate), 11, 2)) Or _
            CInt(Left(strDate, 2)) > CInt(Mid(Trim(strMoneyDate), 11, 2)) And _
            (CInt(Mid(strDate, 4, 2)) < CInt(Mid(Trim(strMoneyDate), 14, 2)) Or _
            CInt(Mid(strDate, 7, 4)) < CInt(Mid(Trim(strMoneyDate), 17, 4)))) Then
                lngAccessMoney = 0
            'Посещение более "?"-и Минут - Входной тариф = Штраф + . . .
            Else
            'Дневное Начальное Время допуска на Предприятие (в Минутах)
                lngAccessTimeD = CInt(Left(Trim(gAccessTimeD), 2)) * 60 + _
                CInt(Mid(Trim(gAccessTimeD), 4, 2))
            'Ночное Начальное Время допуска на Предприятие (в Минутах)
                lngAccessTimeN = CInt(Mid(Trim(gAccessTimeD), 7, 2)) * 60 + _
                CInt(Right(Trim(gAccessTimeD), 2))
                lngAccessMoney = lngAccessMoney - intAccessTariff
            'Посетитель вошел во время текущих суток (после 00.00 часов)
                If Mid(strMoneyDate, 11, 2) = Mid(strMoneyDate, 36, 2) And _
                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2) Then
            'Cуммa оплаты входа (Штраф за опоздание при выходе)
                    lngAccessMoney = gAccessMoneyCell
            
'СПЕЦИАЛЬНО ДЛЯ ВЕНТСПИЛСА "ICE HALL"
                        
'Интервал визита - во время Дневного интервала допуска
If lngTimeReg >= lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
    lngAccessMoney = lngAccessMoney + _
    Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * gAccessMoneyCell

'''            'Интервал визита - во время Дневного интервала допуска
'''                    If lngTimeReg >= lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
'''                        lngAccessMoney = lngAccessMoney + _
'''                        Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * _
'''                        gAccessTimeCell * gAccessHourD / 60
            'Интервал визита - во время Ночного интервала допуска
                    ElseIf lngTimeReg > lngAccessTimeN And lngTimeDel <= 24 * 60 Or _
                    lngTimeReg >= 0 And lngTimeDel < lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            'Интервал визита - Смешанный (частично во время Дневного, а частично
            '   во время Ночного интервалов допуска
                    ElseIf lngTimeReg >= 0 And lngTimeDel <= 24 * 60 Then
            'Интервал визита - Ночь/День
                        If lngTimeReg < lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngTimeDel - lngAccessTimeD) / gAccessTimeCell) _
                            * gAccessTimeCell * gAccessHourD / 60 + _
                            Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60
            'Интервал визита - Ночь/День/Ночь
                        ElseIf lngTimeReg < lngAccessTimeD And lngTimeDel > lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60 + _
                            Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourD / 60 + _
                            Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60
            'Интервал визита - День/Ночь
                        ElseIf lngTimeReg >= lngAccessTimeD And lngTimeDel > lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60 + _
                            Int((lngAccessTimeN - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourD / 60
                        End If
                    End If
            
            
            'Посетитель вошел в предыдущие сутки (до 00.00 часов)
                ElseIf (CInt(Mid(strMoneyDate, 36, 2)) - CInt(Mid(strMoneyDate, 11, 2)) = 1 And _
                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2)) Or _
                (Mid(strMoneyDate, 36, 2) = "01" And _
                CInt(Mid(strMoneyDate, 39, 2)) - CInt(Mid(strMoneyDate, 14, 2)) = 1) Then
            'Интервал визита в предыдущие Сутки - Ночь
                    If lngTimeReg >= lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((24 * 60 - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            'Интервал визита в предыдущие Сутки - День/Ночь
                    ElseIf lngTimeReg >= lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngAccessTimeN - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int((24 * 60 - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            'Интервал визита в предыдущие Сутки - Ночь/День/Ночь
                    ElseIf lngTimeReg < lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60 + _
                        Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int((24 * 60 - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
                    End If
            'Интервал визита в текущие Сутки - Ночь
                    If lngTimeDel <= lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int(lngTimeDel / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            'Интервал визита в текущие Сутки - Ночь/День
                    ElseIf lngTimeDel <= lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int(lngAccessTimeD / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            'Интервал визита в текущие Сутки - Ночь/День/Ночь
                    ElseIf lngTimeDel > lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60 + _
                        Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int(lngAccessTimeD / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
                    End If
                End If
                
            End If
                
        End If
        
            'Восстановление корректного положения ползунков
        hsbSant.Value = lngAccessMoney - Int(lngAccessMoney / 100) * 100
        hsbLat.Value = Int(lngAccessMoney / 100)
        hsbLat_Change
    End If
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Or _
    txtPersonCode.Tag = 1 And _
    Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) = 0 And _
    intAccessDay = 0 Then
            'Установить признак  внесенной информации
        txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
        txtMoneyDate.BackColor = vbCyan
            'Сделать недоступной полосу прокрутки
        hsbLat.Enabled = False
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
        cmdOK.MousePointer = 0
        cmdOutFree.MousePointer = 0
        cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
        cmdOK.SetFocus
    End If
    
End Sub

            'Продлить время и дату удаления для Временного Клиента или время и дату выхода
            '  для Постоянного Клиента (на Предприятиях с ограничением времени непрерывного
            '  пребывания), до которых ему будет разрешен Выход
Private Sub Prolong(ByRef strStatus As String)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Дата (и Время) удаления Клиента
Dim strDate As String
            'Время удаления Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время удаления Клиента
Dim strHour As String
Dim strMinute As String
            'Признак присутствия \ 0 - вошел \ 1 - вышелл \ 2 - зарегистрирован
Dim strPresent As String * 1
            'Признак ("Е" - Окончательно вышел; "D" - Дневной тариф допуска;
            '  "N" - Ночной тариф допуска; "Другой символ"   - Суточный тариф
            '  допуска)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоКоррекции в "Таблице персон"
Dim intAutoCorrectionCode  As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  Постоянному Клиенту непрерывно находиться на Предприятии
Dim intCellLimit As Integer

            'Вычислить время удаления Клиента
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
            'Дата удаления Клиента
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
            'Признак регистрации Клиента
    strPresent = "2"
            'Признак Клиента
    strExpander = "P"

            'Если это Постоянный Клиент превысивший лимит времени непрерывного пребывания
            '  на Предприятии
    If gAcceTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
            'Количество ячеек времени, в течение которого разрешается
            '  Постоянному Клиенту непрерывно находиться на Предприятии
'        intCellLimit = gAccessCellLimit
    Else
        intCellLimit = 0
    End If
    
            'Вычислить "сдвинутые" время и дату Постоянного
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
    
            'Если это Постоянный Клиент превысивший лимит времени непрерывного пребывания
            '  на Предприятии
    If gAcceTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
        strCheckingInfo = ""
            'Формирование упакованной подстроки "Контроль"
        For intCount = 1 To 7 Step 2
            'Дата
            strCheckingInfo = Trim(strCheckingInfo) + _
            Chr(CByte(CInt(Mid(strDate, intCount, 2))))
        Next
            'Часы
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strHour, 1, 2))))
            'Минуты
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            'Упаковка подстроки "Контроль"
        Call frmTablePerson.Pack(strCheckingInfo)
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Extra Payment ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Papildus apmaksa ?", intButtonsAndIcons, "Cancel")
        End If
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Коррекция поля "txtInfo"
            txtInfo = Left(strCheckingInfo, 6) + Trim(Mid(txtInfo, 7))
        Else
            txtMoneyDate.Text = "000,00 Ls" + Mid(Trim(txtMoneyDate.Text), 10)
            Exit Sub
        End If
            
            'Это Временный Клиент и предоплата при входе не равна нулю
    ElseIf imgTime.Visible = True And imgCalendar.Visible = False And gAcceInpCellNumb > 0 Then
        strChecking = ""
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
        strChecking = Left(strChecking, 6) + strPresent + strExpander
    
    End If


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
        gProtocol.strProtocReserve = "AutoCorAcce"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
    End If

End Sub
