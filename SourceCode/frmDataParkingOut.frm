VERSION 5.00
Begin VB.Form frmDataParkingOut 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ParkingOutData"
   ClientHeight    =   3960
   ClientLeft      =   3195
   ClientTop       =   2565
   ClientWidth     =   8745
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
   ScaleWidth      =   8745
   Tag             =   "0"
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
      TabIndex        =   17
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   19
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingOut.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingOut.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingOut.frx":24A4
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
      Begin VB.Line Line8 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
   End
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   6960
      Max             =   99
      TabIndex        =   11
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      LargeChange     =   320
      Left            =   4920
      Max             =   320
      SmallChange     =   320
      TabIndex        =   10
      Top             =   2280
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
      TabIndex        =   9
      Top             =   3240
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
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   7
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   840
      Width           =   972
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   6480
      Top             =   120
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4200
      TabIndex        =   4
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   3120
      Width           =   4215
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
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5280
      X2              =   6240
      Y1              =   120
      Y2              =   120
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   2280
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
      TabIndex        =   12
      Top             =   2280
      Width           =   135
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   8520
      X2              =   8520
      Y1              =   1560
      Y2              =   3720
   End
   Begin VB.Image imgParkingIn 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataParkingOut.frx":28FE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   6240
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   1440
      Y2              =   600
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
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      Picture         =   "frmDataParkingOut.frx":2B10
      Stretch         =   -1  'True
      Top             =   240
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
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   8520
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   7080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   480
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
      TabIndex        =   3
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
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDataParkingOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Строка "Контроль" для Автостоянок
Dim strChecking As String * 8
            'Подстрока "Контроль" поля "txtInfo"
Dim strCheckingInfo As String * 8
            'Вносимая сумма оплаты в Сантимах
Dim lngParkingMoney As Long
            'Количество парковочных дней
Dim intParkingDay As Integer
            'Тариф одного парковочного дня (Сутки)
Dim intParkingTariffFull As Integer
            'Тариф одного парковочного дня (День)
Dim intParkingTariffDay As Integer
            'Тариф одного парковочного дня (Ночь)
Dim intParkingTariffNight As Integer
            'Тариф одного парковочного часа (переменная для рассчетов)
Dim intParkingTariffHour As Integer
            'Тариф Автостоянки (переменная для рассчетов)
Dim intParkingTariff As Integer
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
            'Строка "Таблицы календаря" соответствующая Дате
            '  последнего парковочного дня
Dim intRowNumReg As Integer
            'Столбец "Таблицы календаря", соответствующий Дате
            '  последнего парковочного дня
Dim intColNumReg As Integer

            'Перехват нажатия комбинаций клавиш "Alt"+ {"--", "E" , "L" и "S"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Форма "frmDataParkingOut" доступна
    If frmDataParkingOut.Enabled = True Then
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
            'Сделать доступными кнопки "OK" и "Cancel"
    cmdOK.MousePointer = 0
    cmdCancel.MousePointer = 0
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
            'Установить фокус на кнопкe "ОК"
    If cmdOK.Enabled = True Then cmdOK.SetFocus

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
            'Сделать доступными кнопки "OK" и "Cancel"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdOK.MousePointer = 0
    cmdCancel.MousePointer = 0
            'Установить фокус на кнопкe "ОК"
    If cmdOK.Enabled = True Then cmdOK.SetFocus

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
            
            'Продлить время и дату удаления для Временного Клиента или время и дату выезда
            '  для Постоянного Клиента (на Автостоянках с ограничением времени непрерывного
            '  пребывания АМ на Автостоянке), до которых ему будет разрешен Выезд
    If imgTime.Visible = True Then Call Prolong(strStatus)
            
            'Если это Постоянный Клиент превысивший лимит времени непрерывного пребывания
            '  АМ на Автостоянке
    If gParkTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
            ' Протоколирование события
            
            'ИНФОРМАЦИЯ
        gProtocol.strProtocName = txtInfo.Text
            'ПЕРСОНАЛЬНЫЙ КОД
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            'Статус
        gProtocol.strProtocStatus = strStatus
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
        gProtocol.strProtocReserve = "Extra Paym. " + Left(txtMoneyDate.Text, 9)
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingOut.Tag = 2
            
            'Возврат в вызвавшую процедуру
        cmdCancel_Click
        Exit Sub
    End If
    
            'Вызов процедуры-функции АвтоУдаления
            'ПЕРСОНАЛЬНОГО КОДА для Автостоянки
    intAutoDeletionCode = frmTablePerson.AutoDelParking(txtPersonCode.Text, strStatus)
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
        If gParkingDeletion = 1 Then
            'ФИЗИЧЕСКОЕ Удаление
            gProtocol.strProtocReserve = "AutoDelPark " + _
            Left(Trim(txtMoneyDate.Text), 9)
        Else
            'ЛОГИЧЕСКОЕ Удаление
            gProtocol.strProtocReserve = "LogDelPark " + _
            Left(Trim(txtMoneyDate.Text), 9)
        End If
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Изменения в текстовых полях текущей формы
            '   сохранены в "Таблице персон"
        txtPersonCode.Tag = 0
        txtMoneyDate.Tag = 0
            'Признак (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingOut.Tag = 1
        
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
            '   выполнение процедур невозможно), Временный Клиент Автостоянки,
            '   установлен признак НЕМЕДЛЕННОЕ Удаление и установлен индекс
            '   выездного терминала - открыть терминал
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        imgTime.Visible = True And gParkingDeletion = 1 And gTermOut <> -1 Then
            ' Электронная "Кнопка" не хранит адрес "Controller'a",
            '  выполнено Исключение Клиента Автостоянки
            '  и установлена Опция "Физическое удаление"
            If frmDemo.cmdOpen(gTermOut).Tag = 0 And frmDataParkingOut.Tag = 1 And _
            gParkingDeletion = 1 Then
            'Сделать недоступными элементы управления (Регистрация
            '  и Исключение Клиентов, Информация) для оператора Автостоянки
                frmDemo.imgParkingInData(gTermOut).Enabled = False
                frmDemo.imgParkingOutData(gTermOut).Enabled = False
                frmDemo.imgParkingInfoData(gTermOut).Enabled = False
            'Записать в электронную "Кнопку" адрес "Controller'a",
            '  требующего ручного подтверждения открытия терминала
                vntAddr = CByte(CInt(Trim(gParkAddrTerm(gTermOut))))
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
        gProtocol.strProtocReserve = "Invalid AutoDelParking"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
            'Белый фон текстового поля
        txtMoneyDate.BackColor = vbWhite
            'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
        frmDataParkingOut.Tag = 2
            
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
    If frmDataParkingOut.Tag = 1 And _
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
            'Выход из процедуры
            Exit Sub
        End If
    End If
    
                'Белый фон текстового поля
    txtMoneyDate.BackColor = vbWhite
        'Признак отказа от (Авто)Удаления ПЕРСОНАЛЬНОГО КОДА
    If frmDataParkingOut.Tag = 0 Then frmDataParkingOut.Tag = 2
            'Сделать невидимой текущую форму
    frmDataParkingOut.Visible = False
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
    
            'Сделать доступным элемент управления формы "DataParkingOut"
    txtPersonCode.Enabled = True
            'Сделать недоступными элементы управления формы "DataParkingOut"
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимыми кнопки разрешения выезда Специальных Клиентов
    cmdOutFree.Visible = False
    cmdOutConst.Visible = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми "Иконки"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
    cmdOK.MousePointer = vbNoDrop
    cmdOutFree.MousePointer = vbNoDrop
    cmdOutConst.MousePointer = vbNoDrop
             'Установить флаг завершения Активизации текущей формы
    frmDataParkingOut.Tag = 1

End Sub

            'Деактивизация текущей формы
Private Sub Form_Deactivate()
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    gTablePerson.AccessMinus

End Sub
            
            'Загрузка текущей формы
Private Sub Form_Load()
            'Сделать недоступными элементы управления формы "DataParkingOut"
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            'Сделать невидимой метку текстового поля "txtMoneyDate"
    lblMoneyDate.Visible = False
            'Сделать невидимыми элементы управления формы "DataParkingOut"
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
            'Тариф одного парковочного дня (Сутки)
    intParkingTariffFull = gParkingDN
            'Тариф одного парковочного дня (День)
    intParkingTariffDay = gParkingD
            'Тариф одного парковочного дня (Ночь)
    intParkingTariffNight = gParkingN
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
    cmdOK.MousePointer = vbNoDrop
    cmdOutFree.MousePointer = vbNoDrop
    cmdOutConst.MousePointer = vbNoDrop

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Процедура обработки "Щелчка мыши" на поле Персонального кода
Private Sub txtPersonCode_Click()
            
            'Белый фон текстового поля
    txtPersonCode.BackColor = vbWhite
            'Очистить текстовое поле "Информация" для Автостоянок
    txtInfo.Text = ""
            'Очистить текстовое поле "ДеньгиДата" для Автостоянок
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
            'Сделать недоступными элементы управления формы "DataParkingOut"
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
    cmdOK.MousePointer = vbNoDrop
    cmdOutFree.MousePointer = vbNoDrop
    cmdOutConst.MousePointer = vbNoDrop

End Sub

            'Процедура ввода и анализа "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            'Распакованная подстрока "Контроль" поля
            '  "Name" текущей строки "Таблицы персон"
Dim strCheckingUnPack As String
            'Переменная для Сохранения-Восстановления строки "Контроль"
Dim strCheckingSafe As String
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
            'Время исключения Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Клиента
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
            'Сделать недоступными элементы управления формы "DataParkingOut"
            hsbLat.Enabled = False
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА для Автостоянки
            intAutoFindCode = frmTablePerson.AutoFindParking(txtPersonCode.Text, _
            strInfo, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
            If intAutoFindCode = 0 Then
            'Заполнить текстовое поле "Информация"
                txtInfo.Text = strInfo
            
            'Сохранениe строки "Контроль"
                strCheckingSafe = strChecking
            
            'Распаковка строки "Контроль"
                Call frmTablePerson.UnPack(strDate, strChecking)
            
            'Отображение распакованной строки "Контроль"
                txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
                Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/въезда/выезда АМ Клиента
                If Mid(Trim(strChecking), 7, 1) = "0" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
                ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
                ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
                txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
                End If

            'Вычислить время исключения Клиента
            '  (или выезда Постоянного Клиента с Автостоянки с ограничением времени
            '  непрерывного пребывания АМ Постоянных Клиентов)
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
            'Дата исключения Клиента
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
                strDate = Trim(gProtocol.strProtocDate)
            
            'Анализ статуса Клиента Автостоянки
            
            'Недопустимый для Автостоянки статус Клиента
                If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
                Left(Trim(strStatus), 2) <> "06" Then
            'Окно собщения о неверном  статусе Клиента Автостоянки - на экран
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                    cmdOK.MousePointer = vbNoDrop
                    cmdOutFree.MousePointer = vbNoDrop
                    cmdOutConst.MousePointer = vbNoDrop
                    GoTo PersonCodeError
                End If
            'Бесплатный Клиент
                If Left(Trim(strStatus), 2) = "07" Then
            'Сделать видимой соответствующую "Иконку"
                    imgMoneyFree.Visible = True
            'Формирование шаблона в поле "ДеньгиДаты"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Вся необходимая информация имеется
                    If txtPersonCode.Tag = 1 Then
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                        cmdOK.MousePointer = 0
                        cmdOutFree.MousePointer = 0
                        cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                        cmdOK.SetFocus
                    End If
            'Постоянный Клиент
                ElseIf Left(Trim(strStatus), 2) = "05" Then
            'Сделать видимой соответствующую "Иконку"
                    imgCalendar.Visible = True
            'Установить и сделать видимым соответствующий
            '  Признак АМ Клиента
                    If Right(Trim(strChecking), 1) = "D" Then
            'Дневной тариф допуска к Автостоянке
                        optDay.Value = True
                    ElseIf Right(Trim(strChecking), 1) = "N" Then
            'Ночной тариф допуска к Автостоянке
                        optNight.Value = True
                    ElseIf Right(Trim(strChecking), 1) <> "D" And _
                    Right(Trim(strChecking), 1) <> "N" Then
            'Суточный тариф допуска к Автостоянке
                        optDayNight.Value = True
                    End If
            'Автомобиль не выехал Окончательно с Автостоянки
                    If Right(Trim(strChecking), 1) <> "E" Then
                        fraDayNight.Visible = True
            'Автомобиль выехал Окончательно с Автостоянки
                    Else
                        Exit Sub
                    End If
            'Формирование шаблона в поле "ДеньгиДаты"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Последний оплаченный день парковки еще не наступил
                    If CInt(Mid(txtMoneyDate.Text, 17, 4)) = CInt(Mid(txtMoneyDate.Text, 42, 4)) And _
                    ((CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
                    (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
                    CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
                    CInt(Mid(txtMoneyDate.Text, 36, 2)))) Or _
                    CInt(Mid(txtMoneyDate.Text, 17, 4)) > CInt(Mid(txtMoneyDate.Text, 42, 4)) Then
            
            'Автостоянка с ограничением времени непрерывного пребывания
                        If gParkTimeLimit > 0 Then
            'Распакованная подстрока "Контроль" поля "Информация"
                            Call frmTablePerson.UnPack(strCheckingUnPack, Left(txtInfo, 6) + "  ")
            'Формирование распакованной подстроки "Контроль"
                            strCheckingUnPack = Left(Trim(strCheckingUnPack), 2) + "." + _
                            Mid(Trim(strCheckingUnPack), 3, 2) + "." + _
                            Mid(Trim(strCheckingUnPack), 5, 4) + "/" + _
                            Mid(Trim(strCheckingUnPack), 9, 2) + ":" + _
                            Mid(Trim(strCheckingUnPack), 11, 2) + "/"
                
            'Последний разрешенный день непрерывного присутствия АМ Постоянного
            '  Клиента на Автостоянке уже наступил
                            If Not ((CInt(Mid(strCheckingUnPack, 4, 2)) > CInt(Mid(strDate, 4, 2))) Or _
                            (CInt(Mid(strCheckingUnPack, 4, 2)) = _
                            CInt(Mid(strDate, 4, 2)) And CInt(Left(strCheckingUnPack, 2)) >= _
                            CInt(Left(strDate, 2))) Or _
                            (CInt(Mid(strCheckingUnPack, 4, 2)) < _
                            CInt(Mid(strDate, 4, 2)) And CInt(Mid(strCheckingUnPack, 7, 4)) > _
                            CInt(Right(strDate, 4)))) Then
            'Замена строки "Контроль" подстрокой "Контроль"
                                strChecking = Left(txtInfo, 6) + Right(strChecking, 2)
            'Замена Даты последнего оплаченного дня на Дату последнего разрешенного
            '  дня непрерывного пребывания АМ Постоянного Клиента на Автостоянке
                                strDate = Left(strCheckingUnPack, 2) + "." + _
                                Mid(strCheckingUnPack, 4, 2) + "." + Mid(strCheckingUnPack, 7, 4)
            'Формирование шаблона в поле "ДеньгиДаты"
                                txtMoneyDate.Text = "000,00 Ls=" + Trim(strDate) + _
                                Mid(strCheckingUnPack, 11) + Mid(txtMoneyDate.Text, 28)
            'Сделать видимой ВТОРУЮ соответствующую "Иконку"
                                imgTime.Visible = True
                            ElseIf Not ((CInt(Mid(strCheckingUnPack, 12, 2)) > CInt(strHour)) Or _
                            (CInt(Mid(strCheckingUnPack, 12, 2)) = _
                            CInt(strHour) And CInt(Mid(strCheckingUnPack, 15, 2)) >= _
                            CInt(strMinute))) Then
            'Замена строки "Контроль" подстрокой "Контроль"
                                strChecking = Left(txtInfo, 6) + Right(strChecking, 2)
            'Замена Даты последнего оплаченного дня на Дату последнего разрешенного
            '  дня непрерывного пребывания АМ Постоянного Клиента на Автостоянке
                                strDate = Left(strCheckingUnPack, 2) + "." + _
                                Mid(strCheckingUnPack, 4, 2) + "." + Mid(strCheckingUnPack, 7, 4)
            'Формирование шаблона в поле "ДеньгиДаты"
                                txtMoneyDate.Text = "000,00 Ls=" + Trim(strDate) + _
                                Mid(strCheckingUnPack, 11) + Mid(txtMoneyDate.Text, 28)
            'Сделать видимой ВТОРУЮ соответствующую "Иконку"
                                imgTime.Visible = True
                            Else
            'Вся необходимая информация имеется
                                txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
                                txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                                cmdOK.MousePointer = 0
                                cmdOutFree.MousePointer = 0
                                cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                                cmdOK.SetFocus
                                Exit Sub
                            End If
            'Автостоянка без ограничения времени непрерывного пребывания
                        Else
            'Вся необходимая информация имеется
                            txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
                            txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                            cmdOK.MousePointer = 0
                            cmdOutFree.MousePointer = 0
                            cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                            cmdOK.SetFocus
                            Exit Sub
                        End If
                    End If
            
            'Лимит времени непрерывного пребывания АМ
            '  Постоянного Клиента на Автостоянке превышен
                    If imgTime.Visible = True Then
            'Тариф Автостоянки (переменная для рассчетов) = Средняя
            '  стоимость полных суток парковки (по тарифу одного
            '  парковочного часа)
                        intParkingTariff = (gParkingHourD + gParkingHourN) / 2 * 24
            'Сделать видимой метку текстового поля "txtMoneyDate"
                        lblMoneyDate.Visible = True
            'Сделать доступными элементы управления формы "DataParkingOut"
                        hsbLat.Enabled = True
            'Последний оплаченный день парковки уже наступил
                    Else
            'Определение тарифа Автостоянки (переменной для рассчетов)
            'Дневной тариф допуска к Автостоянке
                        If optDay.Value = True Then
                            intParkingTariff = intParkingTariffDay
            'Ночной тариф допуска к Автостоянке
                        ElseIf optNight.Value = True Then
                            intParkingTariff = intParkingTariffNight
            'Суточный тариф допуска к Автостоянке
                        ElseIf optDayNight.Value = True Then
                            intParkingTariff = intParkingTariffFull
                        End If
            'Сделать видимой метку текстового поля "txtMoneyDate"
                        lblMoneyDate.Visible = True
            'Сделать доступными элементы управления формы "DataParkingOut"
                        hsbLat.Enabled = True
                    End If
            'Временный Клиент
                ElseIf Left(Trim(strStatus), 2) = "06" Then
                
'СПЕЦИАЛЬНО ДЛЯ "SEL_2"
intParkingTariff = gParkingMoneyCell
                
            'Тариф Автостоянки (переменная для рассчетов) = Средняя
            '  стоимость полных суток парковки (по тарифу одного
            '  парковочного часа)
'''                    intParkingTariff = (gParkingHourD + gParkingHourN) / 2 * 24
            'Сделать видимой метку текстового поля "txtMoneyDate"
                    lblMoneyDate.Visible = True
            'Сделать видимой соответствующую "Иконку"
                    imgTime.Visible = True
            'Сделать доступными элементы управления формы "DataParkingOut"
                    hsbLat.Enabled = True
            'Формирование шаблона в поле "ДеньгиДаты"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
                End If
    
            'Не Бесплатный Клиент
                If Left(Trim(strStatus), 2) <> "07" Then
            'Дата Регистрации Клиента Автостоянки (или последнего парковочного дня)
                    strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            'Вычисление Даты  Регистрации Клиента Автостоянки
            '  (или последнего дня действия пропуска)
                    intDayReg = Left(strDate, 2)
                    intMonthReg = Mid(strDate, 4, 2)
                    intYearReg = Right(strDate, 4)
            'Белый фон текстового поля
                    txtMoneyDate.BackColor = vbWhite
            
'СПЕЦИАЛЬНО ДЛЯ "SEL_2"
'Временный Клиент
If Left(Trim(strStatus), 2) = "06" Then
'Имитировать событие "Scroll" - прокрутка для ползунка "Lat"
    hsbLat.Value = hsbLat.Max
End If

'''            'Имитировать событие "Scroll" - прокрутка для ползунка "Lat"
'''                    hsbLat.Value = hsbLat.Max
            'Восстановлениe строки "Контроль"
                    strChecking = strCheckingSafe
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
            cmdOK.MousePointer = vbNoDrop
            cmdOutFree.MousePointer = vbNoDrop
            cmdOutConst.MousePointer = vbNoDrop
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
            cmdOK.MousePointer = vbNoDrop
            cmdOutFree.MousePointer = vbNoDrop
            cmdOutConst.MousePointer = vbNoDrop
        End If
    End If

End Sub

            'Процедура анализа "PersonCode" при АвтоУдалении Клиента
            '  Автостоянки через специальный "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            'Строка "Информация" для Автостоянок
Dim strInfo As String
            'Статус
Dim strStatus As String
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Время исключения Клиента
Dim intHour As Integer
Dim intMinute As Integer
            'Нормализованное (по две цифры) время исключения Клиента
Dim strHour As String
Dim strMinute As String
            'Код возврата при АвтоПоиске в "Таблице персон"
Dim intAutoFindCode  As Integer
             'Ждать завершения Активизации текущей формы
    Do While frmDataParkingOut.Tag = 0
            'Обработать возможные события
        DoEvents
    Loop
            'Занести ПЕРСОНАЛЬНЫЙ КОД в соответствующее
            '  текстовое поле
    txtPersonCode.Text = Trim(vntPersonCode)
            'Сделать недоступным текстовое поле ПЕРСОНАЛЬНОГО
            '  КОДА вызываемой формы "frmDataParkingOut"
    txtPersonCode.Enabled = False
            'Голубой фон текстового поля
    txtPersonCode.BackColor = vbCyan
            'Установить признак  изменений в текстовом поле "PersonCode"
    txtPersonCode.Tag = 1
            'Вызов процедуры-функции АвтоПоиска
            'ПЕРСОНАЛЬНОГО КОДА для Автостоянки
    intAutoFindCode = frmTablePerson.AutoFindParking(txtPersonCode.Text, _
    strInfo, strStatus, strChecking)
            '(Авто)Поиск ПЕРСОНАЛЬНОГО КОДА выполнен успешно
            '   протоколирование события
    If intAutoFindCode = 0 Then
            'Заполнить текстовое поле "Информация" для Автостоянок
        txtInfo.Text = strInfo
            
            'Распаковка строки "Контроль"
        Call frmTablePerson.UnPack(strDate, strChecking)
            
            'Отображение распакованной строки "Контроль" для Автостоянок
        txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/въезда/выезда АМ Клиента
        If Mid(Trim(strChecking), 7, 1) = "0" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
        ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
        ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
        End If

            'Вычислить время исключения Клиента
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
            'Дата исключения Клиента
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(gProtocol.strProtocDate)
            
            'Анализ статуса Клиента Автостоянки
            
            'Недопустимый для Автостоянки статус Клиента
        If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
        Left(Trim(strStatus), 2) <> "06" Then
            'Окно собщения о неверном  статусе Клиента Автостоянки - на экран
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
            cmdOK.MousePointer = vbNoDrop
            cmdOutFree.MousePointer = vbNoDrop
            cmdOutConst.MousePointer = vbNoDrop
            GoTo PersonCodeError
        End If
            'Бесплатный Клиент
        If Left(Trim(strStatus), 2) = "07" Then
            'Сделать видимой соответствующую "Иконку"
            imgMoneyFree.Visible = True
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Вся необходимая информация имеется
            If txtPersonCode.Tag = 1 Then
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                cmdOK.MousePointer = 0
                cmdOutFree.MousePointer = 0
                cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                cmdOK.SetFocus
            End If
            'Постоянный Клиент
        ElseIf Left(Trim(strStatus), 2) = "05" Then
            'Сделать видимой соответствующую "Иконку"
            imgCalendar.Visible = True
            'Установить и сделать видимым соответствующий
            '  Признак АМ Клиента
            If Right(Trim(strChecking), 1) = "D" Then
            'Дневной тариф допуска к Автостоянке
                optDay.Value = True
            ElseIf Right(Trim(strChecking), 1) = "N" Then
            'Ночной тариф допуска к Автостоянке
                optNight.Value = True
            ElseIf Right(Trim(strChecking), 1) <> "D" And _
            Right(Trim(strChecking), 1) <> "N" Then
            'Суточный тариф допуска к Автостоянке
                optDayNight.Value = True
            End If
            'Автомобиль не выехал Окончательно с Автостоянки
            If Right(Trim(strChecking), 1) <> "E" Then
                fraDayNight.Visible = True
            'Автомобиль выехал Окончательно с Автостоянки
            Else
                Exit Function
            End If
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            'Последний оплаченный день парковки еще не наступил
            If (CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
            (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
            CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
            CInt(Mid(txtMoneyDate.Text, 36, 2))) Then
            'Вся необходимая информация имеется
                txtMoneyDate.Tag = 1
            'Голубой фон текстового поля
                txtMoneyDate.BackColor = vbCyan
            'Сделать доступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
                cmdOK.MousePointer = 0
                cmdOutFree.MousePointer = 0
                cmdOutConst.MousePointer = 0
            'Установить фокус на кнопке "ОК_+"
                cmdOK.SetFocus
                Exit Function
            End If
            'Определение тарифа Автостоянки (переменной для рассчетов)
            'Дневной тариф допуска к Автостоянке
            If optDay.Value = True Then
                intParkingTariff = intParkingTariffDay
            'Ночной тариф допуска к Автостоянке
            ElseIf optNight.Value = True Then
                intParkingTariff = intParkingTariffNight
            'Суточный тариф допуска к Автостоянке
            ElseIf optDayNight.Value = True Then
                intParkingTariff = intParkingTariffFull
            End If
            'Сделать видимой метку текстового поля "txtMoneyDate"
            lblMoneyDate.Visible = True
            'Сделать доступными элементы управления формы "DataParkingOut"
            hsbLat.Enabled = True
            'Временный Клиент
        ElseIf Left(Trim(strStatus), 2) = "06" Then
            'Тариф Автостоянки (переменная для рассчетов) = Средняя
            '  стоимость полных суток парковки (по тарифу одного
            '  парковочного часа)
            intParkingTariff = (gParkingHourD + gParkingHourN) / 2 * 24
            'Сделать видимой метку текстового поля "txtMoneyDate"
            lblMoneyDate.Visible = True
            'Сделать видимой соответствующую "Иконку"
            imgTime.Visible = True
            'Сделать доступными элементы управления формы "DataParkingOut"
            hsbLat.Enabled = True
            'Формирование шаблона в поле "ДеньгиДаты"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
        End If
    
            'Не Бесплатный Клиент
        If Left(Trim(strStatus), 2) <> "07" Then
            'Дата Регистрации Клиента Автостоянки (или последнего парковочного дня)
            strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            'Вычисление Даты  Регистрации Клиента Автостоянки
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
    cmdOK.MousePointer = vbNoDrop
    cmdOutFree.MousePointer = vbNoDrop
    cmdOutConst.MousePointer = vbNoDrop
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
            'Установить фокус на кнопкe "Cancel"
    If cmdCancel.Enabled = True Then cmdCancel.SetFocus
    
End Sub

            'Процедура обработки "Щелчка мыши" на поле пароля
Private Sub txtParole_Click()
            'Белый фон текстового поля
    txtParole.BackColor = vbWhite
            'Сделать невидимыми кнопки разрешения выезда Специальных Клиентов
    cmdOutFree.Visible = False
    cmdOutConst.Visible = False
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
            'Сделать видимыми кнопки разрешения выезда Специальных Клиентов
            cmdOutFree.Visible = True
            cmdOutConst.Visible = True
            'Сделать доступным нажатие на кнопки "0 Ls" и "XX San"
            cmdOutFree.MousePointer = 0
            cmdOutConst.MousePointer = 0
            'Пароль неверный
        Else
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Сделать недоступной опцию "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
            'Сделать невидимыми кнопки разрешения выезда Специальных Клиентов
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
            'Установить фокус на кнопкe "Cancel"
        If cmdCancel.Enabled = True Then cmdCancel.SetFocus
    End If

End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Lat"
Private Sub hsbLat_Change()
            'Время Регистрации Клиента
Dim strHourReg As String
Dim strMinuteReg As String
Dim lngTimeReg As Long
            'Дата и Время в ячейке "Reserve" " Таблицы персон"
Dim strDate As String
            'Время исключения Клиента
Dim intHour As Integer
Dim intMinute As Integer
Dim lngTimeDel As Long
            'Нормализованное (по две цифры) время исключения Клиента
Dim strHour As String
Dim strMinute As String
            'Нормализованная информация (две цифры числа) из поля "ДеньгиДаты"
Dim strMoneyDate As String
            'Дневное Начальное Время допуска на Автостоянку (в Минутах)
Dim lngParkingTimeD As Long
            'Ночное Начальное Время допуска на Автостоянку (в Минутах)
Dim lngParkingTimeN As Long
            
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
            
            'Отображение распакованной строки "Контроль" для Автостоянок
    txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            'Признак регистрации/въезда/выезда АМ Клиента
    If Mid(Trim(strChecking), 7, 1) = "0" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
    ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
    ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
    End If
            'Вычислить время исключения Клиента
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
            'Дата исключения Клиента
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
            'Сделать недоступным нажатие на кнопки "OK_-", "0 Ls" и "XX San"
    cmdOK.MousePointer = vbNoDrop
    cmdOutFree.MousePointer = vbNoDrop
    cmdOutConst.MousePointer = vbNoDrop
            'Не нулевое положение одного из ползунков полос прокрутки
    If hsbLat.Value > 0 Or hsbSant.Value > 0 Then
            'Установить признак  внесенной информации
        txtMoneyDate.Tag = 1
            'Вносимая сумма Доплаты в Сантимах
        lngParkingMoney = hsbLat.Value * 100 + hsbSant.Value
            'Количество Доплачиваемых парковочных дней
        intParkingDay = Int(lngParkingMoney / intParkingTariff)
            'Восстановление ИСХОДНОГО сотояния "Календаря"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            'Цикл по Дням "Календаря" (от последнего
            '  парковочного дня или Даты Регистрации Клиента)
        For intParkingDay = intParkingDay To 1 Step -1
            'Количество Доплачиваемых парковочных дней исчерпано
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
    If intParkingDay > 0 And (hsbLat.Value <> 0 Or hsbSant.Value <> 0) Then
            'Количество корректных (без переплаты) парковочных дней
        intParkingDay = Int(lngParkingMoney / intParkingTariff) - intParkingDay
           'Восстановление корректной (без переплаты) суммы Доплаты в Сантимах
        lngParkingMoney = intParkingDay * intParkingTariff
            
            'Временный Клиент - необходима коррекция Суммы в поле "ДеньгиДаты"
        If imgTime.Visible = True Then
            'Время Исключения Клиента (в Минутах)
            lngTimeDel = intHour * 60 + intMinute
            'Нормализованная информация (две цифры числа) из поля "ДеньгиДаты"
            If Mid(Trim(txtMoneyDate.Text), 12, 1) = "." Then
                strMoneyDate = Left(Trim(txtMoneyDate.Text), 10) + "0" + _
                Trim(Mid(Trim(txtMoneyDate.Text), 11))
            Else
                strMoneyDate = Trim(txtMoneyDate.Text)
            End If
            'Время Регистрации Клиента (в Минутах)
            strHourReg = Mid(Trim(strMoneyDate), 22, 2)
            strMinuteReg = Mid(Trim(strMoneyDate), 25, 2)
            lngTimeReg = CInt(strHourReg) * 60 + CInt(strMinuteReg)
           
           'Коррекция суммы Доплаты в Сантимах
            
            
            
'СПЕЦИАЛЬНО ДЛЯ "SEL_2"

            'АМ въехал на Автостоянку во время текущих суток
If Mid(strMoneyDate, 11, 2) = Mid(strMoneyDate, 36, 2) And _
Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2) Then
    lngParkingMoney = gParkingMoneyCell
            'АМ въехал на Автостоянку в предыдущие сутки или еще раньше
Else
    lngParkingMoney = gParkingMoneyCell + (intParkingDay - 1) * intParkingTariff + _
    Int((lngTimeDel + (24 * 60 - lngTimeReg)) / gParkingTimeCell) * gParkingMoneyCell
End If
'''            'Парковка менее "?"-и Минут - Въездной тариф
'''            If Int((lngTimeDel - lngTimeReg) / gParkingTimeCell) = 0 And intParkingDay = 0 Then
'''                lngParkingMoney = gParkingMoneyCell
'''            'Парковка более "?"-и Минут - Въездной тариф + . . .
'''            Else
'''            'Дневное Начальное Время допуска на Автостоянку (в Минутах)
'''                lngParkingTimeD = CInt(Left(Trim(gParkingTimeD), 2)) * 60 + _
'''                CInt(Mid(Trim(gParkingTimeD), 4, 2))
'''            'Ночное Начальное Время допуска на Автостоянку (в Минутах)
'''                lngParkingTimeN = CInt(Mid(Trim(gParkingTimeD), 7, 2)) * 60 + _
'''                CInt(Right(Trim(gParkingTimeD), 2))
'''                lngParkingMoney = lngParkingMoney - intParkingTariff + gParkingMoneyCell
'''            'АМ въехал на Автостоянку во время текущих суток (после 00.00 часов)
'''                If Mid(strMoneyDate, 11, 2) = Mid(strMoneyDate, 36, 2) And _
'''                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2) Then
'''            'Cуммa оплаты въезда
'''                lngParkingMoney = gParkingMoneyCell
'''            'Интервал парковки - во время Дневного интервала допуска к Aвтостоянке
'''                    If lngTimeReg >= lngParkingTimeD And lngTimeDel <= lngParkingTimeN Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngTimeDel - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60
'''            'Интервал парковки - во время Ночного интервала допуска к Aвтостоянке
'''                    ElseIf lngTimeReg > lngParkingTimeN And lngTimeDel <= 24 * 60 Or _
'''                    lngTimeReg >= 0 And lngTimeDel < lngParkingTimeD Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngTimeDel - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки - Смешанный (частично во время Дневного, а частично
'''            '   во время Ночного интервалов допуска к Aвтостоянке
'''                    ElseIf lngTimeReg >= 0 And lngTimeDel <= 24 * 60 Then
'''            'Интервал парковки - Ночь/День
'''                        If lngTimeReg < lngParkingTimeD And lngTimeDel <= lngParkingTimeN Then
'''                            lngParkingMoney = lngParkingMoney + _
'''                            Int((lngTimeDel - lngParkingTimeD) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                            Int((lngParkingTimeD - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки - Ночь/День/Ночь
'''                        ElseIf lngTimeReg < lngParkingTimeD And lngTimeDel > lngParkingTimeN Then
'''                            lngParkingMoney = lngParkingMoney + _
'''                            Int((lngParkingTimeD - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60 + _
'''                            Int((lngParkingTimeN - lngParkingTimeD) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                            Int((lngTimeDel - lngParkingTimeN) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки - День/Ночь
'''                        ElseIf lngTimeReg >= lngParkingTimeD And lngTimeDel > lngParkingTimeN Then
'''                            lngParkingMoney = lngParkingMoney + _
'''                            Int((lngTimeDel - lngParkingTimeN) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60 + _
'''                            Int((lngParkingTimeN - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60
'''                        End If
'''                    End If
'''
'''
'''            'АМ въехал на Автостоянку в предыдущие сутки (до 00.00 часов)
'''                ElseIf (CInt(Mid(strMoneyDate, 36, 2)) - CInt(Mid(strMoneyDate, 11, 2)) = 1 And _
'''                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2)) Or _
'''                (Mid(strMoneyDate, 36, 2) = "01" And _
'''                CInt(Mid(strMoneyDate, 39, 2)) - CInt(Mid(strMoneyDate, 14, 2)) = 1) Then
'''            'Интервал парковки в предыдущие Сутки - Ночь
'''                    If lngTimeReg >= lngParkingTimeN Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((24 * 60 - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки в предыдущие Сутки - День/Ночь
'''                    ElseIf lngTimeReg >= lngParkingTimeD Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngParkingTimeN - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                        Int((24 * 60 - lngParkingTimeN) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки в предыдущие Сутки - Ночь/День/Ночь
'''                    ElseIf lngTimeReg < lngParkingTimeD Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngParkingTimeD - lngTimeReg) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60 + _
'''                        Int((lngParkingTimeN - lngParkingTimeD) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                        Int((24 * 60 - lngParkingTimeN) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''                    End If
'''            'Интервал парковки в текущие Сутки - Ночь
'''                    If lngTimeDel <= lngParkingTimeD Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int(lngTimeDel / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки в текущие Сутки - Ночь/День
'''                    ElseIf lngTimeDel <= lngParkingTimeN Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngTimeDel - lngParkingTimeD) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                        Int(lngParkingTimeD / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''            'Интервал парковки в текущие Сутки - Ночь/День/Ночь
'''                    ElseIf lngTimeDel > lngParkingTimeN Then
'''                        lngParkingMoney = lngParkingMoney + _
'''                        Int((lngTimeDel - lngParkingTimeN) / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60 + _
'''                        Int((lngParkingTimeN - lngParkingTimeD) / gParkingTimeCell) * gParkingTimeCell * gParkingHourD / 60 + _
'''                        Int(lngParkingTimeD / gParkingTimeCell) * gParkingTimeCell * gParkingHourN / 60
'''                    End If
'''                End If
'''
'''            End If
                
        End If
        
            'Восстановление корректного положения ползунков
        hsbSant.Value = lngParkingMoney - Int(lngParkingMoney / 100) * 100
        hsbLat.Value = Int(lngParkingMoney / 100)
            
'СПЕЦИАЛЬНО ДЛЯ "SEL_2"
If Not (imgTime.Visible = True) Then hsbLat_Change
            
'''        hsbLat_Change
    End If
            'Вся необходимая информация имеется
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
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

            'Продлить время и дату удаления для Временного Клиента или время и дату выезда
            '  для Постоянного Клиента (на Автостоянках с ограничением времени непрерывного
            '  пребывания АМ на Автостоянке), до которых ему будет разрешен Выезд
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
            'Признак присутствия \ 0 - въехал \ 1 - выехал \ 2 - зарегистрирован
Dim strPresent As String * 1
            'Признак ("Е" - Окончательно выехал; "D" - Дневной тариф допуска;
            '  "N" - Ночной тариф допуска; "Другой символ"   - Суточный тариф
            '  допуска)
Dim strExpander As String * 1
            'Рабочий счетчик
Dim intCount As Integer
            'Код возврата при АвтоКоррекции в "Таблице персон"
Dim intAutoCorrectionCode  As Integer
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
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
            '  АМ на Автостоянке
    If gParkTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
            'Количество ячеек времени, в течение которого разрешается
            '  АМ Постоянного Клиента непрерывно находиться на Автостоянке
        intCellLimit = gParkingCellLimit
    Else
        intCellLimit = 0
    End If
            
            'Вычислить "сдвинутые" время и дату для Постоянного
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
    
            'Если это Постоянный Клиент превысивший лимит времени непрерывного пребывания
            '  АМ на Автостоянке
    If gParkTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
        
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
            'Это Временный Клиент
    ElseIf imgTime.Visible = True And imgCalendar.Visible = False Then
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
    intAutoCorrectionCode = frmTablePerson.AutoCorParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoCorPark"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
    End If

End Sub
