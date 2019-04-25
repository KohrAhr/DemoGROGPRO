VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmTableCalendar 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_calendar"
   ClientHeight    =   4800
   ClientLeft      =   2400
   ClientTop       =   3015
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6615
   Begin MSACAL.Calendar comCalendar 
      Height          =   495
      Left            =   1320
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
      _Version        =   524288
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   1
      Day             =   1
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   0   'False
      ShowDays        =   0   'False
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.Timer tmrMinute 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Tag             =   "0"
      Top             =   1560
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2520
      Max             =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Frame fraDayType 
      Caption         =   "Day Correction"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   240
      Width           =   4212
      Begin VB.OptionButton optWorkDay 
         Caption         =   " - Workday"
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
         Height          =   192
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton optSpecDay 
         Caption         =   "/^ - Specday"
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
         Height          =   192
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1332
      End
      Begin VB.OptionButton optHoliday 
         Caption         =   "/* - Holyday"
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
         Height          =   192
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.ListBox lstWeekNum 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableCalendar.frx":0000
      Left            =   120
      List            =   "frmTableCalendar.frx":0002
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame fraDay 
      Caption         =   "Day"
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
      Height          =   492
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   4212
      Begin VB.OptionButton optSun 
         Enabled         =   0   'False
         Height          =   192
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optSat 
         Enabled         =   0   'False
         Height          =   192
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optFri 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optThu 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optWed 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optTue 
         Enabled         =   0   'False
         Height          =   192
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optMon 
         Enabled         =   0   'False
         Height          =   192
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   255
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
      TabIndex        =   5
      Top             =   360
      Width           =   1092
   End
   Begin VB.CommandButton cmdNewCalen 
      Caption         =   "NewCalen"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   1092
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
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
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
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableCalendar 
      Height          =   1815
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   54
      Cols            =   8
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
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Calendars variant"
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
      TabIndex        =   24
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<="
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
      TabIndex        =   23
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblVariant99 
      Alignment       =   2  'Center
      Caption         =   "V0"
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
      Left            =   3720
      TabIndex        =   22
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblVariant0 
      Alignment       =   2  'Center
      Caption         =   "V0"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2520
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3840
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6480
      X2              =   2040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   4080
      Y2              =   1560
   End
   Begin VB.Label lblWeekNum 
      Caption         =   "Week Number "
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
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1332
   End
End
Attribute VB_Name = "frmTableCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '"Старый" номер варианта "Таблицы календаря"
Dim intVariantOld As Integer
            '"Новый" номер варианта "Таблицы календаря"
Dim intVariantNew As Integer
            'Текущий номер файла
Dim intFileNum As Integer
           'Текущий номер корректируемой строки "Таблицы календаря"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Таблицы календаря"
Dim intColNumCorr As Integer
            'Строка "Системной таблицы"
Dim gSystem As SystemInfo
            'Строка "Таблицы персон"
Dim gPerson As PersonInfo
            'Строка "Таблицы календаря"
Dim gCalendar As CalendarInfo

            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            '"Старый" номер варианта "Таблицы календаря" не нулевой
    If hsbVariant.Value <> 0 Then
            
            'Были не сохраненные изменения в "Таблице календаря"
        If gChangesTableCalendar = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы календаря" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение таблицы календаря в файле по умолчанию
                cmdSave_Click
            End If
        End If

    Else
            
            'Были не сохраненные изменения в "Таблице календаря"
        If gChangesTableCalendar = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы календаря" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение таблицы календаря в файле по умолчанию
                cmdSave_Click
            End If
        End If

            'Сделать недоступными элементы управления Коррекцией "Таблицы календаря"
        fraDayType.Enabled = False
        optHoliday.Enabled = False
        optSpecDay.Enabled = False
        optWorkDay.Enabled = False
        fraDay.Enabled = False
        optMon.Enabled = False
        optTue.Enabled = False
        optWed.Enabled = False
        optThu.Enabled = False
        optFri.Enabled = False
        optSat.Enabled = False
        optSun.Enabled = False
        lblWeekNum.Enabled = False
        lstWeekNum.Enabled = False
    End If
    
                'Очистить список номеров недель
    lstWeekNum.Clear
        
            'Сделать нулевым текущий вариант "Таблицы интервалов"
    intVariantOld = 0
    hsbVariant.Value = 0
            'Сбросить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = False

            'Сделать невидимой текущую форму
    frmTableCalendar.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub

            'Новый календарь (на следующий год)
Private Sub cmdNewCalen_Click()
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
Dim intColNum As Integer

            'Сделать недоступными элементы управления Коррекцией "Таблицы календаря"
    fraDayType.Enabled = False
    optHoliday.Enabled = False
    optSpecDay.Enabled = False
    optWorkDay.Enabled = False
    fraDay.Enabled = False
    optMon.Enabled = False
    optTue.Enabled = False
    optWed.Enabled = False
    optThu.Enabled = False
    optFri.Enabled = False
    optSat.Enabled = False
    optSun.Enabled = False
    lblWeekNum.Enabled = False
    lstWeekNum.Enabled = False
            
            '"Нулевой" вариант "Таблицы календаря"
    hsbVariant.Value = 0
            'Последняя строка Нулевого варианта "Таблицы календаря"
    grdTableCalendar.Row = grdTableCalendar.Rows - 1
            'Первый столбец последней строки
            ' Нулевого варианта "Таблицы календаря"
    grdTableCalendar.Col = 1
            'Дата формирования нового календаря
    comCalendar.Today
            'Дата "Первый день Первой недели в календаре Нового Года"
    If Left(Trim(grdTableCalendar.Text), 2) = "25" Then
        comCalendar.Value = Str(CInt(comCalendar.Year) + 1) + ".01.01."
    Else
        comCalendar.Value = Str(CInt(comCalendar.Year)) + ".12." + _
        Left(Trim(grdTableCalendar.Text), 2) + "."
    End If
            '"Первый" вариант "Таблицы календаря"
    hsbVariant.Value = 1
    
            'Цикл по всем нефиксированным строкам "Таблицы календаря" - (Даты)
    For intRowNum = 1 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
       grdTableCalendar.Row = intRowNum
            'По всем столбцам "Таблицы календаря"
       For intColNum = 1 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Запись даты в "Таблицу календарь"
            grdTableCalendar.Text = comCalendar.Day
            'Дата следующего дня
            comCalendar.NextDay

            'Установлен признак необходимости автоматического формирования
            '  опций выходных дней в "Таблице календаря" для Нового Года
            If gHolidays = 1 Then
            'Выходной день - Суббота или Воскресенье
                If intColNum >= 6 Then _
                grdTableCalendar.Text = grdTableCalendar.Text + "/*"
            End If
                
        Next
    Next
    
              'Установить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = True
            'Установить фокус на кнопке "Correction"
    If frmTableCalendar.Visible = True Then cmdCorrection.SetFocus

End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()

            ' "Таблица календаря" не содержит нефиксированных строк
    If grdTableCalendar.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        MsgBox ("The table is empty")
    
    Else
            'Сделать доступными элементы управления Коррекцией "Таблицы календаря"
        fraDayType.Enabled = True
        optHoliday.Enabled = True
            'Если нулевой вариант "Tаблицы календаря"
        If hsbVariant.Value = 0 Then optSpecDay.Enabled = True
        optWorkDay.Enabled = True
        optWorkDay.Value = True
        fraDay.Enabled = True
        optMon.Enabled = True
        optMon.Value = True
        optTue.Enabled = True
        optWed.Enabled = True
        optThu.Enabled = True
        optFri.Enabled = True
        optSat.Enabled = True
        optSun.Enabled = True
        lblWeekNum.Enabled = True
        lstWeekNum.Enabled = True
            'Очистить список имен
        lstWeekNum.Clear
    
            'Столбец "Week"
        grdTableCalendar.Col = 0
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
        For intRowNumCorr = 1 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
            grdTableCalendar.Row = intRowNumCorr
            'Заполнение списка "lstWeekNum" записями из "Таблицы календаря"
            lstWeekNum.AddItem grdTableCalendar.Text
        Next
            'Выбрать  элемент списка
        lstWeekNum.ListIndex = 0
            'Номер корректируемой строки - (Week - 1)
        intRowNumCorr = 1
        grdTableCalendar.Row = intRowNumCorr
            'Включить опции
        optMon.Value = True
        optWorkDay.Value = True
    
    End If
    
End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Таблицы календаря"
Private Sub grdTableCalendar_Click()
            'Коррекция "включена"
    If lstWeekNum.Enabled = True Then
            'Номер корректируемой строки "Таблицы календаря"
        intRowNumCorr = grdTableCalendar.RowSel
        grdTableCalendar.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstWeekNum.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Таблицы календаря"
        intColNumCorr = grdTableCalendar.ColSel
        grdTableCalendar.Col = intColNumCorr
            'Выбор корректируемой строки "Таблицы календаря"
        lstWeekNum_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstWeekNum.Left, _
        Y:=lstWeekNum.Top
            'Выбор корректируемого столбца "Таблицы календаря"
        Select Case intColNumCorr
            Case 1
            optMon.Value = True
            Case 2
            optTue.Value = True
            Case 3
            optWed.Value = True
            Case 4
            optThu.Value = True
            Case 5
            optFri.Value = True
            Case 6
            optSat.Value = True
            Case 7
            optSun.Value = True
        End Select
    End If

End Sub

            'Выбор корректируемой строки "Таблицы календаря"
Private Sub lstWeekNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемого столбца "Таблицы календаря" не равен нулю (Коррекция, а
            '  не заполнение списка номеров недель - "lstWeekNum")
        If intColNumCorr <> 0 Then
            'Номер корректируемой строки "Таблицы календаря"
            intRowNumCorr = lstWeekNum.ListIndex + 1
        End If
    End If

End Sub

            'Выбрана опция дня - "Выходной день"
Private Sub optHoliday_GotFocus()
            'Номер корректируемого столбца "Таблицы календаря"
    grdTableCalendar.Col = intColNumCorr
            'Номер корректируемой строки "Таблицы календаря"
    grdTableCalendar.Row = intRowNumCorr
            'Коррекция опции дня в "Таблице календаря"
    If InStr(1, Trim(grdTableCalendar.Text), "/") = 0 Then
        grdTableCalendar.Text = Trim(grdTableCalendar.Text) + "/*"
    Else
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1) + "/*"
    End If
    
            'Установить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = True

End Sub

            'Выбрана опция дня - "Специальный рабочий день"
Private Sub optSpecDay_GotFocus()
            'Номер корректируемого столбца "Таблицы календаря"
    grdTableCalendar.Col = intColNumCorr
            'Номер корректируемой строки "Таблицы календаря"
    grdTableCalendar.Row = intRowNumCorr
            'Коррекция опции дня в "Таблице календаря"
    If InStr(1, Trim(grdTableCalendar.Text), "/") = 0 Then
        grdTableCalendar.Text = Trim(grdTableCalendar.Text) + "/^"
    Else
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1) + "/^"
    End If
    
            'Установить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = True

End Sub

            'Выбрана опция дня - "Рабочий день"
Private Sub optWorkDay_GotFocus()
            'Номер корректируемого столбца "Таблицы календаря"
    grdTableCalendar.Col = intColNumCorr
            'Номер корректируемой строки "Таблицы календаря"
    grdTableCalendar.Row = intRowNumCorr
            'Коррекция опции дня в "Таблице календаря"
    If InStr(1, Trim(grdTableCalendar.Text), "/") <> 0 Then
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1)
    End If
    
              'Установить признак внесенных изменений в "Таблицу календаря"
  gChangesTableCalendar = True

End Sub

            'Выбрана опция "Понедельник"
Private Sub optMon_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 1
            
End Sub

            'Выбрана опция "Вторник"
Private Sub optTue_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 2

End Sub

            'Выбрана опция "Среда"
Private Sub optWed_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 3

End Sub

            'Выбрана опция "Четверг"
Private Sub optThu_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 4

End Sub

            'Выбрана опция "Пятница"
Private Sub optFri_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 5

End Sub

            'Выбрана опция "Суббота"
Private Sub optSat_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 6

End Sub

            'Выбрана опция "Воскресенье"
Private Sub optSun_Click()
            'Номер корректируемого столбца "Таблицы календаря"
    intColNumCorr = 7

End Sub

            'Обработка события "Change" - прокрутка для ползунка "Variant"
Private Sub hsbVariant_Change()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы календаря"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
Dim intColNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Установить ширину столбцов
    SetColWidth
            'Сделать недоступными элементы управления Коррекцией "Таблицы календаря"
    fraDayType.Enabled = False
    optHoliday.Enabled = False
    optSpecDay.Enabled = False
    optWorkDay.Enabled = False
    fraDay.Enabled = False
    optMon.Enabled = False
    optTue.Enabled = False
    optWed.Enabled = False
    optThu.Enabled = False
    optFri.Enabled = False
    optSat.Enabled = False
    optSun.Enabled = False
    lblWeekNum.Enabled = False
    lstWeekNum.Enabled = False
            'Очистить список номеров недель
    lstWeekNum.Clear
            
            'Были несохраненные изменения в "Таблице календаря"
    If gChangesTableCalendar = True Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы календаря" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранить "Новый" номер варианта "Таблицы календаря"
            intVariantNew = hsbVariant.Value
            '"Старый" номер варианта "Таблицы календаря"
            hsbVariant.Value = intVariantOld
            'Сохранение "Таблицы календаря" в файле по умолчанию
            cmdSave_Click
            'Восстановить "Новый" номер варианта "Таблицы календаря"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"Старый" номер варианта "Таблицы календаря"
    intVariantOld = hsbVariant.Value
            'Сбросить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = False

            'Заполнение варианта "Таблицы календаря" из файла
            
            'Вычислить длину записи (строки) "Таблицы календаря"
    lngRecordLen = Len(gCalendar)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            'Файл отсутствует - ?
    On Error GoTo ErrorTableCalendar
            'Дата (Текущая)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Открыть файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам варианта "Таблицы календаря"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        grdTableCalendar.Row = intRowNum
            'Читать строку "Таблицы календаря" из файла в буфер
        Get intFileNum, intRowNum + 1, gCalendar
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Заполнение текущей строки "Таблицы календаря" из буфера
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
        Next
    Next
            'Закрыть файл
    Close intFileNum
    
            'Установить фокус на кнопке "Correction"
    If frmTableCalendar.Visible = True Then cmdCorrection.SetFocus
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = hsbVariant.Value
    
    Exit Sub
ErrorTableCalendar:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")

End Sub
            
            'Сохранение "Таблицы календаря" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы персон"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
Dim intColNum As Integer
            'Вычислить длину записи (строки) "Таблицы календаря"
    lngRecordLen = Len(gCalendar)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем всем строкам "Таблицы календаря"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        grdTableCalendar.Row = intRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Заполнение текущей строки "Таблицы календаря" из буфера
            Select Case intColNum
                Case 0
                gCalendar.strWeekNum = Trim(grdTableCalendar.Text)
                Case 1
                gCalendar.strMonday = Trim(grdTableCalendar.Text)
                Case 2
                gCalendar.strTuesday = Trim(grdTableCalendar.Text)
                Case 3
                gCalendar.strWednesday = Trim(grdTableCalendar.Text)
                Case 4
                gCalendar.strThursday = Trim(grdTableCalendar.Text)
                Case 5
                gCalendar.strFriday = Trim(grdTableCalendar.Text)
                Case 6
                gCalendar.strSaturday = Trim(grdTableCalendar.Text)
                Case 7
                gCalendar.strSunday = Trim(grdTableCalendar.Text)
           End Select
        Next
            'Записать строку "Таблицы календаря" из буфера в файл
            Put intFileNum, intRowNum + 1, gCalendar
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Сбросить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = False
            'Установить фокус на кнопке "Cancel"
    If frmTableCalendar.Visible = True Then cmdCancel.SetFocus
            
End Sub
            'Сохранение таблицы календаря в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы персон"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
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
            'Вычислить длину записи (строки) "Таблицы календаря"
    lngRecordLen = Len(gCalendar)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Открыть выбираемый файл для произвольного доступа или
            '  создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем всем строкам "Таблицы календаря"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        grdTableCalendar.Row = intRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Заполнение текущей строки "Таблицы календаря" из буфера
            Select Case intColNum
                Case 0
                gCalendar.strWeekNum = Trim(grdTableCalendar.Text)
                Case 1
                gCalendar.strMonday = Trim(grdTableCalendar.Text)
                Case 2
                gCalendar.strTuesday = Trim(grdTableCalendar.Text)
                Case 3
                gCalendar.strWednesday = Trim(grdTableCalendar.Text)
                Case 4
                gCalendar.strThursday = Trim(grdTableCalendar.Text)
                Case 5
                gCalendar.strFriday = Trim(grdTableCalendar.Text)
                Case 6
                gCalendar.strSaturday = Trim(grdTableCalendar.Text)
                Case 7
                gCalendar.strSunday = Trim(grdTableCalendar.Text)
           End Select
        Next
            'Записать строку "Таблицы календаря" из буфера в файл
            Put intFileNum, intRowNum + 1, gCalendar
    Next
            'Закрыть  выбранный файл
    Close intFileNum
                'Сбросить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = False

    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Установить фокус на кнопке "Cancel"
    If frmTableCalendar.Visible = True Then cmdCancel.SetFocus
    
End Sub

            'Загрузка формы "Таблица календарь"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы календаря"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
Dim intColNum As Integer
            'Текущий месяц
Dim intMonth As Integer
            'Крайняя правая позиция даты в ячейке "Таблицы календаря" - (Далее следуют опции)
Dim intRightPositionDay

            'Установить ширину столбцов
    SetColWidth
            'Количество вариантов "Таблицы календаря"
    lblVariant99.Caption = "V" + Str(gVarNumCalendar)
    hsbVariant.Max = gVarNumCalendar
            'Переопределить размерность массива
ReDim gToday(gVarNumCalendar) As String * 4
            'Сохранить "Старый" номер варианта "Таблицы календаря"
    intVariantOld = hsbVariant.Value
            
            'Заполнение "Таблицы календаря" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы календаря"
    lngRecordLen = Len(gCalendar)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            'Файл отсутствует - ?
    On Error GoTo ErrorTableCalendar
            'Дата (Текущая)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Текущий месяц
    intMonth = 0
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы календаря"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        grdTableCalendar.Row = intRowNum
            'Читать строку "Таблицы календаря" из файла в буфер
        Get intFileNum, intRowNum + 1, gCalendar
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Заполнение текущей строки "Таблицы календаря" из буфера
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
            'Если нефиксированные строка и столбец "Таблицы календаря"
            If grdTableCalendar.Row <> 0 And grdTableCalendar.Col <> 0 Then
            'Вычисление крайней правой позиции даты в ячейке "Таблицы календаря"
                intRightPositionDay = InStr(1, Trim(grdTableCalendar.Text), "/")
                If intRightPositionDay = 0 Then
            'Опции дня отсутствуют в ячейке "Таблицы календаря" - (Нормальный рабочий день)
                    intRightPositionDay = 2
            'Опции дня присутствуют в ячейке "Таблицы календаря"
                Else
                    intRightPositionDay = intRightPositionDay - 1
                End If
            'Если 1-ое число, то номер месяца +1
                If CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) = 1 _
                Then intMonth = intMonth + 1
            'Если "загружается" текущий месяц
                If intMonth = Mid(gProtocol.strProtocDate, 4, 2) Then
            'Если даты совпали
                    If CInt(Left(gProtocol.strProtocDate, 2)) = _
                    CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) Then
            'Дата (Текущая) с опциями
                        gToday(0) = Trim(grdTableCalendar.Text)
            'Номер строки "Таблицы календаря", где расположена ячейка Текущего дня
                        gRowNum = intRowNum
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня
                        gColNum = intColNum
                    End If
                End If
            End If
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
                
            'Цикл по всем ненулевым вариантам "Таблицы календаря"
    For intVariantNew = 1 To gVarNumCalendar Step 1
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(intVariantNew)) + ".dat"
    
            'Открыть умалчиваемый файл для произвольного доступа
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Читать строку ненулевого варианта "Таблицы календаря" из файла в буфер
        Get intFileNum, gRowNum + 1, gCalendar
            'Текущая дата (с опциями) для ненулевого варианта  "Таблицы календаря"
        Select Case gColNum
            Case 1
            gToday(intVariantNew) = gCalendar.strMonday
            Case 2
            gToday(intVariantNew) = gCalendar.strTuesday
            Case 3
            gToday(intVariantNew) = gCalendar.strWednesday
            Case 4
            gToday(intVariantNew) = gCalendar.strThursday
            Case 5
            gToday(intVariantNew) = gCalendar.strFriday
            Case 6
            gToday(intVariantNew) = gCalendar.strSaturday
            Case 7
            gToday(intVariantNew) = gCalendar.strSunday
        End Select
            'Закрыть умалчиваемый файл
        Close intFileNum
    Next
                
                'Сбросить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = False
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableCalendar:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")
    
End Sub

            'Загрузка формы "Таблица календарь" в Новом Году
Private Sub NewYear()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы календаря"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы календаря"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы календаря"
Dim intColNum As Integer
            'Текущий месяц
Dim intMonth As Integer
            'Крайняя правая позиция даты в ячейке "Таблицы календаря" - (Далее следуют опции)
Dim intRightPositionDay

            'Установить ширину столбцов
    SetColWidth
            'Количество вариантов "Таблицы календаря"
    lblVariant99.Caption = "V" + Str(gVarNumCalendar)
    hsbVariant.Max = gVarNumCalendar
            'Переопределить размерность массива
ReDim gToday(gVarNumCalendar) As String * 4
            'Сохранить "Старый" номер варианта "Таблицы календаря"
    intVariantOld = hsbVariant.Value
            
            'Заполнение "Таблицы календаря" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы календаря"
    lngRecordLen = Len(gCalendar)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + "1" + ".dat"
    
    
            'Файл отсутствует - ?
    On Error GoTo ErrorTableCalendar
            'Дата (Текущая)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Текущий месяц
    intMonth = 0
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы календаря"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        grdTableCalendar.Row = intRowNum
            'Читать строку "Таблицы календаря" из файла в буфер
        Get intFileNum, intRowNum + 1, gCalendar
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            grdTableCalendar.Col = intColNum
            'Заполнение текущей строки "Таблицы календаря" из буфера
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
            'Если нефиксированные строка и столбец "Таблицы календаря"
            If grdTableCalendar.Row <> 0 And grdTableCalendar.Col <> 0 Then
            'Вычисление крайней правой позиции даты в ячейке "Таблицы календаря"
                intRightPositionDay = InStr(1, Trim(grdTableCalendar.Text), "/")
                If intRightPositionDay = 0 Then
            'Опции дня отсутствуют в ячейке "Таблицы календаря" - (Нормальный рабочий день)
                    intRightPositionDay = 2
            'Опции дня присутствуют в ячейке "Таблицы календаря"
                Else
                    intRightPositionDay = intRightPositionDay - 1
                End If
            'Если 1-ое число, то номер месяца +1
                If CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) = 1 _
                Then intMonth = intMonth + 1
            'Если "загружается" текущий месяц
                If intMonth = Mid(gProtocol.strProtocDate, 4, 2) Then
            'Если даты совпали
                    If CInt(Left(gProtocol.strProtocDate, 2)) = _
                    CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) Then
            'Дата (Текущая) с опциями
                        gToday(0) = Trim(grdTableCalendar.Text)
            'Номер строки "Таблицы календаря", где расположена ячейка Текущего дня
                        gRowNum = intRowNum
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня
                        gColNum = intColNum
                    End If
                End If
            End If
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Позиционировать движок полосы прокрутки на нулевую "Таблицу календарь"
    hsbVariant.Value = 0
                'Установить признак внесенных изменений в "Таблицу календаря"
    gChangesTableCalendar = True
            'Сохранить "Таблицу календарь" в умалчиваемом файле (Для Нового Года)
    Call cmdSave_Click
            
    
    Exit Sub
ErrorTableCalendar:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("New Year TableCalendar Error !")
    
End Sub

            'Процедура контроля времени (смена минут, часа, дня и года)
Private Sub tmrMinute_Timer()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Код возврата при вызове функции "Shell"
Dim vntShell As Variant
            'Длина строки "Системной таблицы", "Таблицы протокола" или "Таблицы календаря"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Системной таблицы"
Dim intRowNum As Integer
            'Текущий номер столбца "Системной таблицы"
Dim intColNum As Integer
            'Дата
Dim strDate As String
            'Время
Dim strTime As String
Dim intHour As Integer
Dim intMinute As Integer
Dim strHour As String
Dim strMinute As String
            'Рабочий счетчик
Dim intCount As Integer
            'Строка отсылаемого сообщения
Dim strMessage As String
            
            'Текущая дата
    strDate = Trim(Format(Now, "dd/mm/yyyy"))
            
            'Текущее время
    strTime = Format(Now, "h:mm:ss")
            'Часы
    intHour = Hour(strTime)
    If intHour < 10 Then
        strHour = "0" + Trim(Str(intHour))
    Else
        strHour = Trim(Str(intHour))
    End If
            'Минуты
    intMinute = Minute(strTime)
    If intMinute < 10 Then
        strMinute = "0" + Trim(Str(intMinute))
    Else
        strMinute = Trim(Str(intMinute))
    End If
    
            'Вывод Даты и Времени
    frmDemo.lblTime.Caption = "   " + strDate + "   " + strHour + ":" + strMinute

        
            'Если установлен признак необходимости сжатия "Таблицы персон":
            '   устанавливается всегда в "Host Computer'e" и в тех случаях
            '   в "Препроцессоре", когда последний использует свою
            '   собственную "Таблицу персон" - "ЗЕРКАЛЬНАЯ Таблицa персон"
    If gCompresTablPers = 1 Then
            'Если режим "Выполнение"
        If frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
            'Если установлен запрос на реальное удаление строк из
            '   "Таблицы персон", "Таблицa персон" доступна для
            '   вычеркивания строк и контроллер терминала свободен
            If gRealDelPerson = True And gTablePerson.Access < 1 And _
            frmDemo.lblInform(frmDemo.tmrTermContr.Tag).Tag = 0 Then
            'Физически удалить логически удаленные строки 'Таблицы персон"
                Call frmTablePerson.RealDelPerson
            End If
            'Если установлен признак внесенных изменений в
            '   "Таблицу персон" - сохранить таблицу в умалчиваемом файле
            If gChangesTablePerson = True Then
                Call frmTablePerson.SaveTablePerson
            End If
        End If
    End If
            
            'Если это не "Host Computer" и режим "Выполнение"
    If gPreprocName <> "" And frmDemo.Enabled = True And _
    frmDemo.chkSetup.Value = 1 Then
            
            'Если есть признак запрета формирования баз Протокола и
            '   Бухгалтерии - закрыть/открыть "Таблицу протокола"
        If gMSBase = 0 Then
            'Закрыть файл "Таблицы протокола"
            Close gProtocFileNum
            'Вычислить длину записи (строки) "Таблицы протокола"
            lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла "Таблицы протокола"
            gProtocFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            strPathFileName = strPathFileName + "TableProtocol.dat"
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
            Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            'Номер первой свободной строки "Таблицы протокола"
            gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    
    End If
    
    
            'Час истек и режим "Выполнение"
    If strHour <> frmDemo.lblTime.Tag And _
    frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
            'Запомнить новое время (Час) формирования баз
        frmDemo.lblTime.Tag = strHour
            
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Автостоянке или на Предприятии
        If gParkingPlaceNum <> 0 Or gAccessPlaceNum <> 0 Then
            
            If gParkingPlaceNum <> 0 Then
            'Имеется технологический перерыв в работе
            '  Автостоянки
                If Not (Left(gDefaultParkTime, 2) = "00" And _
                Mid(gDefaultParkTime, 4, 2) = "00" And _
                Mid(gDefaultParkTime, 7, 2) = "24" And _
                Mid(gDefaultParkTime, 10, 2) = "00") And _
                Left(gDefaultParkTime, 2) = intHour And _
                Mid(gDefaultParkTime, 4, 2) >= intMinute Then
            'Исходное количество свободных мест на Автостоянке
                    gParkFreePlaces = gParkingPlaceNum
                ElseIf Not (Left(gDefaultParkTime, 2) = "00" And _
                Mid(gDefaultParkTime, 4, 2) = "00" And _
                Mid(gDefaultParkTime, 7, 2) = "24" And _
                Mid(gDefaultParkTime, 10, 2) = "00") And _
                Mid(gDefaultParkTime, 7, 2) = intHour Then
            'Нулевое количество свободных мест на Автостоянке
                    gParkFreePlaces = 0
                End If
            'Подготовка инициализация дисплея-указателя
                strMessage = "ParkFreePlaces=" + CStr(gParkFreePlaces)
            End If
            
            If gAccessPlaceNum <> 0 Then
            'Имеется технологический перерыв в работе
            '  Предприятия
                If Not (Left(gDefaultAcceTime, 2) = "00" And _
                Mid(gDefaultAcceTime, 4, 2) = "00" And _
                Mid(gDefaultAcceTime, 7, 2) = "24" And _
                Mid(gDefaultAcceTime, 10, 2) = "00") And _
                Left(gDefaultAcceTime, 2) = intHour And _
                Mid(gDefaultAcceTime, 4, 2) >= intMinute Then
            'Исходное количество свободных мест на Предприятии
                    gAcceFreePlaces = gAccessPlaceNum
                ElseIf Not (Left(gDefaultAcceTime, 2) = "00" And _
                Mid(gDefaultAcceTime, 4, 2) = "00" And _
                Mid(gDefaultAcceTime, 7, 2) = "24" And _
                Mid(gDefaultAcceTime, 10, 2) = "00") And _
                Mid(gDefaultAcceTime, 7, 2) = intHour Then
            'Нулевое количество свободных мест на Предприятии
                    gAcceFreePlaces = 0
                End If
            'Подготовка инициализация дисплея-указателя
                strMessage = "AcceFreePlaces=" + CStr(gAcceFreePlaces)
            End If
            
            'Вывести информацию на дисплей
            Call frmDemo.Display(strMessage)
            'Сохранить новую 'Таблицу персон"
            Call frmTablePerson.SaveTablePerson
        End If
            
            'Установлен признак формирования баз Протокола и Бухгалтерии
        If gMSBase = 1 Then
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
            Call frmDemo.BasesConvert
        End If
    
    End If
            
            'Сутки истекли или необходима коррекция Текущего Года
            '  в "Системной таблице"
    If strDate <> frmTableCalendar.Tag Or _
    Right(Trim(frmTableCalendar.Tag), 4) <> Trim(Str(gYear)) Then
    
            'Если это не "Host Computer" и режим "Выполнение"
        If gPreprocName <> "" And frmDemo.Enabled = True And _
        frmDemo.chkSetup.Value = 1 Then
            'Формирование СООБЩЕНИЯ "Host Computer'у" о необходимости
            '  синхронизации времени для данного Препроцессора
            qMsgOutput.Body = "Time"
            ' Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
            qInfoOutput.FormatName = "DIRECT=OS:" + gHost + "\Private$\GeneralQueue"
            ' Открыть очередь сообщений с параметрами (для передачи
            '   сообщений, доступ к очереди разрешен всем)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' Отослать СООБЩЕНИЕ
            qMsgOutput.Send qQueueOutput
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Автостоянке
            If gParkingPlaceNum <> 0 Then
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                qMsgOutput.Body = "ParkFreePlaces "
            ' Отослать СООБЩЕНИЕ
                qMsgOutput.Send qQueueOutput
            End If
            ' Если имеется дисплей-указатель количества свободных мест
            '   на Предприятии
            If gAccessPlaceNum <> 0 Then
            ' Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
                qMsgOutput.Body = "AcceFreePlaces "
            ' Отослать СООБЩЕНИЕ
                qMsgOutput.Send qQueueOutput
            End If
            ' Закрыть очередь СООБЩЕНИЙ
            qQueueOutput.Close
        End If
    
            'Новый Год наступил или необходима коррекция Текущего Года
        If Right(Trim(strDate), 4) <> Right(Trim(frmTableCalendar.Tag), 4) Or _
        Right(Trim(strDate), 4) <> Trim(Str(gYear)) Then
            'Загрузить "Таблицу календарь" Нового Года
            Call NewYear
            'Запомнить новую дату
            frmTableCalendar.Tag = strDate
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
            For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка  "Системной таблицы"
                frmTableSystem.grdTableSystem.Row = intRowNum
            'Фиксированный столбец "Системной таблицы" (Объект)
                frmTableSystem.grdTableSystem.Col = 0
                If Trim(frmTableSystem.grdTableSystem.Text) = "gYear" Then
            'Текущий столбец "Системной таблицы"=1(Константа)
                    frmTableSystem.grdTableSystem.Col = 1
            'Установить Текущий год
                    gYear = Right(Trim(strDate), 4)
                    frmTableSystem.grdTableSystem.Text = gYear
                End If
            Next
            'Сохранить новую 'Системную таблицу"
            Call frmTableSystem.SaveTableSystem
            
            'Сформировать новую "Таблицу календарь" для следующего Нового Года
            Call cmdNewCalen_Click
            'Позиционировать движок полосы прокрутки на первую
            '  "Таблицу календарь" (Для Нового Года)
            hsbVariant.Value = 1
            'Сохранить "Таблицу календарь" в умалчиваемом файле (Для Нового Года)
            Call cmdSave_Click
            'Позиционировать движок полосы прокрутки на нулевую "Таблицу календарь"
            hsbVariant.Value = 0
        
        Else
            'Запомнить новую дату
            frmTableCalendar.Tag = strDate
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня = 7
            If gColNum = 7 Then
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня
                gColNum = 1
            'Номер строки "Таблицы календаря", где расположена ячейка Текущего дня
            gRowNum = gRowNum + 1
            Else
            'Номер столбца "Таблицы календаря", где расположена ячейка Текущего дня
                gColNum = gColNum + 1
            End If
        End If
            
            'Вычислить длину записи (строки) "Таблицы календаря"
        lngRecordLen = Len(gCalendar)
            
            'Цикл по всем вариантам "Таблицы календаря"
        For intVariantNew = 0 To gVarNumCalendar Step 1
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            strPathFileName = strPathFileName + "TableCalendar" + _
            Trim(Str(intVariantNew)) + ".dat"
    
            'Файл отсутствует - ?
            On Error GoTo ErrorTableCalendar
            'Открыть умалчиваемый файл для произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Читать строку варианта "Таблицы календаря" из файла в буфер
            Get intFileNum, gRowNum + 1, gCalendar
            'Текущая дата (с опциями) для варианта  "Таблицы календаря"
            Select Case gColNum
                Case 1
                gToday(intVariantNew) = gCalendar.strMonday
                Case 2
                gToday(intVariantNew) = gCalendar.strTuesday
                Case 3
                gToday(intVariantNew) = gCalendar.strWednesday
                Case 4
                gToday(intVariantNew) = gCalendar.strThursday
                Case 5
                gToday(intVariantNew) = gCalendar.strFriday
                Case 6
                gToday(intVariantNew) = gCalendar.strSaturday
                Case 7
                gToday(intVariantNew) = gCalendar.strSunday
            End Select
            'Закрыть умалчиваемый файл
            Close intFileNum
        Next
    End If
    
    Exit Sub
ErrorTableCalendar:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")

End Sub
            
            'Процедура установки ширины и выравнивания столбцов "Таблицы календаря"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Фиксированный столбец
    grdTableCalendar.ColWidth(intColNumber) = 1070
    grdTableCalendar.ColAlignment(intColNumber) = 0
            'Цикл по всем нефиксированным столбцам
    For intColNumber = 1 To grdTableCalendar.Cols - 1 Step 1
        grdTableCalendar.ColWidth(intColNumber) = 415
        grdTableCalendar.ColAlignment(intColNumber) = 0
    Next
    
End Sub


