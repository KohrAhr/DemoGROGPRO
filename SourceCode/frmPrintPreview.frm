VERSION 5.00
Begin VB.Form frmPrintPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "print_preview"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   11190
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
   ScaleHeight     =   7785
   ScaleWidth      =   11190
   Begin VB.CheckBox chkPrintPage 
      Caption         =   "Print Page"
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
      Left            =   1800
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtDateFrom 
      Height          =   372
      Left            =   6960
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
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
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CheckBox chkProtocol 
      Caption         =   "Connect Protocol Base"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   4560
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox txtReservOrNote 
      Height          =   372
      Left            =   8640
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.TextBox txtDateTo 
      Height          =   372
      Left            =   6960
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox txtCodeOrPassword 
      Height          =   372
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.TextBox txtName 
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Data datProtocol 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   516
      Left            =   5880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdFirst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   360
      Picture         =   "frmPrintPreview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   252
   End
   Begin VB.CommandButton cmdLast 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      Picture         =   "frmPrintPreview.frx":0112
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   252
   End
   Begin VB.CommandButton cmdPrintPage 
      Caption         =   "&PrintPage"
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
      Left            =   8160
      TabIndex        =   3
      Top             =   7080
      Width           =   1212
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Pre&vious"
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
      Left            =   720
      TabIndex        =   2
      Top             =   7080
      Width           =   1212
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
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
      TabIndex        =   1
      Top             =   7080
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   7080
      Width           =   1212
   End
   Begin VB.Label lblDateTo 
      Alignment       =   2  'Center
      Caption         =   "To"
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
      Left            =   6360
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.Label lblDateFrom 
      Alignment       =   2  'Center
      Caption         =   "From"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   492
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Текущий номер строки формы "frmPrintPreview"
Dim intRowPrintNum As Integer
            'Количество строк на одной странице формы "frmPrintPreview"
Dim intRowPrintQuan As Integer
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "Архив протокола")
Dim lngRowNum As Long
            'Текущий номер столбца таблиц ("TablePerson", "TableCalendar","TableProtocol,
            '  "TableSystem", "TableTime", "TableTerminal")
Dim intColNum As Integer
            'Начальная дата в "Таблице протокола"
Dim strDateFrom As String
            'Количество записей в "Архиве протокола"
Dim lngArchivesRowNum As Long
             'Номер файла "Архив протокола"
Dim intFileNum As Integer
           'Количество строк в "Базе Протокола"
Dim intProtocolBaseCount As Integer
           'Буфер печати строки "Системной таблицы"
Dim strTableSystem(5) As String
           'Буфер печати строки "Таблицы персон"
Dim strTablePerson(6) As String
            'Буфер печати строки "Таблицы календаря"
Dim strTableCalendar(8) As String
           'Буфер печати строки "Таблицы времени"
Dim strTableTime(3) As String
           'Буфер печати строки "Таблицы терминалов"
Dim strTableTerminal(4) As String
            
            'Изменение значения опции "База Протокола"
Private Sub chkProtocol_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
            'Текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
    
            'Очистить форму
    frmPrintPreview.Cls
            'Опция "База Протокола" выключена
    If chkProtocol.Value = 0 Then
            'Кнопка "Find" доступна для нажатия
        cmdFind.MousePointer = 0
            'Сделать доступной кнопку "First"
        cmdFirst.Enabled = True
            'Сделать доступной кнопку "Last"
        cmdLast.Enabled = True
            'Сделать доступной кнопку "Previous"
        cmdPrevious.Enabled = True
            'Сделать недоступной кнопку "Next"
        cmdNext.Enabled = False
            'Сделать доступной кнопку "PrintPage"
        cmdPrintPage.Enabled = True
            'Сделать доступной кнопку "Cancel"
        cmdCancel.Enabled = True
            'Сделать невидимой кнопку "Find"
        cmdFind.Visible = False
            'Сделать невидимой  опцию "PrintPage"
        chkPrintPage.Visible = False
            'Сделать невидимым поле ввода "Name"
        txtName.Visible = False
            'Сделать невидимым поле ввода "CodeOrPassword"
        txtCodeOrPassword.Visible = False
            'Сделать невидимыми поля ввода "Date"
        txtDateFrom.Visible = False
        txtDateTo.Visible = False
            'Сделать невидимыми метки полей ввода "Date"
        lblDateFrom.Visible = False
        lblDateTo.Visible = False
            'Сделать невидимым поле ввода "ReservOrNote"
        txtReservOrNote.Visible = False

            
            'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
            
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
                If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
        
            'Опция "База Протокола" включена
    Else
            
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            'Установка свойств элемента "Data" доступа к "Базе Протокола"
        datProtocol.DatabaseName = strPathFileName + "ProtocolBase.mdb"
        datProtocol.RecordSource = "Protocol"
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
        frmPrintPreview.MousePointer = vbHourglass
            'Кнопка "Find" доступна для нажатия
        cmdFind.MousePointer = 0
            'Сделать недоступной кнопку "First"
        cmdFirst.Enabled = False
            'Сделать недоступной кнопку "Last"
        cmdLast.Enabled = False
            'Сделать недоступной кнопку "Previous"
        cmdPrevious.Enabled = False
            'Сделать недоступной кнопку "Next"
        cmdNext.Enabled = False
            'Сделать недоступной кнопку "PrintPage"
        cmdPrintPage.Enabled = False
            'Сделать недоступной кнопку "Cancel"
        cmdCancel.Enabled = False
        
            'Определить количество записей в "Базе Протокола"
        datProtocol.Refresh
        datProtocol.Recordset.MoveLast
        intProtocolBaseCount = datProtocol.Recordset.RecordCount
            'Обновить "Базу Протокола"
        datProtocol.Recordset.MoveFirst
            'Цикл по всем строкам "Таблицы протокола"
        For lngRowNum = 1 To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, lngRowNum, gProtocol
            'Дата "From"
            If lngRowNum = 1 Then
                txtDateFrom.Text = gProtocol.strProtocDate
                strDateFrom = txtDateFrom.Text
            End If
            'Обновить текущую запись "Базы Протокола"
            datProtocol.Recordset.Edit
            datProtocol.Recordset.Fields("Name").Value = gProtocol.strProtocName
            datProtocol.Recordset.Fields("CodeOrPassword").Value = _
            gProtocol.strProtocPersonCode
            datProtocol.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
            datProtocol.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            datProtocol.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            datProtocol.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
            datProtocol.Recordset.Update
            'Не последняя запись старой "Базы Протокола"
            If lngRowNum < intProtocolBaseCount Then
                datProtocol.Recordset.MoveNext
            'Последняя запись старой "Базы Протокола"
            Else
                datProtocol.Recordset.AddNew
                datProtocol.Recordset.Update
                datProtocol.Recordset.MoveNext
            End If
        Next
            'Удаление одной лишней записи из  "Базы Протокола"
        If lngRowNum > intProtocolBaseCount Then
            datProtocol.Recordset.Delete
            'Удаление лишних записей из  "Базы Протокола"
        Else
            For lngRowNum = lngRowNum To intProtocolBaseCount Step 1
                datProtocol.Recordset.Delete
                datProtocol.Recordset.MoveNext
            Next
        End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
        lngRowNum = 1
            'Восстановить стандартный курсор мыши
        frmPrintPreview.MousePointer = 0
            'Дата "To"
        txtDateTo.Text = Format(Now, "dd/mm/yyyy")
            'Очистка текстовых полей ввода
        txtName.Text = ""
        txtCodeOrPassword.Text = ""
        txtReservOrNote.Text = ""
            'Сделать видимой кнопку "Find"
        cmdFind.Visible = True
            'Установить  опцию "PrintPage"
        chkPrintPage.Value = 0
            'Сделать видимой  опцию "PrintPage"
        chkPrintPage.Visible = True
            'Сделать видимым поле ввода "Name"
        txtName.Visible = True
            'Сделать видимым поле ввода "CodeOrPassword"
        txtCodeOrPassword.Visible = True
            'Сделать видимыми поля ввода "Date"
        txtDateFrom.Visible = True
        txtDateTo.Visible = True
            'Сделать видимыми метки полей ввода "Date"
        lblDateFrom.Visible = True
        lblDateTo.Visible = True
            'Сделать видимым поле ввода "ReservOrNote"
        txtReservOrNote.Visible = True
    
            'Установить смещение для строки печати
        frmPrintPreview.CurrentY = 1350
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = 7
             'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
   End If
    
End Sub

            'Обработка события "Cancel"
Private Sub cmdCancel_Click()
            'Закрыть выбранный  файл "Архив протокола", если он открыт
    If intFileNum <> Empty Then Close intFileNum
            'Убрать с экрана форму
    frmPrintPreview.Hide

End Sub
            
            'Обработка события нажатия кнопки "Find"
Private Sub cmdFind_Click()
            'Ключ-дата Протокола
Dim lngDateProtocol As Long
            'Ключ-дата Начальная
Dim lngDateFrom As Long
            'Ключ-дата Конечная
Dim lngDateTo As Long
            'Кнопка "Find" доступна для нажатия
    If cmdFind.MousePointer = 0 Then
            'Сделать недоступным нажатие на кнопку "Find"
        cmdFind.MousePointer = vbNoDrop
                'Неверный формат введенной даты "From"
        If Mid(Trim(txtDateFrom.Text), 3, 1) <> "." Or Mid(Trim(txtDateFrom.Text), 6, 1) <> "." _
        Or Len(Trim(txtDateFrom.Text)) <> 10 Then
            'Установка фокуса на текстовое поле "DateFrom"
            txtDateFrom.SetFocus
            'Звуковой сигнал
            frmDemo.BeepSound
            'Досрочный выход из процедуры
            Exit Sub
        End If
                'Неверный формат введенной даты "To"
        If Mid(Trim(txtDateTo.Text), 3, 1) <> "." Or Mid(Trim(txtDateTo.Text), 6, 1) <> "." _
        Or Len(Trim(txtDateTo.Text)) <> 10 Then
            'Установка фокуса на текстовое поле "DateTo"
            txtDateTo.SetFocus
            'Звуковой сигнал
            frmDemo.BeepSound
            'Досрочный выход из процедуры
            Exit Sub
        End If
            'Очистить форму и принтер
        frmPrintPreview.Cls
        Printer.EndDoc
            'Установить смещение для строки печати
        frmPrintPreview.CurrentY = 1350
        If chkPrintPage.Value = 1 Then Printer.CurrentY = 5
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = 7
            'Печать заголовков столбцов на форму
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
             'Печать заголовков столбцов на принтер
        If chkPrintPage.Value = 1 Then Printer.Print Tab(3); "Name"; Tab(25); _
        "Code or Password"; Tab(55); "Status"; Tab(70); "Time"; Tab(95); "Date"; _
        Tab(100); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
        If chkPrintPage.Value = 1 Then Printer.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Ключ-дата Начальная
        lngDateFrom = CLng(Mid(Trim(txtDateFrom.Text), 7, 4)) * 365 + _
        CLng(Mid(Trim(txtDateFrom.Text), 4, 2)) * 31 + _
        CLng(Mid(Trim(txtDateFrom.Text), 1, 2))
            'Ключ-дата Конечная
        lngDateTo = CLng(Mid(Trim(txtDateTo.Text), 7, 4)) * 365 + _
        CLng(Mid(Trim(txtDateTo.Text), 4, 2)) * 31 + _
        CLng(Mid(Trim(txtDateTo.Text), 1, 2))
            'Цикл по всем строкам "Таблицы протокола"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, lngRowNum, gProtocol
            'Обработка ошибки - Неверный формат даты в Протоколе
            If Val(Mid(gProtocol.strProtocDate, 7, 4)) = 0 Or _
            Val(Mid(gProtocol.strProtocDate, 4, 2)) = 0 Or _
            Val(Mid(gProtocol.strProtocDate, 1, 2)) = 0 Then
            'Ключ-дата Протокола - Условная
                lngDateProtocol = lngDateFrom
            Else
            'Ключ-дата Протокола
            lngDateProtocol = CLng(Mid(gProtocol.strProtocDate, 7, 4)) * 365 + _
                CLng(Mid(gProtocol.strProtocDate, 4, 2)) * 31 + _
                CLng(Mid(gProtocol.strProtocDate, 1, 2))
            End If
            
            'Строка "Таблицы протокола" удовлетворяет критерию поиска
            If Len(Trim(txtName.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocName), _
            Trim(txtName.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            'Вывести строку "Таблицы протокола" на форму
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            'Вывести строку "Таблицы протокола" на принтер
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            'Строка "Таблицы протокола" удовлетворяет критерию поиска
            ElseIf Len(Trim(txtCodeOrPassword.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocPersonCode), _
            Trim(txtCodeOrPassword.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            'Вывести строку "Таблицы протокола" на форму
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            'Вывести строку "Таблицы протокола" на принтер
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            'Строка "Таблицы протокола" удовлетворяет критерию поиска
            ElseIf Len(Trim(txtReservOrNote.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocReserve), _
            Trim(txtReservOrNote.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            'Вывести строку "Таблицы протокола"
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            'Вывести строку "Таблицы протокола" на принтер
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            'Строка "Таблицы протокола" удовлетворяет критерию поиска
            ElseIf Len(Trim(txtName.Text)) = 0 And Len(Trim(txtCodeOrPassword.Text)) = 0 _
            And Len(Trim(txtReservOrNote.Text)) = 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            'Вывести строку "Таблицы протокола"
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            'Вывести строку "Таблицы протокола" на принтер
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            End If
            
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступным нажатие на кнопку "Find"
                If lngRowNum < gProtocRowNum - 1 Then cmdFind.MousePointer = 0
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
            'Вывести данные на печать
        If chkPrintPage.Value = 1 Then Printer.EndDoc
    End If
    
End Sub

            'Нажата кнопка "Next"
Private Sub cmdNext_Click()
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = lngRowNum + 1
    
            'Очистить форму
    frmPrintPreview.Cls
            'Сделать доступной кнопку "Previous"
    cmdPrevious.Enabled = True
            'Сделать недоступной кнопку "Next"
    cmdNext.Enabled = False

            
            'Предварительная печать "Системной таблицы" на форму "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            'Вывести строку "Системной таблицы"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы персон" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            'Статус - Клиент Автостоянки
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            'Вывести строку "Таблицы персон"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы календаря" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(95); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            'Вывести строку "Таблицы календаря"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(95); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
        Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Архива протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Архива протокола"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            'Читать строку "Архива протокола" из файла в буфер
            Get intFileNum, lngRowNum, gProtocol
            'Вывести строку "Архива протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
                
                'Предварительная печать "Таблицы времени" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            'Вывести строку "Таблицы времени"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы терминалов" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            'Вывести строку "Таблицы терминалов"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
End If

End Sub
            
            'Нажата кнопка "Previous"
Private Sub cmdPrevious_Click()
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "Архива протокола")
    lngRowNum = lngRowNum - (intRowPrintQuan + intRowPrintNum - 5) + 1
    If intRowPrintNum <= intRowPrintQuan Then lngRowNum = lngRowNum - 1
    If lngRowNum < 1 Then lngRowNum = 1
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
    
            'Очистить форму
    frmPrintPreview.Cls
            'Сделать недоступной кнопку "Previous"
    If lngRowNum = 1 Then cmdPrevious.Enabled = False
            'Сделать доступной кнопку "Next"
    cmdNext.Enabled = True

            
            'Предварительная печать "Системной таблицы" на форму "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Системной таблицы"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            'Вывести строку "Системной таблицы"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы персон" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            'Статус - Клиент Автостоянки
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            'Вывести строку "Таблицы персон"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы календаря" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            'Вывести строку "Таблицы календаря"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
        Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Архива протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Архива протокола"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            'Читать строку "Архива протокола" из файла в буфер
            Get intFileNum, lngRowNum, gProtocol
            'Вывести строку "Архива протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
                
                'Предварительная печать "Таблицы времени" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            'Вывести строку "Таблицы времени"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы терминалов" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            'Вывести строку "Таблицы терминалов"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
    End If

End Sub
            
            'Нажата кнопка "First"
Private Sub cmdFirst_Click()
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
            'Номер первой строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = 1
    
            'Очистить форму
    frmPrintPreview.Cls
            'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
            'Сделать недоступной кнопку "Next"
    cmdNext.Enabled = False

            
            'Предварительная печать "Системной таблицы" на форму "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            'Вывести строку "Системной таблицы"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы персон" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            'Статус - Клиент Автостоянки
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            'Вывести строку "Таблицы персон"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы календаря" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            'Вывести строку "Таблицы календаря"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
        Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Архива протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Архива протокола"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            'Читать строку "Архива протокола" из файла в буфер
            Get intFileNum, lngRowNum, gProtocol
            'Вывести строку "Архива протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
                
                'Предварительная печать "Таблицы времени" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            'Вывести строку "Таблицы времени"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы терминалов" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            'Вывести строку "Таблицы терминалов"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
End If

End Sub
            
            'Нажата кнопка "Last"
Private Sub cmdLast_Click()
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
    
            'Очистить форму
    frmPrintPreview.Cls
            'Сделать доступной кнопку "Previous"
    cmdPrevious.Enabled = True
            'Сделать недоступной кнопку "Next"
    cmdNext.Enabled = False
            
            
            'Предварительная печать "Системной таблицы" на форму "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Системной таблицы"
    lngRowNum = frmTableSystem.grdTableSystem.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            'Вывести строку "Системной таблицы"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            'Предварительная печать "Таблицы персон" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Таблицы персон"
    lngRowNum = gTablePerson.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            'Статус - Клиент Автостоянки
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            'Вывести строку "Таблицы персон"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            'Предварительная печать "Таблицы календаря" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Таблицы календаря"
    lngRowNum = frmTableCalendar.grdTableCalendar.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            'Вывести строку "Таблицы календаря"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Таблицы протокола"
        lngRowNum = gProtocRowNum + 3 - intRowPrintQuan
        If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
            cmdPrevious.Enabled = False
            lngRowNum = 1
        End If
            'Цикл по всем строкам последней страницы "Таблицы протокола"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
        Next
    
            'Предварительная печать "Архива протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            
            'Номер первой строки последней страницы "Таблицы протокола"
        lngRowNum = lngArchivesRowNum + 4 - intRowPrintQuan
        If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
            cmdPrevious.Enabled = False
            lngRowNum = 1
        End If
            'Цикл по всем строкам "Архива протокола"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            'Читать строку "Архива протокола" из файла в буфер
            Get intFileNum, lngRowNum, gProtocol
            'Вывести строку "Архива протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
        Next
                
                'Предварительная печать "Таблицы времени" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Таблицы времени"
    lngRowNum = frmTableTime.grdTableTime.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            'Вывести строку "Таблицы времени"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            'Предварительная печать "Таблицы терминалов" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Номер первой строки последней страницы "Таблицы терминалов"
    lngRowNum = frmTableTerminal.grdTableTerminal.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                'Сделать недоступной кнопку "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            'Вывести строку "Таблицы терминалов"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
End If

End Sub
            'Нажата кнопка "PrintPage"
Private Sub cmdPrintPage_Click()
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = lngRowNum - (intRowPrintNum - 3) + 1
    If intRowPrintNum <= intRowPrintQuan Then lngRowNum = lngRowNum - 1
    If lngRowNum < 1 Then lngRowNum = 1
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
    
            'Очистить принтер от "остатков" предыдущей печати
    Printer.EndDoc
    
            'Печать "Системной таблицы"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
        Tab(70); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Системной таблицы"
            For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
                frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
                strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
            Next
            'Вывести на печать строку "Системной таблицы"
            Printer.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
            Tab(55); strTableSystem(2); Tab(70); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            'Печать "Таблицы персон"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
        Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
        For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
            gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
            For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
                gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
                strTablePerson(intColNum) = gTablePerson.Text
            Next
            'Статус - Клиент Автостоянки
            If Left(Trim(strTablePerson(2)), 2) = "07" Or _
            Left(Trim(strTablePerson(2)), 2) = "05" Or _
            Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
                strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
            ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
            Left(Trim(strTablePerson(2)), 2) = "08" Or _
            Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
                strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
            End If
            'Вывести на печать строку "Таблицы персон"
            Printer.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
            Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
            Tab(115); strTablePerson(5)
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            'Печать "Таблицы календаря"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
        Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
        Tab(115); "Sunday"
            'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
        For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
            For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
                frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
                strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
            Next
            'Вывести на печать строку "Таблицы календаря"
            Printer.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
            Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
            Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
                'Печать "Таблицы протокола"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(75); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
           'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
            Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
            Printer.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(75); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
        
            'Печать "Таблицы времени"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
        For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
            frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
            For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
                frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
                strTableTime(intColNum) = frmTableTime.grdTableTime.Text
            Next
            'Вывести на печать строку "Таблицы времени"
            Printer.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); Tab(55); strTableTime(2)
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            'Печать "Таблицы терминалов"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
        Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
        Tab(70); "Expander"
            'Вывести пустую строку
        Printer.Print
            'Текущий номер строки на странице печати
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
        For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
            For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
                frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
                strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
            Next
            'Вывести на печать строку "Таблицы терминалов"
            Printer.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
            Tab(55); strTableTerminal(2); Tab(70); strTableTerminal(3)
            'Текущий номер строки на странице печати
            intRowPrintNum = intRowPrintNum + 1
            'Страница печати заполнена - завершить печать
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
    End If
    
            'Данных для печати больше нет
    Printer.EndDoc

End Sub

            'Форма "Активирована"
Private Sub Form_Activate()
             'Полное имя файла "Архив протокола" (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Архива протокола"
Dim lngRecordLen As Long
            'Длина в байтах "Архива протокола"
Dim lngFileLength As Long

            'Выбрать принтер "По Умолчанию - широкий"
    Set Printer = Printers(0)

            'Сделать "Пустым" номер файла "Архив протокола"
    intFileNum = Empty
           'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = 1
            'Количество строк на одной странице формы "frmPrintPreview"
    intRowPrintQuan = gRowPrintQuan
            'Текущий номер строки таблиц ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "Архив протокола")
    lngRowNum = 1
            
            'Очистить форму
    frmPrintPreview.Cls
            'Предварительная печать "Системной таблицы" на форму "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            'По всем столбцам "Системной таблицы"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            frmTableSystem.grdTableSystem.Col = intColNum
            'Заполнение буфера для печати строки "Системной таблицы"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            'Вывести строку "Системной таблицы"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы персон" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = lngRowNum
            'По всем столбцам "Таблицы персон"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            'Текущий столбец "Таблицы персон"
            gTablePerson.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы персон"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            'Статус - Клиент Автостоянки
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            'Статус - Посетитель Предприятия
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            'Маскирование Запакованных символов в поле Резерв (Примечание)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            'Вывести строку "Таблицы персон"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы календаря" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы календаря"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            'Текущая строка "Таблицы календаря"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            'По всем столбцам "Таблицы календаря"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            'Текущий столбец "Таблицы календаря"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы календаря"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            'Вывести строку "Таблицы календаря"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
                'Предварительная печать "Таблицы протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            'Сделать видимой опцию "Connect Protocol Base"
    chkProtocol.Visible = True
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Таблицы протокола"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            'Читать строку "Таблицы протокола" из файла в буфер
        Get gProtocFileNum, lngRowNum, gProtocol
            'Вывести строку "Таблицы протокола"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Архива протокола" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then

                'Полное имя файла "Архив протокола" (с указанием "пути" к нему)
        strPathFileName = gPathFileName
            'Вычислить длину записи (строки) "Архива протокола"
        lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Открыть выбранный архивный файл для произвольного доступа
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Определить длину в байтах выбранного файла "Архив протокола"
        lngFileLength = LOF(intFileNum)
            'Вычислить количество записей в выбранном файле "Архив протокола"
        lngArchivesRowNum = lngFileLength / lngRecordLen
        
            'Печать заголовков столбцов
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            'Вывести пустую строку
        frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем строкам "Архива протокола"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            'Читать строку "Архива протокола" из файла в буфер
            Get intFileNum, lngRowNum, gProtocol
            'Вывести строку "Архива протокола"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            'Текущий номер строки формы "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
            If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
                Exit For
            End If
        Next
    
            'Предварительная печать "Таблицы времени" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        frmTableTime.grdTableTime.Row = lngRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
            frmTableTime.grdTableTime.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы времени"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            'Вывести строку "Таблицы времени"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
            'Предварительная печать "Таблицы терминалов" на форму "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            'Печать заголовков столбцов
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            'Вывести пустую строку
    frmPrintPreview.Print
            'Текущий номер строки формы "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            'Заполнение буфера для печати строки "Таблицы терминалов"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            'Вывести строку "Таблицы терминалов"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            'Текущий номер строки формы "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            'Страница формы "frmPrintPreview" заполнена
        If intRowPrintNum > intRowPrintQuan Then
            'Сделать доступной кнопку "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            'Приостановить вывод строк на форму "frmPrintPreview"
            Exit For
        End If
    Next
    
    End If

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Контроль ввода нецифровой информации в текстовое поле "DateFrom"
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
            'Введен нецифровой символ
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            'Отмена символа
        KeyAscii = 0
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
                'Неверный формат введенной даты
    If (KeyAscii = 0 Or KeyAscii = vbKeyReturn) And _
    (Mid(Trim(txtDateFrom.Text), 3, 1) <> "." Or Mid(Trim(txtDateFrom.Text), 6, 1) <> "." _
    Or Len(Trim(txtDateFrom.Text)) <> 10) Then
            'Восстановление формата даты "From"
        txtDateFrom.Text = Trim(strDateFrom)
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
            'Кнопка "Find" доступна для нажатия
    cmdFind.MousePointer = 0
    
End Sub

            'Контроль ввода нецифровой информации в текстовое поле "DateTo"
Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
            'Введен нецифровой символ
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            'Отмена символа
        KeyAscii = 0
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Неверный формат введенной даты
    If (KeyAscii = 0 Or KeyAscii = vbKeyReturn) And _
    (Mid(Trim(txtDateTo.Text), 3, 1) <> "." Or Mid(Trim(txtDateTo.Text), 6, 1) <> "." _
    Or Len(Trim(txtDateTo.Text)) <> 10) Then
            'Восстановление формата даты "To"
        txtDateTo.Text = Format(Now, "dd/mm/yyyy")
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
            'Кнопка "Find" доступна для нажатия
    cmdFind.MousePointer = 0
    
End Sub

            'Контроль длины введенной информации в текстовое поле "CodeOrPassword"
Private Sub txtCodeOrPassword_Change()
            'Длина введенной строки больше допустимой
    If Len(Trim(txtCodeOrPassword.Text)) > 16 Then
            'Очистка текстового поля
        txtCodeOrPassword.Text = ""
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
            'Кнопка "Find" доступна для нажатия
    cmdFind.MousePointer = 0

End Sub

            'Контроль длины введенной информации в текстовое поле "Name"
Private Sub txtName_Change()
            'Длина введенной строки больше допустимой
    If Len(Trim(txtName.Text)) > 16 Then
            'Очистка текстового поля
        txtName.Text = ""
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
            'Кнопка "Find" доступна для нажатия
    cmdFind.MousePointer = 0

End Sub

            'Контроль длины введенной информации в текстовое поле "ReservOrNote"
Private Sub txtReservOrNote_Change()
            'Длина введенной строки больше допустимой
    If Len(Trim(txtReservOrNote.Text)) > 22 Then
            'Очистка текстового поля
        txtReservOrNote.Text = ""
            'Звуковой сигнал
        frmDemo.BeepSound
    End If
            'Восстановить текущий номер строки таблицы "TableProtocol"
    lngRowNum = 1
            'Кнопка "Find" доступна для нажатия
    cmdFind.MousePointer = 0

End Sub
