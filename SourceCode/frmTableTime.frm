VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableTime 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_time"
   ClientHeight    =   4935
   ClientLeft      =   2640
   ClientTop       =   2745
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
   ScaleHeight     =   4935
   ScaleWidth      =   6615
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2400
      Max             =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox txtExpander 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4320
      TabIndex        =   19
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Frame fraColName 
      Caption         =   "Options"
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
      Height          =   1452
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optTime 
         Caption         =   "Time"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   732
      End
      Begin VB.OptionButton optExpander 
         Caption         =   "Expander"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1452
      End
      Begin VB.CheckBox chkFromToTime 
         Caption         =   "v From     To"
         Enabled         =   0   'False
         Height          =   492
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.HScrollBar hsbMinute 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4320
      Max             =   59
      TabIndex        =   12
      Top             =   840
      Width           =   1452
   End
   Begin VB.HScrollBar hsbHour 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4320
      Max             =   23
      TabIndex        =   9
      Top             =   480
      Width           =   1452
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete..."
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
      TabIndex        =   8
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
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
      TabIndex        =   7
      Top             =   2880
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
      Left            =   3720
      TabIndex        =   6
      Top             =   4320
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4320
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
      TabIndex        =   4
      Top             =   4320
      Width           =   1212
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
      TabIndex        =   2
      Top             =   360
      Width           =   1092
   End
   Begin VB.ListBox lstInterval 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableTime.frx":0000
      Left            =   120
      List            =   "frmTableTime.frx":0002
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableTime 
      Height          =   1695
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   3
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
      Left            =   3600
      TabIndex        =   24
      Top             =   2040
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
      Left            =   2040
      TabIndex        =   23
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<==="
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
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Intervals variant"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6480
      X2              =   1920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1920
      X2              =   1920
      Y1              =   1800
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1920
      X2              =   120
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   4200
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   1800
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Label lblMinute59 
      Alignment       =   2  'Center
      Caption         =   "59min"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblMinute0 
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
      Left            =   4080
      TabIndex        =   13
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblHour23 
      Alignment       =   2  'Center
      Caption         =   "23h"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblHour0 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblIntervals 
      Alignment       =   2  'Center
      Caption         =   "Intervals "
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmTableTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Текущий номер корректируемой строки "Таблицы времени"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Таблицы времени"
Dim intColNumCorr As Integer
            '"Старый" номер варианта "Таблицы времени"
Dim intVariantOld As Integer
            '"Новый" номер варианта "Таблицы времени"
Dim intVariantNew As Integer
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Таблицы времени - интервалов"
Dim gTime As TimeInfo

            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            '"Старый" номер варианта "Таблицы времени" не нулевой
    If hsbVariant.Value <> 0 Then

            'Были несохраненные изменения в "Таблице времени"
        If gChangesTableTime = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы времени" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение "Таблицы времени" в файле по умолчанию
                cmdSave_Click
            End If
        End If
            
    Else

            'Были несохраненные изменения в "Таблице времени"
        If gChangesTableTime = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы времени" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение "Таблицы времени" в файле по умолчанию
                cmdSave_Click
            End If
        End If
            
            'Сделать недоступными элементы управления Коррекцией "Таблицы времени"
        lblIntervals.Enabled = False
        lstInterval.Enabled = False
        fraColName.Enabled = False
        optTime.Enabled = False
        chkFromToTime.Enabled = False
        optExpander.Enabled = False
        lblHour0.Enabled = False
        lblHour23.Enabled = False
        lblMinute0.Enabled = False
        lblMinute59.Enabled = False
        hsbHour.Enabled = False
        hsbMinute.Enabled = False
        txtExpander.Enabled = False
    End If
    
            'Очистить текстовые поля
    txtExpander.Text = ""
            'Очистить список интервалов
    lstInterval.Clear
            
            'Сделать нулевым текущий вариант "Таблицы интервалов"
    intVariantOld = 0
    hsbVariant.Value = 0
            'Сбросить признак внесенных изменений в "Таблицу времени"
    gChangesTableTime = False

            'Сделать невидимой текущую форму
    frmTableTime.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()

            ' "Таблица времени" не содержит нефиксированных строк
    If grdTableTime.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        MsgBox ("The table is empty")
    
    Else
            'Сделать доступными некоторые элементы управления Коррекцией "Таблицы времени"
        lblIntervals.Enabled = True
        lstInterval.Enabled = True
        fraColName.Enabled = True
        optTime.Enabled = True
        optTime.Value = True
        chkFromToTime.Enabled = True
        optExpander.Enabled = True
        lblHour0.Enabled = True
        lblHour23.Enabled = True
        lblMinute0.Enabled = True
        lblMinute59.Enabled = True
        hsbHour.Enabled = True
        hsbMinute.Enabled = True
            'Очистить текстовые поля
        txtExpander.Text = ""
            'Очистить список имен
        lstInterval.Clear
    
            'Столбец "Intervals"
        grdTableTime.Col = 0
            'Цикл по всем нефиксированным строкам "Таблицы времени"
        For intRowNumCorr = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
            grdTableTime.Row = intRowNumCorr
            'Заполнение списка "lstInterval" записями из "Таблицы времени"
            lstInterval.AddItem grdTableTime.Text
        Next
            'Выбрать  элемент списка
        lstInterval.ListIndex = 0
            'Номер корректируемой строки - (1)
        intRowNumCorr = 1
        grdTableTime.Row = intRowNumCorr
            'Включить опцию
        optTime_Click
    End If
    
End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Таблицы времени"
Private Sub grdTableTime_Click()
            'Коррекция "включена"
    If lstInterval.Enabled = True Then
            'Номер корректируемой строки "Таблицы времени"
        intRowNumCorr = grdTableTime.RowSel
        grdTableTime.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstInterval.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Таблицы времени"
        intColNumCorr = grdTableTime.ColSel
        grdTableTime.Col = intColNumCorr
            'Выбор корректируемой строки "Таблицы времени"
        lstInterval_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstInterval.Left, Y:=lstInterval.Top
            'Выбор корректируемого столбца "Таблицы времени"
        Select Case intColNumCorr
            Case 1
            optTime.Value = True
            Case 2
            optExpander.Value = True
            'Установить фокус на текстовом поле для Коррекции
            txtExpander.SetFocus
        End Select
    End If
        
End Sub

            'Выбор корректируемой строки "Таблицы времени"
Private Sub lstInterval_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемой строки "Таблицы времени"
        intRowNumCorr = lstInterval.ListIndex + 1
        grdTableTime.Row = intRowNumCorr
        grdTableTime.Col = 2
            'Копирование ячейки "Таблицы времени" в текстовое поле для Коррекции
        txtExpander.Text = grdTableTime.Text
            'Восстановить номер корректируемого столбца "Таблицы времени"
        grdTableTime.Col = intColNumCorr
    End If

End Sub

            'Выбрана опция - "Time"
Private Sub optTime_Click()
            'Номер корректируемого столбца "Таблицы времени"
    intColNumCorr = 1
    grdTableTime.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы времени"
    lblHour0.Enabled = True
    lblHour23.Enabled = True
    lblMinute0.Enabled = True
    lblMinute59.Enabled = True
    hsbHour.Enabled = True
    hsbMinute.Enabled = True
    txtExpander.Enabled = False

End Sub

            'Выбрана опция "Expander"
Private Sub optExpander_Click()
            'Номер корректируемого столбца "Таблицы времени"
    intColNumCorr = 2
    grdTableTime.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы времени"
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
            'Копирование ячейки "Таблицы времени" в текстовое поле для Коррекции
    txtExpander.Text = grdTableTime.Text
    txtExpander.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtExpander.SetFocus

End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Hour"
Private Sub hsbHour_Scroll()
    hsbHour_Change
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Hour"
Private Sub hsbHour_Change()
            'Начало временного интервала
    If chkFromToTime.Value = 1 Then
            'Изменение ячейки "Time" в "Таблице времени"
        If hsbHour.Value < 10 Then
            grdTableTime.Text = "0" + Trim(Str(hsbHour.Value)) + Mid(grdTableTime.Text, 3)
        Else
            grdTableTime.Text = Trim(Str(hsbHour.Value)) + Mid(grdTableTime.Text, 3)
        End If
            'Конец временного интервала
    Else
            'Изменение ячейки "Time" в "Таблице времени"
        If hsbHour.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 6) + "0" + Trim(Str(hsbHour.Value)) _
            + Mid(grdTableTime.Text, 9)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 6) + Trim(Str(hsbHour.Value)) _
            + Mid(grdTableTime.Text, 9)
        End If
    End If
            'Установить признак  внесенных изменений в "Таблицу времени"
    gChangesTableTime = True
    
End Sub
            
            'Обработка события "Scroll" - прокрутка для ползунка "Minute"
Private Sub hsbMinute_Scroll()
    hsbMinute_Change
End Sub
            
            'Обработка события "Change" - прокрутка для ползунка "Minute"
Private Sub hsbMinute_Change()
            'Начало временного интервала
    If chkFromToTime.Value = 1 Then
            'Изменение ячейки "Time" в "Таблице времени"
        If hsbMinute.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 3) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 6)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 3) + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 6)
        End If
            'Конец временного интервала
    Else
            'Изменение ячейки "Time" в "Таблице времени"
        If hsbMinute.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 9) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 12)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 9) + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 12)
        End If
    End If
            'Установить признак  внесенных изменений в "Таблицу времени"
    gChangesTableTime = True

End Sub
            
            'Процедура ввода и анализа Корректируемого поля "Expander"
Private Sub txtExpander_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtExpander.Text)) < 9 Then
            'Изменение поля "Expander" в "Таблице времени"
            grdTableTime.Text = Trim(txtExpander.Text)
            'Установить признак  внесенных изменений в "Таблицу времени"
            gChangesTableTime = True
            'Неверный формат данных
            'Установить фокус на кнопке "Save"
            cmdSave.SetFocus
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            'Обработка события "Change" - прокрутка для ползунка "Variant"
Private Sub hsbVariant_Change()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы времени"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы времени"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы времени"
Dim intColNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Установить ширину столбцов
    SetColWidth
            'Сделать недоступными элементы управления Коррекцией "Таблицы времени"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            'Очистить текстовые поля
    txtExpander.Text = ""
            'Очистить список интервалов
    lstInterval.Clear
            
            'Были несохраненные изменения в "Таблице времени"
    If gChangesTableTime = True Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы времени" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранить "Новый" номер варианта "Таблицы времени"
            intVariantNew = hsbVariant.Value
            '"Старый" номер варианта "Таблицы времени"
            hsbVariant.Value = intVariantOld
            'Сохранение "Таблицы времени" в файле по умолчанию
            cmdSave_Click
            'Восстановить "Новый" номер варианта "Таблицы времени"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"Старый" номер варианта "Таблицы времени"
    intVariantOld = hsbVariant.Value
            'Количество удалений/добавлений строк в "Таблице времени"
    gAddDelRowTableTime = 0
            'Сбросить признак внесенных изменений в "Таблицу времени"
    gChangesTableTime = False

            'Заполнение варианта "Таблицы времени" из файла
            
            'Вычислить длину записи (строки) "Таблицы времени"
    lngRecordLen = Len(gTime)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
                
                                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableTime
            'Количество строк "Таблицы времени" равно размеру файла по умолчанию +1
    grdTableTime.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам  варианта "Таблицы времени"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        grdTableTime.Row = intRowNum
            'Читать строку "Таблицы времени" из файла в буфер
        Get intFileNum, intRowNum, gTime
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
            grdTableTime.Col = intColNum
            'Заполнение текущей строки "Таблицы времени" из буфера
            Select Case intColNum
                Case 0
                grdTableTime.Text = gTime.strIntervalNum
                Case 1
                 grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                Case 2
                grdTableTime.Text = gTime.strExpander
            End Select
        Next
    Next
            'Закрыть файл
    Close intFileNum
    
            'Установить фокус на кнопке "Correction"
    cmdCorrection.SetFocus
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = hsbVariant.Value
    
    Exit Sub
ErrorTableTime:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableTime Error !")
    
End Sub
            
            'Добавление строки в "Таблицу времени"
Private Sub cmdAdd_Click()
    
            'Сделать недоступными элементы управления Коррекцией "Таблицы времени"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            'Очистить список интервалов
    lstInterval.Clear
    
            'Формирование номера интервала времени
    gTime.strIntervalNum = "Interval-" + Str(grdTableTime.Rows)
            'Добавление строки в конец "Таблицы времени"
    grdTableTime.AddItem gTime.strIntervalNum
            'Формирование шаблона для интервала времени
    grdTableTime.Row = grdTableTime.Rows - 1
    grdTableTime.Col = 1
    grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
    "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
            'Количество удалений/добавлений строк в "Таблице времени"
    gAddDelRowTableTime = gAddDelRowTableTime + 1
            'Установить признак внесенных изменений в "Таблицу времени"
    gChangesTableTime = True
            'Установить фокус на кнопке "Add"
    cmdAdd.SetFocus
    
End Sub
            
            'Удаление строки из "Таблицы времени"
Private Sub cmdDelete_Click()
            'Текущий номер нефиксированной строки "Таблицы времени"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы времени"
Dim intColNum As Integer
    
            'Сделать недоступными элементы управления Коррекцией "Таблицы времени"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            'Очистить список интервалов
    lstInterval.Clear
    
            'Загрузить (не показывая) форму "frmSelectRow"
    Load frmSelectRow
            'Инициализировать этикетку "lblColName" формы "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Interval"
    
            'Столбец "Intervals"
    grdTableTime.Col = 0
             'Очистить список объектов
    frmSelectRow.lstSelectRow.Clear
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        grdTableTime.Row = intRowNum
            'Заполнение списка "lstSelectRow" записями из "Таблицы времени"
        frmSelectRow.lstSelectRow.AddItem grdTableTime.Text
    Next
            'Выбрать элемент списка
    frmSelectRow.lstSelectRow.ListIndex = 0
            'Вывести на экран форму "frmSelectRow" с уровнем модальности 1
    frmSelectRow.Show 1
            'Строка не выбрана
    If frmSelectRow.Tag = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The row isn't selected !"
            'Последняя строка не может быть удалена
    ElseIf frmSelectRow.lstSelectRow.ListCount = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The last row isn't selected !"
            'Удаление строки из "Таблицы времени"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            'Номер удаляемой строки
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            'Удаление строки
        grdTableTime.RemoveItem intRowNum
            'Удаленная строка не последняя в "Таблице времени"
        If intRowNum < grdTableTime.Rows Then
             'Цикл по всем строкам "Таблицы времени", начиная за удаленной строкой
            For intRowNum = intRowNum To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
                grdTableTime.Row = intRowNum
            'Заполнение списка "lstSelectRow" записями из "Таблицы времени"
                grdTableTime.Text = "interval-" + Str(intRowNum)
            Next
        End If
           'Количество удалений/добавлений строк в "Таблице времени"
        gAddDelRowTableTime = gAddDelRowTableTime - 1
            'Установить признак внесенных изменений в "Таблицу времени"
        gChangesTableTime = True
    End If
            'Выгрузить форму "frmSelectRow"
    UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
    Set frmSelectRow = Nothing
            'Установить фокус на кнопке "Delete"
    cmdDelete.SetFocus
    
End Sub
            
            'Сохранение "Таблицы времени" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы времени"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы времени"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы времени"
Dim intColNum As Integer
            'Вычислить длину записи (строки) "Таблицы времени"
    lngRecordLen = Len(gTime)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
   
            'Строк, удаленных из "Таблицы времени" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
    If gAddDelRowTableTime < 0 Then
            'Удалить "старый" умалчиваемый файл
        Kill strPathFileName
    End If
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        grdTableTime.Row = intRowNum
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
            grdTableTime.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы времени" в файл
            Select Case intColNum
                Case 0
                gTime.strIntervalNum = grdTableTime.Text
                Case 1
                gTime.strTime = Left(grdTableTime.Text, 2) + Mid(grdTableTime.Text, 4, 2) + _
                Mid(grdTableTime.Text, 7, 2) + Mid(grdTableTime.Text, 10, 2)
                Case 2
                gTime.strExpander = grdTableTime.Text
            End Select
        Next
            'Записать строку "Таблицы времени" в файл
        Put intFileNum, intRowNum, gTime
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице времени"
    gAddDelRowTableTime = 0
            'Сбросить признак внесенных изменений в "Таблицу времени"
    gChangesTableTime = False
            'Установить фокус на кнопке "Cancel"
    cmdCancel.SetFocus
            
End Sub
            
            'Сохранение "Таблицы времени" в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы времени"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы времени"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы времени"
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
            'Запись "Таблицы времени" в выбранный файл
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Таблицы времени"
        lngRecordLen = Len(gTime)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Строк, удаленных из "Таблицы времени" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
        If gAddDelRowTableTime < 0 Then
            'Удалить "старый" файл, если он существует
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы времени"
        For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
            grdTableTime.Row = intRowNum
            'По всем столбцам "Таблицы времени"
            For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
                grdTableTime.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы времени" в файл
                Select Case intColNum
                    Case 0
                    gTime.strIntervalNum = grdTableTime.Text
                    Case 1
                    gTime.strTime = Left(grdTableTime.Text, 2) + Mid(grdTableTime.Text, 4, 2) + _
                    Mid(grdTableTime.Text, 7, 2) + Mid(grdTableTime.Text, 10, 2)
                    Case 2
                    gTime.strExpander = grdTableTime.Text
                End Select
            Next
            'Записать строку "Таблицы времени" в файл
        Put intFileNum, intRowNum, gTime
        Next
            'Закрыть умалчиваемый файл
        Close intFileNum
            'Количество удалений/добавлений строк в "Таблице времени"
        gAddDelRowTableTime = 0
            'Сбросить признак внесенных изменений в "Таблицу времени"
        gChangesTableTime = False
    End If
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Установить фокус на кнопке "Cancel"
    cmdCancel.SetFocus
    
End Sub

            'Загрузка формы "Таблица времени"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы времени"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы времени"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы времени"
Dim intColNum As Integer
            'Количество столбцов в массиве интервалов доступа
Dim intIntervalNum As Integer

            'Установить ширину столбцов
    SetColWidth
            'Количество вариантов "Таблицы времени"
    lblVariant99.Caption = "V" + Str(gVarNumTime)
    hsbVariant.Max = gVarNumTime
            'Сохранить "Старый" номер варианта "Таблицы времени"
    intVariantOld = hsbVariant.Value
    
            'Текущая строка = 0 (Заголовки столбцов)
    grdTableTime.Row = 0
    grdTableTime.Col = 0
    grdTableTime.Text = "Intervals"
            'Записать в ячейку (строка 0, столбец 1)
    grdTableTime.Col = 1
    grdTableTime.Text = "Time"
            'Записать в ячейку (строка 0, столбец 2)
    grdTableTime.Col = 2
    grdTableTime.Text = "Expander"
            
            'Заполнение "Таблицы времени" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы времени"
    lngRecordLen = Len(gTime)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableTime
                'Количество строк "Таблицы времени" равно размеру файла по умолчанию +1
    grdTableTime.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы времени"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            'Текущая строка "Таблицы времени"
        grdTableTime.Row = intRowNum
            'Читать строку "Таблицы времени" из файла в буфер
        Get intFileNum, intRowNum, gTime
            'По всем столбцам "Таблицы времени"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            'Текущий столбец "Таблицы времени"
            grdTableTime.Col = intColNum
            'Заполнение текущей строки "Таблицы времени" из буфера
            Select Case intColNum
                Case 0
                grdTableTime.Text = gTime.strIntervalNum
                Case 1
                 grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                Case 2
                grdTableTime.Text = gTime.strExpander
            End Select
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
    
            'Максимальное количество столбцов в массивe интервалов доступа
    intIntervalNum = grdTableTime.Rows
            'Цикл по всем ненулевым вариантам "Таблицы времени"
    For intVariantNew = 1 To gVarNumTime Step 1
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTime" + Trim(Str(intVariantNew)) + ".dat"
    
            'Максимальное количество столбцов в массиве интервалов доступа
        If FileLen(strPathFileName) / lngRecordLen + 1 > intIntervalNum Then
            intIntervalNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    Next
            'Переопределить размерность массива интервалов доступа
ReDim gInterval(gVarNumTime + 1, intIntervalNum) As String * 11
            'Переопределить размерность массива дополнительных терминалов
            '  и календарей для всех вариантов "Таблицы времени"
ReDim gTerCal(gVarNumTime + 1, intIntervalNum) As String * 12
    
            'Цикл по всем вариантам "Таблицы времени"
    For intVariantNew = 0 To gVarNumTime Step 1
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTime" + Trim(Str(intVariantNew)) + ".dat"
    
            'Количество "значащих" столбцов в текущей строке массива
            '  интервалов доступа
        intIntervalNum = FileLen(strPathFileName) / lngRecordLen + 1
        gInterval(intVariantNew, 0) = intIntervalNum
            'Количество "значащих" столбцов в текущей строке массива
            '  дополнительных терминалов и календарей
        gTerCal(intVariantNew, 0) = intIntervalNum
    
            'Открыть файл для произвольного доступа
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам  варианта "Таблицы времени"
        For intRowNum = 1 To intIntervalNum - 1 Step 1
            'Читать строку "Таблицы времени" из файла в буфер
            Get intFileNum, intRowNum, gTime
                 gInterval(intVariantNew, intRowNum) = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                 gTerCal(intVariantNew, intRowNum) = gTime.strIntervalNum
        Next
            'Закрыть файл
        Close intFileNum
    Next
    
            'Количество удалений/добавлений строк в "Таблице времени"
    gAddDelRowTableTime = 0
            'Сбросить признак внесенных изменений в "Таблицу времени"
    gChangesTableTime = False
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableTime:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableTime Error !")
    
End Sub
            
            'Процедура установки ширины и выравнивания столбцов "Таблицы времени"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Цикл по всем столбцам
    For intColNumber = 0 To grdTableTime.Cols - 1 Step 1
        grdTableTime.ColWidth(intColNumber) = 970
        grdTableTime.ColAlignment(intColNumber) = 0
    Next
    
End Sub


