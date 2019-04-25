VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableTerminal 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_terminal"
   ClientHeight    =   6390
   ClientLeft      =   2400
   ClientTop       =   1485
   ClientWidth     =   7110
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
   ScaleHeight     =   6390
   ScaleWidth      =   7110
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   31
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdDefaultTerm 
      Cancel          =   -1  'True
      Caption         =   "DefaultTerm"
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
      Left            =   5760
      TabIndex        =   30
      Top             =   5760
      Width           =   1212
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
      Left            =   5880
      TabIndex        =   29
      Top             =   5040
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
      Left            =   4680
      TabIndex        =   28
      Top             =   5040
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
      TabIndex        =   27
      Top             =   5040
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
      TabIndex        =   26
      Top             =   5040
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   25
      Top             =   5760
      Width           =   1212
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2520
      Max             =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   1092
   End
   Begin VB.TextBox txtExpander 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5520
      TabIndex        =   17
      Top             =   2280
      Width           =   1212
   End
   Begin VB.TextBox txtDescription 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5520
      TabIndex        =   15
      Top             =   1800
      Width           =   1212
   End
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   288
      Left            =   6480
      TabIndex        =   12
      Top             =   1320
      Width           =   252
   End
   Begin VB.TextBox txtAddress 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
      Width           =   372
   End
   Begin VB.TextBox txtTerminal 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5280
      TabIndex        =   9
      Top             =   480
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
      Height          =   2412
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optDescription 
         Caption         =   "Description"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1452
      End
      Begin VB.OptionButton optAddrPort 
         Caption         =   "Address and Port"
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
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1452
      End
      Begin VB.OptionButton optTerminal 
         Caption         =   "Terminal "
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
         Height          =   312
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
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
         TabIndex        =   4
         Top             =   2040
         Width           =   1452
      End
   End
   Begin VB.ListBox lstTerminal 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableTerminal.frx":0000
      Left            =   120
      List            =   "frmTableTerminal.frx":0002
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1092
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableTerminal 
      Height          =   1455
      Left            =   2160
      TabIndex        =   19
      Top             =   3480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   1
      Cols            =   4
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
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<==="
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
      Left            =   4800
      TabIndex        =   24
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Terminals variant"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
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
      Top             =   3000
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
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   2760
      Y2              =   4920
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4680
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3480
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   6960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2040
      X2              =   6960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblExpander 
      Alignment       =   2  'Center
      Caption         =   "Expander "
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
      Left            =   4080
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Description "
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
      Left            =   4080
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      Caption         =   " 2-8 Port "
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
      Height          =   495
      Left            =   6480
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      Caption         =   "01-15 Addr. "
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
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblAddrAndPort 
      Alignment       =   2  'Center
      Caption         =   "Addr.  and  Port "
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
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblTerminal 
      Alignment       =   2  'Center
      Caption         =   "Terminal "
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
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblTerminals 
      Alignment       =   2  'Center
      Caption         =   "Terminals "
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmTableTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Текущий номер корректируемой строки "Таблицы терминалов"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Таблицы терминалов"
Dim intColNumCorr As Integer
            '"Старый" номер варианта "Таблицы терминалов"
Dim intVariantOld As Integer
            '"Новый" номер варианта "Таблицы терминалов"
Dim intVariantNew As Integer
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Таблицы терминалов"
Dim gTerminal As TerminalInfo

            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            '"Старый" номер варианта "Таблицы терминалов" не нулевой
    If hsbVariant.Value <> 0 Then
    
            'Были не сохраненные изменения в "Таблице терминалов"
        If gChangesTableTerminal = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы терминалов" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение таблицы терминалов в файле по умолчанию
                cmdSave_Click
            End If
        End If
            
    Else
    
            'Были не сохраненные изменения в "Таблице терминалов"
        If gChangesTableTerminal = True Then
            'Издать звуковой сигнал
            frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы терминалов" - на экран
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
            If strResponse = vbYes Then
            'Сохранение таблицы терминалов в файле по умолчанию
                cmdSave_Click
            End If
        End If
            
            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы терминалов"
        fraColName.Enabled = False
        optTerminal.Enabled = False
        optAddrPort.Enabled = False
        optDescription.Enabled = False
        optExpander.Enabled = False
        lblTerminal.Enabled = False
        txtTerminal.Enabled = False
        lblAddrAndPort.Enabled = False
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblDescription.Enabled = False
        txtDescription.Enabled = False
        lblExpander.Enabled = False
        txtExpander.Enabled = False
        lblTerminals.Enabled = False
        lstTerminal.Enabled = False
    End If
            
            'Очистить текстовые поля
    txtTerminal.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtDescription.Text = ""
    txtExpander.Text = ""
            'Очистить список терминалов
    lstTerminal.Clear
            
            'Сделать нулевым текущий вариант "Таблицы терминалов"
    intVariantOld = 0
    hsbVariant.Value = 0
            'Сбросить признак внесенных изменений в "Таблицу терминалов"
    gChangesTableTerminal = False
            'Сделать невидимой текущую форму
    frmTableTerminal.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()

            ' "Таблица терминалов" не содержит нефиксированных строк
    If grdTableTerminal.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        MsgBox ("The table is empty")
    
    Else
            'Сделать доступными некоторые элементы управления
            '  Коррекцией "Таблицы терминалов"
        fraColName.Enabled = True
        optTerminal.Enabled = True
        optTerminal.Value = True
        optAddrPort.Enabled = True
        optDescription.Enabled = True
        optExpander.Enabled = True
        lblTerminal.Enabled = True
        txtTerminal.Enabled = True
        lstTerminal.Enabled = True
            'Очистить текстовые поля
        txtTerminal.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtDescription.Text = ""
        txtExpander.Text = ""
            'Очистить список имен
        lstTerminal.Clear
    
            'Столбец "Terminal"
        grdTableTerminal.Col = 0
                'Цикл по всем нефиксированным строкам "Таблицы терминалов"
        For intRowNumCorr = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
            grdTableTerminal.Row = intRowNumCorr
            'Заполнение списка "lstName" записями из "Таблицы терминалов"
            lstTerminal.AddItem grdTableTerminal.Text
        Next
            'Выбрать  элемент списка
        lstTerminal.ListIndex = 0
            'Номер корректируемой строки - (1)
        intRowNumCorr = 1
        grdTableTerminal.Row = intRowNumCorr
            'Включить опцию
        optTerminal_Click
    End If
    
End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Таблицы терминалов"
Private Sub grdTableTerminal_Click()
            'Коррекция "включена"
    If lstTerminal.Enabled = True Then
            'Номер корректируемой строки "Таблицы терминалов"
        intRowNumCorr = grdTableTerminal.RowSel
        grdTableTerminal.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstTerminal.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Таблицы терминалов"
        intColNumCorr = grdTableTerminal.ColSel
        grdTableTerminal.Col = intColNumCorr
            'Выбор корректируемой строки "Таблицы терминалов"
        lstTerminal_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstTerminal.Left, Y:=lstTerminal.Top
            'Выбор корректируемого столбца "Таблицы терминалов"
        Select Case intColNumCorr
            Case 1
            optAddrPort.Value = True
            'Установить фокус на текстовом поле для Коррекции
            txtAddress.SetFocus
            Case 2
            optDescription.Value = True
            'Установить фокус на текстовом поле для Коррекции
            txtDescription.SetFocus
            Case 3
            optExpander.Value = True
            'Установить фокус на текстовом поле для Коррекции
            txtExpander.SetFocus
        End Select
    End If
        
End Sub

            'Выбор корректируемой строки "Таблицы терминалов"
Private Sub lstTerminal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемой строки "Таблицы терминалов"
        intRowNumCorr = lstTerminal.ListIndex + 1
        grdTableTerminal.Row = intRowNumCorr
        grdTableTerminal.Col = 0
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Коррекции
        txtTerminal.Text = grdTableTerminal.Text
            'Номер корректируемого столбца "Таблицы терминалов"
        grdTableTerminal.Col = 1
            'Копирование ячейки "Таблицы терминалов" в текстовые поля для Коррекции
        txtAddress.Text = Left(grdTableTerminal.Text, 2)
        txtPort.Text = Mid(grdTableTerminal.Text, 3, 1)
        grdTableTerminal.Col = 2
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Коррекции
        txtDescription.Text = grdTableTerminal.Text
        grdTableTerminal.Col = 3
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Коррекции
        txtExpander.Text = grdTableTerminal.Text
            'Восстановить номер корректируемого столбца "Таблицы терминалов"
        grdTableTerminal.Col = intColNumCorr
    End If

End Sub

            'Выбрана опция - "Terminal"
Private Sub optTerminal_Click()
            'Номер корректируемого столбца "Таблицы терминалов"
    intColNumCorr = 0
    grdTableTerminal.Col = intColNumCorr
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Коррекции
    txtTerminal.Text = grdTableTerminal.Text
            'Сделать (не)доступными некоторые элем. управл. Коррекцией "Таблицы терминалов"
    lblTerminal.Enabled = True
    txtTerminal.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtTerminal.SetFocus
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
            'Номер отображаемого столбца "Таблицы терминалов"
    grdTableTerminal.Col = 1
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Отображения
    txtAddress.Text = Left(Trim(grdTableTerminal.Text), 2)
    lblPort.Enabled = False
    txtPort.Enabled = False
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Отображения
    txtPort.Text = Mid(Trim(grdTableTerminal.Text), 3, 1)
    lblDescription.Enabled = False
    txtDescription.Enabled = False
            'Номер отображаемого столбца "Таблицы терминалов"
    grdTableTerminal.Col = 2
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Отображения
    txtDescription.Text = grdTableTerminal.Text
    lblExpander.Enabled = False
    txtExpander.Enabled = False
            'Номер отображаемого столбца "Таблицы терминалов"
    grdTableTerminal.Col = 3
            'Копирование ячейки "Таблицы терминалов" в текстовое поле для Отображения
    txtExpander.Text = grdTableTerminal.Text
            'Восстановить номер корректируемого столбца "Таблицы терминалов"
    grdTableTerminal.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "AddrPort"
Private Sub optAddrPort_Click()
            'Номер корректируемого столбца "Таблицы терминалов"
    intColNumCorr = 1
    grdTableTerminal.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Таблицы терминалов"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = True
    lblAddress.Enabled = True
    txtAddress.Enabled = True
            'Копирование ячейки "Таблицы терминалов" в текстовые поля для Коррекции
    txtAddress.Text = Left(grdTableTerminal.Text, 2)
    txtPort.Text = Mid(grdTableTerminal.Text, 3, 1)
            'Установить фокус на текстовом поле для Коррекции
    txtAddress.SetFocus
    lblPort.Enabled = True
    txtPort.Enabled = True
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False

End Sub
            
            'Выбрана опция - "Description"
Private Sub optDescription_Click()
            'Номер корректируемого столбца "Таблицы терминалов"
    intColNumCorr = 2
    grdTableTerminal.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Таблицы терминалов"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = True
    txtDescription.Enabled = True
            'Копирование ячейки "Таблицы терминалов" в текстовые поля для Коррекции
    txtDescription.Text = grdTableTerminal.Text
            'Установить фокус на текстовом поле для Коррекции
    txtDescription.SetFocus
    lblExpander.Enabled = False
    txtExpander.Enabled = False

End Sub
            
            'Выбрана опция - "Expander"
Private Sub optExpander_Click()
            'Номер корректируемого столбца "Таблицы терминалов"
    intColNumCorr = 3
    grdTableTerminal.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Таблицы терминалов"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = True
    txtExpander.Enabled = True
            'Копирование ячейки "Таблицы терминалов" в текстовые поля для Коррекции
    txtExpander.Text = grdTableTerminal.Text
            'Установить фокус на текстовом поле для Коррекции
    txtExpander.SetFocus

End Sub
            
            'Процедура ввода и анализа Корректируемого имени "Terminal"
Private Sub txtTerminal_KeyPress(KeyAscii As Integer)
            'Имя введено
    If KeyAscii = vbKeyReturn Then
            'Имя в допустимом диапазоне
        If Len(Trim(txtTerminal.Text)) < 17 Then
            'Изменение имени "Terminal" в "Таблице терминалов"
            grdTableTerminal.Text = Trim(txtTerminal.Text)
            'Установить признак  внесенных изменений в "Таблицу терминалов"
            gChangesTableTerminal = True
            'Включить опцию "optAddrPort"
            optAddrPort.Value = True
            Exit Sub
            'Имя в недопустимом диапазоне
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Address and Port - Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            'Адрес введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo AddressError
            'Адрес в допустимом диапазоне адресов (01/15,  00 - групповой адрес)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text >= 0 _
        And txtAddress.Text < 16 Then
            'Изменение ячейки "Address and Port" в "Таблице терминалов"
            grdTableTerminal.Text = Trim(txtAddress.Text) + Mid(grdTableTerminal.Text, 3)
            'Установить признак  внесенных изменений в "Таблицу терминалов"
            gChangesTableTerminal = True
            'Установить фокус на текстовом поле "Port"
            txtPort.SetFocus
            Exit Sub
            'Адреса в недопустимом диапазоне адресов
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            'Процедура ввода и анализа Корректируемого "Address and Port - Port"
Private Sub txtPort_KeyPress(KeyAscii As Integer)
            'Номер порта введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo PortError
            'Номер порта в допустимом диапазоне (2/8)
        If Len(Trim(txtPort.Text)) = 1 And txtPort.Text > 1 And txtPort.Text < 9 Then
            'Изменение ячейки "Address and Port" в "Таблице терминалов"
            If Len(Trim(grdTableTerminal.Text)) = 0 Then
                grdTableTerminal.Text = "01" + Trim(txtPort.Text)
            Else
                grdTableTerminal.Text = Left(grdTableTerminal.Text, 2) + Trim(txtPort.Text)
            End If
            'Установить признак  внесенных изменений в "Таблицу терминалов"
            gChangesTableTerminal = True
            'Включить опцию "optDescription"
            optDescription.Value = True
            Exit Sub
            'Номер порта в недопустимом диапазоне
PortError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого поля "Description"
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtDescription.Text)) < 17 Then
            'Изменение поля "Expander" в "Таблице терминалов"
            grdTableTerminal.Text = Trim(txtDescription.Text)
            'Установить признак  внесенных изменений в "Таблицу терминалов"
            gChangesTableTerminal = True
            'Включить опцию "optExpander"
            optExpander.Value = True
            'Неверный формат данных
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого поля "Expander"
Private Sub txtExpander_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtExpander.Text)) < 9 Then
            'Изменение поля "Expander" в "Таблице терминалов"
            grdTableTerminal.Text = Trim(txtExpander.Text)
            'Установить признак  внесенных изменений в "Таблицу терминалов"
            gChangesTableTerminal = True
            'Установить фокус на кнопке "Save"
            cmdSave.SetFocus
            'Неверный формат данных
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            'Обработка события "Change" - прокрутка для ползунка "Variant"
Private Sub hsbVariant_Change()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы терминалов"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы терминалов"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы терминалов"
Dim intColNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String

            'Установить ширину столбцов
    SetColWidth
            'Сделать недоступными элементы управления Коррекцией "Таблицы терминалов"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            'Очистить текстовые поля
    txtTerminal.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtDescription.Text = ""
    txtExpander.Text = ""
            'Очистить список интервалов
    lstTerminal.Clear
            
            'Были несохраненные изменения в "Таблице терминалов"
    If gChangesTableTerminal = True Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы терминалов" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранить "Новый" номер варианта "Таблицы терминалов"
            intVariantNew = hsbVariant.Value
            '"Старый" номер варианта "Таблицы терминалов"
            hsbVariant.Value = intVariantOld
            'Сохранение "Таблицы терминалов" в файле по умолчанию
            cmdSave_Click
            'Восстановить "Новый" номер варианта "Таблицы терминалов"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"Старый" номер варианта "Таблицы терминалов"
    intVariantOld = hsbVariant.Value
            'Количество удалений/добавлений строк в "Таблице терминалов"
    gAddDelRowTableTerminal = 0
            'Сбросить признак внесенных изменений в "Таблицу терминалов"
    gChangesTableTerminal = False

            'Заполнение варианта "Таблицы терминалов" из файла
            
            'Вычислить длину записи (строки) "Таблицы терминалов"
    lngRecordLen = Len(gTerminal)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + _
    Trim(Str(hsbVariant.Value)) + ".dat"
                
                                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableTerminal
            'Количество строк "Таблицы терминалов" равно размеру файла по умолчанию +1
    grdTableTerminal.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам  варианта "Таблицы терминалов"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        grdTableTerminal.Row = intRowNum
            'Читать строку "Таблицы терминалов" из файла в буфер
        Get intFileNum, intRowNum, gTerminal
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            grdTableTerminal.Col = intColNum
            'Заполнение текущей строки "Таблицы терминалов" из буфера
            Select Case intColNum
                Case 0
                grdTableTerminal.Text = gTerminal.strTerminal
                Case 1
                 grdTableTerminal.Text = gTerminal.strAddrPort
                Case 2
                grdTableTerminal.Text = gTerminal.strDescription
                Case 3
                grdTableTerminal.Text = gTerminal.strExpander
            End Select
        Next
    Next
            'Закрыть файл
    Close intFileNum
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = hsbVariant.Value
            'Установить фокус на кнопке "Correction"
    If frmTableTerminal.Visible = True Then cmdCorrection.SetFocus
    Exit Sub
ErrorTableTerminal:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableTerminal Error !")
    
End Sub
            
            'Добавление строки в "Таблицу терминалов"
Private Sub cmdAdd_Click()
    
            'Сделать недоступными элементы управления Коррекцией "Таблицы терминалов"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            'Очистить список интервалов
    lstTerminal.Clear
    
            'Формирование номера терминала
    gTerminal.strTerminal = "Terminal-" + Str(grdTableTerminal.Rows)
            'Добавление строки в конец "Таблицы терминалов"
    grdTableTerminal.AddItem gTerminal.strTerminal
            'Количество удалений/добавлений строк в "Таблице терминалов"
    gAddDelRowTableTerminal = gAddDelRowTableTerminal + 1
            'Установить признак внесенных изменений в "Таблицу терминалов"
    gChangesTableTerminal = True
            'Установить фокус на кнопке "Add"
    If frmTableTerminal.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            'Удаление строки из "Таблицы терминалов"
Private Sub cmdDelete_Click()
            'Текущий номер нефиксированной строки "Таблицы терминалов"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы терминалов"
Dim intColNum As Integer
    
            'Сделать недоступными элементы управления Коррекцией "Таблицы терминалов"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            'Очистить список интервалов
    lstTerminal.Clear
    
            'Загрузить (не показывая) форму "frmSelectRow"
    Load frmSelectRow
            'Инициализировать этикетку "lblColName" формы "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Terminal"
    
            'Столбец "Terminal"
    grdTableTerminal.Col = 0
             'Очистить список объектов
    frmSelectRow.lstSelectRow.Clear
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        grdTableTerminal.Row = intRowNum
            'Заполнение списка "lstSelectRow" записями из "Таблицы терминалов"
        frmSelectRow.lstSelectRow.AddItem grdTableTerminal.Text
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
            'Удаление строки из "Таблицы терминалов"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            'Номер удаляемой строки
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            'Удаление строки
        grdTableTerminal.RemoveItem intRowNum
           'Количество удалений/добавлений строк в "Таблице терминалов"
        gAddDelRowTableTerminal = gAddDelRowTableTerminal - 1
            'Установить признак внесенных изменений в "Таблицу терминалов"
        gChangesTableTerminal = True
    End If
            'Выгрузить форму "frmSelectRow"
    UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
    Set frmSelectRow = Nothing
            'Установить фокус на кнопке "Delete"
    If frmTableTerminal.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            'Сохранение "Таблицы терминалов" в файле по умолчанию
Public Function SaveTableTerminal()
    Call cmdSave_Click
    SaveTableTerminal = 0
    
End Function
            
            'Сохранение "Таблицы терминалов" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы терминалов"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы терминалов"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы терминалов"
Dim intColNum As Integer
            'Вычислить длину записи (строки) "Таблицы терминалов"
    lngRecordLen = Len(gTerminal)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + _
    Trim(Str(hsbVariant.Value)) + ".dat"
   
            'Строк, удаленных из "Таблицы терминалов" больше количества
            '  добавленных, ' т.е. умалчиваемый файл станет короче
    If gAddDelRowTableTerminal < 0 Then
            'Удалить "старый" умалчиваемый файл
        Kill strPathFileName
    End If
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        grdTableTerminal.Row = intRowNum
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            grdTableTerminal.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы терминалов" в файл
            Select Case intColNum
                Case 0
                gTerminal.strTerminal = grdTableTerminal.Text
                Case 1
                gTerminal.strAddrPort = grdTableTerminal.Text
                Case 2
                gTerminal.strDescription = grdTableTerminal.Text
                Case 3
                gTerminal.strExpander = grdTableTerminal.Text
            End Select
        Next
            'Записать строку "Таблицы терминалов" в файл
        Put intFileNum, intRowNum, gTerminal
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице терминалов"
    gAddDelRowTableTerminal = 0
            'Сбросить признак внесенных изменений в "Таблицу терминалов"
    gChangesTableTerminal = False
            'Установить фокус на кнопке "Cancel"
    If frmTableTerminal.Visible = True Then cmdCancel.SetFocus
            
End Sub
            
            'Сохранение "Таблицы терминалов" в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы терминалов"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы терминалов"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы терминалов"
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
            'Запись "Таблицы терминалов" в выбранный файл
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Таблицы терминалов"
        lngRecordLen = Len(gTerminal)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Строк, удаленных из "Таблицы терминалов" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
        If gAddDelRowTableTerminal < 0 Then
            'Удалить "старый" файл, если он существует
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
        For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
            grdTableTerminal.Row = intRowNum
            'По всем столбцам "Таблицы терминалов"
            For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
                grdTableTerminal.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы терминалов" в файл
                Select Case intColNum
                Case 0
                gTerminal.strTerminal = grdTableTerminal.Text
                Case 1
                gTerminal.strAddrPort = grdTableTerminal.Text
                Case 2
                gTerminal.strDescription = grdTableTerminal.Text
                Case 3
                gTerminal.strExpander = grdTableTerminal.Text
                End Select
            Next
            'Записать строку "Таблицы терминалов" в файл
        Put intFileNum, intRowNum, gTerminal
        Next
            'Закрыть умалчиваемый файл
        Close intFileNum
            'Количество удалений/добавлений строк в "Таблице терминалов"
        gAddDelRowTableTerminal = 0
            'Сбросить признак внесенных изменений в "Таблицу терминалов"
        gChangesTableTerminal = False
    End If
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Установить фокус на кнопке "Cancel"
    If frmTableTerminal.Visible = True Then cmdCancel.SetFocus
    
End Sub


            'Загрузка формы "Таблица терминалов"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы терминалов"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы терминалов"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы терминалов"
Dim intColNum As Integer
            'Количество столбцов в массиве терминалов доступа
Dim intTerminalNum As Integer

            'Установить ширину столбцов
    SetColWidth
            'Количество вариантов "Таблицы терминалов"
    lblVariant99.Caption = "V" + Str(gVarNumTerminal)
    hsbVariant.Max = gVarNumTerminal
            'Сохранить "Старый" номер варианта "Таблицы терминалов"
    intVariantOld = hsbVariant.Value
    
            'Текущая строка = 0 (Заголовки столбцов)
    grdTableTerminal.Row = 0
    grdTableTerminal.Col = 0
    grdTableTerminal.Text = "Terminal"
            'Записать в ячейку (строка 0, столбец 1)
    grdTableTerminal.Col = 1
    grdTableTerminal.Text = "Address and Port"
            'Записать в ячейку (строка 0, столбец 2)
    grdTableTerminal.Col = 2
    grdTableTerminal.Text = "Description"
            'Записать в ячейку (строка 0, столбец 3)
    grdTableTerminal.Col = 3
    grdTableTerminal.Text = "Expander"
            
            'Заполнение "Таблицы терминалов" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы терминалов"
    lngRecordLen = Len(gTerminal)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(hsbVariant.Value)) + ".dat"
                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableTerminal
                'Количество строк "Таблицы терминалов" равно размеру файла по умолчанию +1
    grdTableTerminal.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы терминалов"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            'Текущая строка "Таблицы терминалов"
        grdTableTerminal.Row = intRowNum
            'Читать строку "Таблицы терминалов" из файла в буфер
        Get intFileNum, intRowNum, gTerminal
            'По всем столбцам "Таблицы терминалов"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            'Текущий столбец "Таблицы терминалов"
            grdTableTerminal.Col = intColNum
            'Заполнение текущей строки "Таблицы терминалов" из буфера
            Select Case intColNum
                Case 0
                grdTableTerminal.Text = gTerminal.strTerminal
                Case 1
                grdTableTerminal.Text = gTerminal.strAddrPort
                Case 2
                grdTableTerminal.Text = gTerminal.strDescription
                Case 3
                grdTableTerminal.Text = gTerminal.strExpander
            End Select
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            
            'Максимальное количество столбцов в массивe терминалов доступа
    intTerminalNum = grdTableTerminal.Rows
            'Цикл по всем ненулевым вариантам "Таблицы терминалов"
    For intVariantNew = 1 To gVarNumTerminal Step 1
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(intVariantNew)) + ".dat"
    
            'Максимальное количество солбцов в массиве терминалов доступа
        If FileLen(strPathFileName) / lngRecordLen + 1 > intTerminalNum Then
            intTerminalNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    Next
            'Переопределить размерность массива терминалов доступа
ReDim gAddrPort(gVarNumTerminal + 1, intTerminalNum) As String * 4
    
            'Цикл по всем вариантам "Таблицы терминалов"
    For intVariantNew = 0 To gVarNumTerminal Step 1
            'Получить свободный номер файла
        intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(intVariantNew)) + ".dat"
    
            'Количество "значащих" столбцов в текущей строке массива терминалов доступа
        intTerminalNum = FileLen(strPathFileName) / lngRecordLen + 1
        gAddrPort(intVariantNew, 0) = intTerminalNum
    
            'Открыть файл для произвольного доступа
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам  варианта "Таблицы терминалов"
        For intRowNum = 1 To intTerminalNum - 1 Step 1
            'Читать строку "Таблицы терминалов" из файла в буфер
            Get intFileNum, intRowNum, gTerminal
                 gAddrPort(intVariantNew, intRowNum) = Trim(gTerminal.strAddrPort) + "0"
            'Присутствует Имя (признак) Препроцессора в столбце "Expander"
            '   "Таблицы терминалов"
            If gTerminal.strExpander = gPreprocName Then
            'Заменить номер порта в текущей строке таблицы терминалов на
            '   первую букву Имени Препроцессора
                gAddrPort(intVariantNew, intRowNum) = _
                Left(gAddrPort(intVariantNew, intRowNum), 2) + _
                Left(gPreprocName, 1) + "0"
            End If
        Next
            'Закрыть файл
        Close intFileNum
    Next
            
            'Количество удалений/добавлений строк в "Таблице терминалов"
    gAddDelRowTableTerminal = 0
            'Сбросить признак внесенных изменений в "Таблицу терминалов"
    gChangesTableTerminal = False
            'Индицировать номер варианта в текстовом поле "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableTerminal:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableTerminal Error !")
    
End Sub
            
            'Процедура восстановления умалчиваемых значений
            '' параметров всех вариантов "Таблицы терминалов"
Private Sub cmdDefaultTerm_Click()
            
            '"Старый" номер варианта "Таблицы терминалов" не нулевой
    If hsbVariant.Value <> 0 Then
            'Сделать нулевым текущий вариант "Таблицы терминалов"
        intVariantOld = 0
        hsbVariant.Value = 0
    Else
    
            'Сделать недоступными элементы управления
            '  Коррекцией "Таблицы терминалов"
        fraColName.Enabled = False
        optTerminal.Enabled = False
        optAddrPort.Enabled = False
        optDescription.Enabled = False
        optExpander.Enabled = False
        lblTerminal.Enabled = False
        txtTerminal.Enabled = False
        lblAddrAndPort.Enabled = False
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblDescription.Enabled = False
        txtDescription.Enabled = False
        lblExpander.Enabled = False
        txtExpander.Enabled = False
        lblTerminals.Enabled = False
        lstTerminal.Enabled = False
            'Очистить текстовые поля
        txtTerminal.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtDescription.Text = ""
        txtExpander.Text = ""
            'Очистить список интервалов
        lstTerminal.Clear
    End If
            
            'Загрузка формы "Таблица терминалов"
    Form_Load
            'Установить фокус на кнопке "Correction"
    If frmTableTerminal.Visible = True Then cmdCorrection.SetFocus
            
End Sub
            
            'Процедура установки ширины и выравнивания столбцов "Таблицы терминалов"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Цикл по всем столбцам
    For intColNumber = 0 To grdTableTerminal.Cols - 1 Step 1
        grdTableTerminal.ColWidth(intColNumber) = 1480
        grdTableTerminal.ColAlignment(intColNumber) = 0
    Next
    
End Sub
