VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableSystem 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_system"
   ClientHeight    =   6810
   ClientLeft      =   1035
   ClientTop       =   1095
   ClientWidth     =   9060
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
   ScaleHeight     =   6810
   ScaleWidth      =   9060
   Begin VB.OptionButton optReset 
      Caption         =   "Reset"
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
      Left            =   6000
      TabIndex        =   32
      Top             =   2520
      Value           =   -1  'True
      Width           =   732
   End
   Begin VB.ListBox lstIndex 
      Enabled         =   0   'False
      Height          =   900
      ItemData        =   "frmTableSystem.frx":0000
      Left            =   6840
      List            =   "frmTableSystem.frx":0002
      TabIndex        =   31
      Top             =   2400
      Width           =   1815
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
      Height          =   3252
      Left            =   2040
      TabIndex        =   19
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optIndex 
         Caption         =   "Index"
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
         TabIndex        =   29
         Top             =   2280
         Width           =   1452
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Object "
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
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.OptionButton optConsAddrTerm 
         Caption         =   "Constant or Address and Terminal"
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
         Height          =   612
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1452
      End
      Begin VB.OptionButton optType 
         Caption         =   "Type"
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
         TabIndex        =   21
         Top             =   1560
         Width           =   1452
      End
      Begin VB.OptionButton optAppendix 
         Caption         =   "Appendix"
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
         TabIndex        =   20
         Top             =   2880
         Width           =   1452
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
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
      Height          =   612
      Left            =   3960
      TabIndex        =   14
      Top             =   1560
      Width           =   4812
      Begin VB.OptionButton optConstant 
         Caption         =   "Constant"
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
         Top             =   240
         Value           =   -1  'True
         Width           =   1092
      End
      Begin VB.OptionButton optReader 
         Caption         =   "Reader"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   852
      End
      Begin VB.OptionButton optProcessor 
         Caption         =   "Processor"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optWriter 
         Caption         =   "Writer"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtObject 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5160
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtConstant 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5160
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtAddress 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7680
      TabIndex        =   11
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox txtAppendix 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5280
      TabIndex        =   10
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox txtTerm 
      Enabled         =   0   'False
      Height          =   288
      Left            =   8280
      TabIndex        =   9
      Top             =   960
      Width           =   252
   End
   Begin VB.ListBox lstObject 
      Enabled         =   0   'False
      Height          =   1950
      ItemData        =   "frmTableSystem.frx":0004
      Left            =   120
      List            =   "frmTableSystem.frx":0006
      TabIndex        =   7
      Top             =   4080
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
      TabIndex        =   6
      Top             =   1440
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
      Top             =   6240
      Width           =   1212
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
      Left            =   6600
      TabIndex        =   3
      Top             =   6240
      Width           =   1092
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
      Left            =   7800
      TabIndex        =   2
      Top             =   6240
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
      TabIndex        =   1
      Top             =   6240
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
      TabIndex        =   0
      Top             =   6240
      Width           =   1092
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableSystem 
      Height          =   2295
      Left            =   2160
      TabIndex        =   5
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   9
      Cols            =   5
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
   Begin VB.Line Line19 
      X1              =   3960
      X2              =   3960
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line Line18 
      X1              =   3960
      X2              =   3960
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   3960
      X2              =   6720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   6720
      X2              =   8760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   8760
      X2              =   8760
      Y1              =   3480
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   3960
      X2              =   8760
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   8880
      X2              =   8880
      Y1              =   3720
      Y2              =   120
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   120
      X2              =   8880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   8880
      X2              =   2040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   6120
      Y2              =   3720
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   5160
      X2              =   7440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblIndex 
      Alignment       =   2  'Center
      Caption         =   "Index "
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
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblObject 
      Alignment       =   2  'Center
      Caption         =   "Object "
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
      TabIndex        =   28
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblConstant 
      Alignment       =   2  'Center
      Caption         =   "Constant "
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
      TabIndex        =   27
      Top             =   960
      Width           =   975
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
      Left            =   7560
      TabIndex        =   26
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   7440
      X2              =   7440
      Y1              =   1440
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8760
      X2              =   8760
      Y1              =   1560
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   7440
      X2              =   8760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblAppendix 
      Alignment       =   2  'Center
      Caption         =   "Appendix "
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
      TabIndex        =   25
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblTerm 
      Alignment       =   2  'Center
      Caption         =   " 0-3 Term. "
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
      Left            =   8160
      TabIndex        =   24
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblObjects 
      Alignment       =   2  'Center
      Caption         =   "Objects "
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
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   6120
      Y2              =   6120
   End
End
Attribute VB_Name = "frmTableSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Текущий номер корректируемой строки "Системной таблицы"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Системной таблицы"
Dim intColNumCorr As Integer
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Системной таблицы"
Dim gSystem As SystemInfo

            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            'Были не сохраненные изменения в "Системной таблице"
    If gChangesTableSystem = True Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Системной таблицы" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранение "Системной таблицы" в файле по умолчанию
            cmdSave_Click
        End If
    End If
    
            'Сделать недоступными элементы управления Коррекцией "Системной таблицы"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            'Очистить текстовые поля
    txtObject.Text = ""
    txtConstant.Text = ""
    txtAddress.Text = ""
    txtTerm.Text = ""
    txtAppendix.Text = ""
            'Очистить список объектов
    lstObject.Clear
    lstIndex.Clear
    
            'Сбросить признак внесенных изменений в "Системную таблицу"
    gChangesTableSystem = False
            'Сделать невидимой текущую форму
    frmTableSystem.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()

            ' "Системная таблица" не содержит нефиксированных строк
    If grdTableSystem.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        MsgBox ("The table is empty")
    
    Else
            'Сделать доступными некоторые элем. управления Коррекцией "Системной таблицы"
        fraColName.Enabled = True
        optObject.Enabled = True
        optObject.Value = True
        optConsAddrTerm.Enabled = True
        optType.Enabled = True
        optIndex.Enabled = True
        optAppendix.Enabled = True
        lblObject.Enabled = True
        txtObject.Enabled = True
        lblObjects.Enabled = True
        lstObject.Enabled = True
            'Очистить текстовые поля
        txtObject.Text = ""
        txtConstant.Text = ""
        txtAddress.Text = ""
        txtTerm.Text = ""
        txtAppendix.Text = ""
            'Очистить списки
        lstObject.Clear
        lstIndex.Clear
    
            'Столбец "Objects"
        grdTableSystem.Col = 0
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNumCorr = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
            grdTableSystem.Row = intRowNumCorr
            'Заполнение списка "lstObject" записями из "Системной таблицы"
            lstObject.AddItem grdTableSystem.Text
            'Заполнение списка "lstIndex" записями из "Системной таблицы"
            lstIndex.AddItem grdTableSystem.Text
        Next
            'Выбрать  элемент списка
        lstObject.ListIndex = 0
        lstIndex.ListIndex = 0
            'Номер корректируемой строки - (1)
        intRowNumCorr = 1
        grdTableSystem.Row = intRowNumCorr
            'Включить опцию
        optObject_Click
    
    End If
    
End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Системной таблицы"
Private Sub grdTableSystem_Click()
            'Коррекция "включена"
    If lstObject.Enabled = True Then
            'Номер корректируемой строки "Системной таблицы"
        intRowNumCorr = grdTableSystem.RowSel
        grdTableSystem.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstObject.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Системной таблицы"
        intColNumCorr = grdTableSystem.ColSel
        grdTableSystem.Col = intColNumCorr
            'Выбор корректируемой строки "Системной таблицы"
        lstObject_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstObject.Left, Y:=lstObject.Top
            'Выбор корректируемого столбца "Системной таблицы"
        Select Case intColNumCorr
            Case 1
            If optConsAddrTerm.Value = True Then
                optConsAddrTerm_Click
            Else
                optConsAddrTerm.Value = True
            End If
            Case 2
            optType.Value = True
            Case 3
            optIndex.Value = True
            Case 4
            optAppendix.Value = True
        End Select
    End If
        
End Sub

            'Выбор корректируемой строки "Системной таблицы"
Private Sub lstObject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемой строки "Системной таблицы"
        intRowNumCorr = lstObject.ListIndex + 1
        grdTableSystem.Row = intRowNumCorr
        grdTableSystem.Col = 0
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
        txtObject.Text = grdTableSystem.Text
            'Номер анализируемого столбца "Системной таблицы" - "Type"
        grdTableSystem.Col = 2
            'Установлена опция "Constant"
        If Left(grdTableSystem.Text, 2) = "00" Then
            'Номер корректируемого столбца "Системной таблицы"
            grdTableSystem.Col = 1
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
            txtConstant.Text = grdTableSystem.Text
            'Очиска недоступных для Коррекции текстовых полей
            txtAddress.Text = ""
            txtTerm.Text = ""
            'Не установлена опция "Constant"
        Else
            'Номер корректируемого столбца "Системной таблицы"
            grdTableSystem.Col = 1
            'Копирование ячейки "Системной таблицы" в текстовые поля для Коррекции
            txtAddress.Text = Left(grdTableSystem.Text, 2)
            txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
            'Очиска недоступного для Коррекции текстового поля
            txtConstant.Text = ""
        End If
        grdTableSystem.Col = 3
            'Номер строки "Системной таблицы", на которую указывает ячейка "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            lstIndex.ListIndex = grdTableSystem.Text - 1
            'Ячейка "Index" пустая
        Else
            lstIndex.ListIndex = 0
        End If
        grdTableSystem.Col = 4
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
        txtAppendix.Text = grdTableSystem.Text
            'Восстановить номер корректируемого столбца"Системной таблицы"
        grdTableSystem.Col = intColNumCorr
    End If
    
End Sub

            'Выбрана опция - "Object"
Private Sub optObject_Click()
            'Номер корректируемого столбца "Системной таблицы"
    intColNumCorr = 0
    grdTableSystem.Col = intColNumCorr
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
    txtObject.Text = grdTableSystem.Text
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Системной таблицы"
    lblObject.Enabled = True
    txtObject.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtObject.SetFocus
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
            'Номер анализируемого столбца "Системной таблицы" - "Type"
    grdTableSystem.Col = 2
            'Установлена опция "Constant"
    If Left(grdTableSystem.Text, 2) = "00" Then
            'Номер отображаемого столбца "Системной таблицы"
        grdTableSystem.Col = 1
            'Копирование ячейки "Системной таблицы" в текстовое поле для Отображения
        txtConstant.Text = grdTableSystem.Text
            'Очистка недоступных для Отображения текстовых полей
        txtAddress.Text = ""
        txtTerm.Text = ""
            'Не установлена опция "Constant"
    Else
            'Номер отображаемого столбца "Системной таблицы"
        grdTableSystem.Col = 1
            'Копирование ячейки "Системной таблицы" в текстовое поле для Отображения
        txtConstant.Text = grdTableSystem.Text
            'Копирование ячейки "Системной таблицы" в текстовые поля для Отображения
        txtAddress.Text = Left(grdTableSystem.Text, 2)
        txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
            'Очистка недоступного для Отображения текстового поля
        txtConstant.Text = ""
    End If
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            'Номер отображаемого столбца "Системной таблицы"
    grdTableSystem.Col = 4
            'Копирование ячейки "Системной таблицы" в текстовое поле для Отображения
    txtAppendix.Text = grdTableSystem.Text
    
            'Восстановить номер корректируемого столбца "Системной таблицы"
    grdTableSystem.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "ConsAddrTerm"
Private Sub optConsAddrTerm_Click()
            'Номер корректируемого столбца "Системной таблицы"
    intColNumCorr = 1
    grdTableSystem.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Системной таблицы"
    lblObject.Enabled = False
    txtObject.Enabled = False
            'Номер анализируемого столбца "Системной таблицы" - "Type"
    grdTableSystem.Col = 2
            'Установлена опция "Constant"
    If Left(grdTableSystem.Text, 2) = "00" Then
            'Номер корректируемого столбца "Системной таблицы"
        grdTableSystem.Col = intColNumCorr
        lblConstant.Enabled = True
        txtConstant.Enabled = True
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblTerm.Enabled = False
        txtTerm.Enabled = False
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
        txtConstant.Text = grdTableSystem.Text
                'Установить фокус на текстовом поле для Коррекции
        txtConstant.SetFocus
            'Очиска недоступных для Коррекции текстовых полей
        txtAddress.Text = ""
        txtTerm.Text = ""
            'Не установлена опция "Constant"
    Else
            'Номер корректируемого столбца "Системной таблицы"
        grdTableSystem.Col = intColNumCorr
        lblConstant.Enabled = False
        txtConstant.Enabled = False
        lblAddress.Enabled = True
        txtAddress.Enabled = True
        lblTerm.Enabled = True
        txtTerm.Enabled = True
            'Копирование ячейки "Системной таблицы" в текстовые поля для Коррекции
        txtAddress.Text = Left(grdTableSystem.Text, 2)
        txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
                'Установить фокус на текстовом поле для Коррекции
        txtAddress.SetFocus
            'Очиска недоступного для Коррекции текстового поля
        txtConstant.Text = ""
    End If
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub

            'Выбрана опция - "Type"
Private Sub optType_Click()
            'Номер корректируемого столбца "Системной таблицы"
    intColNumCorr = 2
    grdTableSystem.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Системной таблицы"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = True
    optConstant.Enabled = True
            'Включить опцию "Constant"
    optConstant.Value = True
    optReader.Enabled = True
    optWriter.Enabled = True
    optProcessor.Enabled = True
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub
            
            'Выбрана опция - "Index"
Private Sub optIndex_Click()
            'Номер корректируемого столбца "Системной таблицы"
    intColNumCorr = 3
    grdTableSystem.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Системной таблицы"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = True
            'Номер строки "Системной таблицы", на которую указывает ячейка "Index"
    If Trim(grdTableSystem.Text) <> "" Then
        lstIndex.ListIndex = grdTableSystem.Text - 1
            'Ячейка "Index" пустая
    Else
        lstIndex.ListIndex = 0
    End If
    optReset.Enabled = True
    lstIndex.Enabled = True
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub

            'Выбрана опция "Appendix"
Private Sub optAppendix_Click()
            'Номер корректируемого столбца "Системной таблицы"
    intColNumCorr = 4
    grdTableSystem.Col = intColNumCorr
            'Сделать (не)доступными некоторые элем-ы управл. Коррекцией "Системной таблицы"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
            'Копирование ячейки "Системной таблицы" в текстовое поле для Коррекции
    txtAppendix.Text = grdTableSystem.Text
    lblAppendix.Enabled = True
    txtAppendix.Enabled = True
                'Установить фокус на текстовом поле для Коррекции
    txtAppendix.SetFocus

End Sub
            
            'Процедура ввода и анализа Корректируемого имени "Object"
Private Sub txtObject_KeyPress(KeyAscii As Integer)
            'Имя введено
    If KeyAscii = vbKeyReturn Then
            'Имя в допустимом диапазоне
        If Len(Trim(txtObject.Text)) < 17 Then
            'Изменение имени "Object" в "Системной таблице"
          grdTableSystem.Text = Trim(txtObject.Text)
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Включить опцию "optConsAddrTerm"
            optConsAddrTerm.Value = True
            'Имя в недопустимом диапазоне
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Constant"
Private Sub txtConstant_KeyPress(KeyAscii As Integer)
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Константа в допустимом диапазоне
        If Len(Trim(txtConstant.Text)) < 17 Then
            'Изменение ячейки "Cons.,Addr.,Term." в "Системной таблице"
            grdTableSystem.Text = Trim(txtConstant.Text)
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Включить опцию "optType"
            optType.Value = True
            'Константа в недопустимом диапазоне
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            'Адрес введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo AddressError
            'Адрес в допустимом диапазоне адресов (01/15,  00 - групповой адрес)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text > 0 And txtAddress.Text < 16 Then
            'Изменение ячейки "Cons.,Addr.,Term." в "Системной таблице"
            If Len(Trim(grdTableSystem.Text)) < 3 Then
            'Изменение ячейки "Cons.,Addr.,Term." в "Системной таблице"
                txtTerm.Text = "0"
                grdTableSystem.Text = Trim(txtAddress.Text) + Trim(txtTerm.Text)
            Else
                grdTableSystem.Text = Trim(txtAddress.Text) + Mid(grdTableSystem.Text, 3)
            End If
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Установить фокус на текстовое поле "txtTerm"
            txtTerm.SetFocus
            Exit Sub
            'Адреса в недопустимом диапазоне адресов
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            'Процедура ввода и анализа Корректируемого "Term"
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
            'Номер терминала введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo TermError
            'Номер терминала в допустимом диапазоне (0/3)
        If Len(Trim(txtTerm.Text)) = 1 And txtTerm.Text >= 0 And txtTerm.Text < 4 Then
            'Изменение ячейки "Cons.,Addr.,Term." в "Системной таблице"
            
            If Len(Trim(grdTableSystem.Text)) < 3 Then
            'Изменение ячейки "Cons.,Addr.,Term." в "Системной таблице"
                txtAddress.Text = "01"
                grdTableSystem.Text = Trim(txtAddress.Text) + Trim(txtTerm.Text)
            Else
                grdTableSystem.Text = Left(grdTableSystem.Text, 2) + Trim(txtTerm.Text) + _
                Mid(grdTableSystem.Text, 4)
            End If
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Включить опцию "optType"
            optType.Value = True
            Exit Sub
            'Номер терминала в недопустимом диапазоне
TermError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            'Выбрана опция - "Constant"
Private Sub optConstant_GotFocus()
            'Изменение ячейки "Type" в "Системной таблице"
    grdTableSystem.Text = "00 - Constant"
            'Установить признак  внесенных изменений в "Системную таблицу"
    gChangesTableSystem = True

End Sub

            'Выбрана опция - "Reader"
Private Sub optReader_GotFocus()
            'Изменение ячейки "Type" в "Системной таблице"
    grdTableSystem.Text = "01 - Reader"
            'Установить признак  внесенных изменений в "Системную таблицу"
    gChangesTableSystem = True

End Sub

            'Выбрана опция - "Writer"
Private Sub optWriter_GotFocus()
            'Изменение ячейки "Type" в "Системной таблице"
    grdTableSystem.Text = "02 - Writer"
            'Установить признак  внесенных изменений в "Системную таблицу"
    gChangesTableSystem = True

End Sub

            'Выбрана опция - "Processor"
Private Sub optProcessor_GotFocus()
            'Изменение ячейки "Type" в "Системной таблице"
    grdTableSystem.Text = "03 - Processor"
            'Установить признак  внесенных изменений в "Системную таблицу"
    gChangesTableSystem = True

End Sub

            'Выбрана опция - "Reset"
Private Sub optReset_GotFocus()
            'Не пустая текущая ячейка "Index"
    If Trim(grdTableSystem.Text) <> "" Then
            'Номер строки "Системной таблицы", на которую ссылается текущая ячейка "Index"
        grdTableSystem.Row = grdTableSystem.Text
            'Строка "Системной таблицы", на которую ссылается текущая ячейка "Index"
            '  тоже содержит не пустую ячейку "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            'Связывание возикающего "разрыва" в списковой структуре
            lstIndex.ListIndex = grdTableSystem.Text - 1
            grdTableSystem.Row = intRowNumCorr
            'Коррекция содержимого текущей ячейки "Index" в "Системной таблице"
            grdTableSystem.Text = lstIndex.ListIndex + 1
        Else
            'Очистка текущей ячейки "Index" в "Системной таблице"
            grdTableSystem.Row = intRowNumCorr
            grdTableSystem.Text = ""
            lstIndex.ListIndex = 0
        End If
            'Установить признак  внесенных изменений в "Системную таблицу"
        gChangesTableSystem = True
    End If

End Sub

            'Выбор строки "Системной таблицы", на которую будет указывать ячейка "Index"
Private Sub lstIndex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = 1 Then
            'Верный номер ссылки - индекс не указывает сам на себя
        If lstIndex.ListIndex + 1 <> grdTableSystem.Row Then
            'Изменение ячейки "Index" в "Системной таблице"
            grdTableSystem.Text = lstIndex.ListIndex + 1
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Неверный формат данных
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого поля "Appendix"
Private Sub txtAppendix_KeyPress(KeyAscii As Integer)
            'Информация введена
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtAppendix.Text)) < 9 Then
            'Изменение поля "Appendix" в "Системной таблице"
            grdTableSystem.Text = Trim(txtAppendix.Text)
            'Установить признак  внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Установить фокус на кнопке "Save"
            cmdSave.SetFocus
            'Неверный формат данных
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Добавление строки в "Системную таблицу"
Private Sub cmdAdd_Click()
            'Пустая строка "Системной таблицы"
Dim strSystem As String
    strSystem = ""
    
            'Сделать недоступными элементы управления Коррекцией "Системной таблицы"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            'Очистить список объектов
    lstObject.Clear
    lstIndex.Clear
    
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя Имя объекта
    strSystem = InputBox("Objects Name:", "Add ...")
            'Имя объекта не выбрано
    If strSystem = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox " The object isn't selected !"
            'Выбрана таблица для просмотра
    Else
            'Добавление строки в конец "Системной таблицы"
        grdTableSystem.AddItem strSystem
            'Количество удалений/добавлений строк в "Системной таблице"
        gAddDelRowTableSystem = gAddDelRowTableSystem + 1
            'Установить признак внесенных изменений в "Системную таблицу"
        gChangesTableSystem = True
    End If
            'Установить фокус на кнопке "Add"
    cmdAdd.SetFocus
    
End Sub
            
            'Удаление строки из "Системной таблицы"
Private Sub cmdDelete_Click()
            'Текущий номер нефиксированной строки "Системной таблицы"
Dim intRowNum As Integer
Dim intRowNumSys As Integer
            'Текущий номер столбца "Системной таблицы"
Dim intColNum As Integer

            'Сделать недоступными элементы управления Коррекцией "Системной таблицы"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            'Очистить список объектов
    lstObject.Clear
    lstIndex.Clear
            
            'Загрузить (не показывая) форму "frmSelectRow"
    Load frmSelectRow
            'Инициализировать этикетку "lblColName" формы "frmSelectRow"
    frmSelectRow.lblColName.Caption = "System"
    
            'Столбец "Objects"
    grdTableSystem.Col = 0
             'Очистить список объектов
    frmSelectRow.lstSelectRow.Clear
           'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        grdTableSystem.Row = intRowNum
            'Заполнение списка "lstSelectRow" записями из "Системной таблицы"
        frmSelectRow.lstSelectRow.AddItem grdTableSystem.Text
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
            'Удаление строки из "Системной таблицы"
            'Последняя строка не может быть удалена
    ElseIf frmSelectRow.lstSelectRow.ListCount = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The last row isn't selected !"
            'Удаление строки из "Таблицы времени"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            'Номер удаляемой строки
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            
            'Столбец "Index"
        grdTableSystem.Col = 3
            'Не пустая текущая ячейка "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            'Издать звуковой сигнал
           frmDemo.BeepSound
            'Удаление невозможно
            MsgBox "Deletion impossible. The Index isn't empty !"
        Else
            'Удаление строки
            grdTableSystem.RemoveItem intRowNum
            'Количество удалений/добавлений строк в "Системной таблице"
            gAddDelRowTableSystem = gAddDelRowTableSystem - 1
            'Установить признак внесенных изменений в "Системную таблицу"
            gChangesTableSystem = True
            'Цикл по всем нефиксир. строкам "Системной таблицы" - коррекция ячеек "Index"
            For intRowNumSys = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
                grdTableSystem.Row = intRowNumSys
            'Строка содержит указатель
                If Trim(grdTableSystem.Text) <> "" Then
            'Очистка ячейки "Index", которая указывает на удаленную строку "Системной таблицы"
                    If grdTableSystem.Text = intRowNum Then grdTableSystem.Text = ""
            'Уменьшить на -1 содержимое непустых ячеек "Index", которые указывают на номера
            '  строк "Системной таблицы" большие, чем номер удаленной строки
                    If Trim(grdTableSystem.Text) > intRowNum _
                    Then grdTableSystem.Text = grdTableSystem.Text - 1
                End If
            Next
        End If
            
    End If
            'Выгрузить форму "frmSelectRow"
    UnLoad frmSelectRow
            'Освободить память, занимаемую выгруженной формой
    Set frmSelectRow = Nothing
            'Установить фокус на кнопке "Delete"
    cmdDelete.SetFocus
    
End Sub
            
            'Сохранение "Системной таблицы" в файле по умолчанию
Public Function SaveTableSystem()
    Call cmdSave_Click
    SaveTableSystem = 0
    
End Function
            
            'Сохранение "Системной таблицы" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Системной таблицы"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Системной таблицы"
Dim intRowNum As Integer
            'Текущий номер столбца "Системной таблицы"
Dim intColNum As Integer
            'Вычислить длину записи (строки) "Системной таблицы"
    lngRecordLen = Len(gSystem)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableSystem.dat"
    
            'Строк, удаленных из "Системной таблицы" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
    If gAddDelRowTableSystem < 0 Then
            'Удалить "старый" умалчиваемый файл
        Kill strPathFileName
    End If
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        grdTableSystem.Row = intRowNum
            'По всем столбцам "Системной таблицы"
        For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            grdTableSystem.Col = intColNum
            'Заполнение буфера для записи текущей строки "Системной таблицы" в файл
            Select Case intColNum
                Case 0
                gSystem.strObject = grdTableSystem.Text
                Case 1
                gSystem.strConsAddrTerm = grdTableSystem.Text
                Case 2
                gSystem.strType = Left(grdTableSystem.Text, 2)
                Case 3
                gSystem.strIndex = Left(grdTableSystem.Text, 5)
                Case 4
                gSystem.strAppendix = grdTableSystem.Text
            End Select
        Next
            'Записать строку "Системной таблицы" в файл
        Put intFileNum, intRowNum, gSystem
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Системной таблицы"
    gAddDelRowTableSystem = 0
            'Сбросить признак внесенных изменений в "Системной таблицы"
    gChangesTableSystem = False
            
End Sub
            
            'Сохранение "Системной таблицы" в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Системной таблицы"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Системной таблицы"
Dim intRowNum As Integer
            'Текущий номер столбца "Системной таблицы"
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
            'Запись "Системной таблицы" в выбранный файл
    Else
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Системной таблицы"
        lngRecordLen = Len(gSystem)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Строк, удаленных из "Системной таблицы" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
        If gAddDelRowTableSystem < 0 Then
            'Удалить "старый" файл, если он существует
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            'Открыть выбранный файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
            grdTableSystem.Row = intRowNum
            'По всем столбцам "Системной таблицы"
            For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
                grdTableSystem.Col = intColNum
            'Заполнение буфера для записи текущей строки "Системной таблицы" в файл
                Select Case intColNum
                    Case 0
                    gSystem.strObject = grdTableSystem.Text
                    Case 1
                    gSystem.strConsAddrTerm = grdTableSystem.Text
                    Case 2
                    gSystem.strType = Left(grdTableSystem.Text, 2)
                    Case 3
                    gSystem.strIndex = Left(grdTableSystem.Text, 5)
                    Case 4
                    gSystem.strAppendix = grdTableSystem.Text
                End Select
            Next
            'Записать строку "Системной таблицы" в файл
            Put intFileNum, intRowNum, gSystem
        Next
            'Закрыть выбранный файл
        Close intFileNum
             'Количество удалений/добавлений строк в "Системной таблице"
        gAddDelRowTableSystem = 0
               'Сбросить признак внесенных изменений в "Системной таблицы"
        gChangesTableSystem = False
    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Установить фокус на кнопке "Cancel"
    cmdCancel.SetFocus
    
End Sub

            'Загрузка формы "Системной таблицы"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Системной таблицы"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Системной таблицы"
Dim intRowNum As Integer
            'Текущий номер столбца "Системной таблицы"
Dim intColNum As Integer

            'Установить ширину столбцов
    SetColWidth
            'Текущая строка = 0 (Заголовки столбцов)
    grdTableSystem.Row = 0
    grdTableSystem.Col = 0
    grdTableSystem.Text = "Objects"
            'Записать в ячейку (строка 0, столбец 1)
    grdTableSystem.Col = 1
    grdTableSystem.Text = "Cons.,Addr.,Term."
            'Записать в ячейку (строка 0, столбец 2)
    grdTableSystem.Col = 2
    grdTableSystem.Text = "Type"
            'Записать в ячейку (строка 0, столбец 3)
    grdTableSystem.Col = 3
    grdTableSystem.Text = "Index"
            'Записать в ячейку (строка 0, столбец 4)
    grdTableSystem.Col = 4
    grdTableSystem.Text = "Appendix"
    
            
            'Заполнение "Системной таблицы" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Системной таблицы"
    lngRecordLen = Len(gSystem)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableSystem.dat"
                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableSystem
                'Количество строк "Системной таблицы" равно размеру файла по умолчанию +1
    grdTableSystem.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        grdTableSystem.Row = intRowNum
            'Читать строку "Системной таблицы" из файла в буфер
        Get intFileNum, intRowNum, gSystem
            'По всем столбцам "Системной таблицы"
        For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            'Текущий столбец "Системной таблицы"
            grdTableSystem.Col = intColNum
            'Заполнение текущей строки "Системной таблицы" из буфера
            Select Case intColNum
                Case 0
                grdTableSystem.Text = gSystem.strObject
                Case 1
                grdTableSystem.Text = gSystem.strConsAddrTerm
                Case 2
                grdTableSystem.Text = gSystem.strType
                If gSystem.strType = "00" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Constant"
                If gSystem.strType = "01" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Reader"
                If gSystem.strType = "02" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Writer"
                If gSystem.strType = "03" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Processor"
                Case 3
                grdTableSystem.Text = Left(gSystem.strIndex, 5)
                Case 4
                grdTableSystem.Text = gSystem.strAppendix
            End Select
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Системной таблице"
    gAddDelRowTableSystem = 0
            'Сбросить признак внесенных изменений в "Системной таблицы"
    gChangesTableSystem = False
    
    Exit Sub
ErrorTableSystem:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableSystem Error !")
    
End Sub
            
            'Процедура установки ширины и выравнивания столбцов "Системной таблицы"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Цикл по всем столбцам
    For intColNumber = 0 To grdTableSystem.Cols - 1 Step 1
        grdTableSystem.ColWidth(intColNumber) = 1600
        grdTableSystem.ColAlignment(intColNumber) = 0
    Next
    
End Sub
