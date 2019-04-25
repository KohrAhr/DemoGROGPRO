VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_information"
   ClientHeight    =   7125
   ClientLeft      =   1185
   ClientTop       =   1080
   ClientWidth     =   9510
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
   ScaleHeight     =   7125
   ScaleWidth      =   9510
   Begin VB.TextBox txtCategory 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   33
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtSite 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   32
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtBrigade 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   31
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtDepartment 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   30
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtTNumber 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   29
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtSurName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   28
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   27
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CommandButton cmdFormTableInfoFromTablePerson 
      Cancel          =   -1  'True
      Caption         =   "F&Orm 'TableInfo' From 'TablePerson'"
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
      Left            =   7200
      TabIndex        =   22
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtDeletion 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7320
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtRegistration 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7320
      TabIndex        =   19
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtRemark 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   1212
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
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
      Left            =   5520
      TabIndex        =   12
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete..."
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
      Left            =   6720
      TabIndex        =   11
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Sa&VeAs..."
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
      Left            =   2760
      TabIndex        =   9
      Top             =   6480
      Width           =   1092
   End
   Begin VB.ListBox lstPersonID 
      Enabled         =   0   'False
      Height          =   1740
      ItemData        =   "frmTableInfo.frx":0000
      Left            =   120
      List            =   "frmTableInfo.frx":0002
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCorrection 
      Caption         =   "Co&Rrection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   1215
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
      Height          =   4095
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optCategory 
         Caption         =   "Category"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   1455
      End
      Begin VB.OptionButton optSite 
         Caption         =   "Site"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton optBrigade 
         Caption         =   "Brigade"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton optDepartment 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optTNumber 
         Caption         =   "T_Number"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optSurName 
         Caption         =   "SurName"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optRemark 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1455
      End
      Begin VB.OptionButton optPersonID 
         Caption         =   "PersonID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optCardID 
         Caption         =   "CardID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.TextBox txtPersonID 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCardID 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdDefaultPers 
      Caption         =   "D&Efault from HDD"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find..."
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
      Left            =   4320
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableInfo 
      Height          =   1815
      Left            =   2160
      TabIndex        =   14
      Top             =   4560
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   9
      Cols            =   13
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
   Begin VB.Label lblDeletion 
      Alignment       =   2  'Center
      Caption         =   "Deletion"
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
      Left            =   7320
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblRegistration 
      Alignment       =   2  'Center
      Caption         =   "Registration"
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
      Left            =   7320
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblPersOrTerm 
      Alignment       =   2  'Center
      Caption         =   "PersonID "
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
      Left            =   360
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   6120
      Y2              =   6360
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   4440
      Y2              =   6360
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   9360
      X2              =   2040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   9360
      X2              =   9360
      Y1              =   4440
      Y2              =   120
   End
End
Attribute VB_Name = "frmTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             'Введенный пароль
Dim strPassword As String
            'Текущий номер корректируемой строки "Таблицы информации"
Dim intRowNumCorr As Integer
            'Текущий номер корректируемого столбца "Таблицы информации"
Dim intColNumCorr As Integer
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Таблицы информации"
Dim gInfo As ExtendInfo
            'Строка отсылаемого сообщения
Dim strMessage As String

            'Возврат в вызвавшую процедуру
Private Sub cmdCancel_Click()
            'Переменная "Кнопки + Иконки" в окне сообщений
    Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
    Dim strResponse As String
            
            'Были не сохраненные изменения в "Таблице информации"
    If gChangesTableInfo = True Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Окно собщения с запросом сохранения "Таблицы информации" - на экран
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            'Нажата кнопка "Да"
        If strResponse = vbYes Then
            'Сохранение "Tаблицы персон" в файле по умолчанию
            cmdSave_Click
        End If
    End If
    
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить текстовые поля
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            'Очистить список имен
    lstPersonID.Clear
            'Сбросить признак внесенных изменений в "Таблицу информации"
    gChangesTableInfo = False
            'Сделать невидимой текущую форму
    frmTableInfo.Visible = False
            'Сделать доступной форму "frmDemo"
    frmDemo.Enabled = True
            'Сделать видимой форму "frmDemo"
    frmDemo.Show
    
End Sub
            
            'Коррекция
Private Sub cmdCorrection_Click()
            
            ' "Таблица персон" не содержит нефиксированных строк
    If grdTableInfo.Rows = 1 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о невозможности коррекции
        If frmDemo.optEnglish = True Then
            MsgBox ("The TableInfo is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
    
    Else
            'Сделать доступными элементы управления Коррекцией
            '  "Таблицы информации"
        fraColName.Enabled = True
        optPersonID.Value = True
        txtPersonID.Enabled = True
        lstPersonID.Enabled = True
        txtCardID.Enabled = True
        txtTNumber.Enabled = True
        txtName.Enabled = True
        txtSurName.Enabled = True
        txtDepartment.Enabled = True
        txtBrigade.Enabled = True
        txtSite.Enabled = True
        txtCategory.Enabled = True
        txtRemark.Enabled = True
            'Очистить текстовые поля
        txtPersonID.Text = ""
        txtCardID.Text = ""
        txtRegistration.Text = ""
        txtDeletion.Text = ""
        txtTNumber.Text = ""
        txtName.Text = ""
        txtSurName.Text = ""
        txtDepartment.Text = ""
        txtBrigade.Text = ""
        txtSite.Text = ""
        txtCategory.Text = ""
        txtRemark.Text = ""
            'Очистить список имен
        lstPersonID.Clear
    
            'Столбец "Person or Terminal"
        grdTableInfo.Col = 0
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNumCorr = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNumCorr
            'Заполнение списка "lstPersonID" записями из "Таблицы информации"
            lstPersonID.AddItem grdTableInfo.Text
        Next
            'Выбрать  элемент списка
        lstPersonID.ListIndex = 0
            'Номер корректируемой строки - (1)
        intRowNumCorr = 1
        grdTableInfo.Row = intRowNumCorr
            'Включить опцию
        optPersonID_Click
    
    End If
    
End Sub

            'Процедура формирования "Таблицы информации"
            '  из "Таблицы Персон"
Private Sub cmdFormTableInfoFromTablePerson_Click()
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить текстовые поля
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            'Очистить список имен
    lstPersonID.Clear
            
            'Время
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            'Удаление из "Таблицы информации" всех существующих строк
    grdTableInfo.Rows = 2
    grdTableInfo.Row = 1
            'Цикл по всем нефиксированным строкам "Таблицы персон"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            'Текущая строка "Таблицы персон"
        gTablePerson.Row = intRowNum
            'Если длина Персонального кода в "Таблице персон" равна 16-и байтам
            '   - корректная (ранее не удаленная и не искаженная информация)
        gTablePerson.Col = 1
        If Len(Trim(gTablePerson.Text)) = 16 Then
            'Добавление строки в конец "Таблицы информации"
            If grdTableInfo.Row = grdTableInfo.Rows - 1 Then
                grdTableInfo.AddItem ""
                grdTableInfo.Row = grdTableInfo.Rows - 1
            End If
            'Изменение ячейки "PersonID" в "Таблице информации"
            gTablePerson.Col = 0
            grdTableInfo.Col = 0
            
            ' Если в поле "Info" нет признака Гостя
            If Left(Trim(gTablePerson.Text), 1) <> gVisitor Then
                grdTableInfo.Text = Trim(gTablePerson.Text)
                If Len(Trim(grdTableInfo.Text)) = 16 Then
                    If Mid(Trim(grdTableInfo.Text), 16, 1) = "+" Or _
                    Mid(Trim(grdTableInfo.Text), 16, 1) = "-" Then
                        grdTableInfo.Text = Trim(Left(Trim(grdTableInfo.Text), 15))
                    End If
                End If
            'Изменение ячейки "CardID" в "Таблице информации"
                gTablePerson.Col = 1
                grdTableInfo.Col = 1
                grdTableInfo.Text = Trim(gTablePerson.Text)
            'Текущий столбец "Таблицы информации" = 2 (Время и Дата Регистрации)
                grdTableInfo.Col = 2
                grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
                "  ||  " + Trim(gProtocol.strProtocDate)
            'Количество удалений/добавлений строк в "Таблице информации"
                grdTableInfo.Tag = grdTableInfo.Tag + 1
            'Установить признак внесенных изменений в "Таблицу информации"
                gChangesTableInfo = True
            'B поле "Info" есть признак Гостя
            Else
                grdTableInfo.RemoveItem grdTableInfo.Row
                grdTableInfo.Row = grdTableInfo.Rows - 1
            End If
        
        End If
    Next
            
            'Установить признак "Удалить старый файл" при
            '  выполнении команды "Save" или "Save As"
    grdTableInfo.Tag = -1
            'Установить фокус на кнопке "Correction"
    If frmTableInfo.Visible = True Then cmdCorrection.SetFocus


End Sub
            
            'Процедура восстановления умалчиваемых значений
            ' параметров "Таблицы информации"
Private Sub cmdDefaultPers_Click()
            
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить текстовые поля
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            'Очистить список имен
    lstPersonID.Clear
            
            'Загрузка формы "Таблица информации"
    Form_Load
            'Установить фокус на кнопке "Correction"
    If frmTableInfo.Visible = True Then cmdCorrection.SetFocus

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор корректируемой ячейки "Таблицы информации"
Private Sub grdTableInfo_Click()
            'Коррекция "включена"
    If lstPersonID.Enabled = True Then
            'Номер корректируемой строки "Таблицы информации"
        intRowNumCorr = grdTableInfo.RowSel
        grdTableInfo.Row = intRowNumCorr
            'Номер выбранного элемента списка
        lstPersonID.ListIndex = intRowNumCorr - 1
            'Номер корректируемого столбца "Таблицы информации"
        intColNumCorr = grdTableInfo.ColSel
        grdTableInfo.Col = intColNumCorr
            'Выбор корректируемой строки "Таблицы информации"
        lstPersonID_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstPersonID.Left, Y:=lstPersonID.Top
            'Выбор корректируемого столбца "Таблицы информации"
        Select Case intColNumCorr
            'Установить фокус на текстовом поле для Коррекции
            Case 0
            optPersonID.Value = True
            txtPersonID.SetFocus
            Case 1
            optCardID.Value = True
            txtCardID.SetFocus
            Case 2
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 3
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 4
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 5
            optName.Value = True
            txtName.SetFocus
            Case 6
            optSurName.Value = True
            txtSurName.SetFocus
            Case 7
            optDepartment.Value = True
            txtDepartment.SetFocus
            Case 8
            optBrigade.Value = True
            txtBrigade.SetFocus
            Case 9
            optSite.Value = True
            txtSite.SetFocus
            Case 10
            optCategory.Value = True
            txtCategory.SetFocus
            Case 11
            optRemark.Value = True
            txtRemark.SetFocus
            Case 12
            optRemark.Value = True
            txtRemark.SetFocus
        End Select
    End If
        
End Sub

            'Выбор корректируемой строки "Таблицы информации"
Private Sub lstPersonID_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            'Нажата левая кнопка "мыши"
    If Button = vbLeftButton Then
            'Номер корректируемой строки "Таблицы информации"
        intRowNumCorr = lstPersonID.ListIndex + 1
        grdTableInfo.Row = intRowNumCorr
            
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Коррекции или Отображения
        grdTableInfo.Col = 0
        txtPersonID.Text = grdTableInfo.Text
        grdTableInfo.Col = 1
        txtCardID.Text = grdTableInfo.Text
        grdTableInfo.Col = 2
        txtRegistration.Text = grdTableInfo.Text
        grdTableInfo.Col = 3
        txtDeletion.Text = grdTableInfo.Text
        grdTableInfo.Col = 4
        txtTNumber.Text = grdTableInfo.Text
        grdTableInfo.Col = 5
        txtName.Text = grdTableInfo.Text
        grdTableInfo.Col = 6
        txtSurName.Text = grdTableInfo.Text
        grdTableInfo.Col = 7
        txtDepartment.Text = grdTableInfo.Text
        grdTableInfo.Col = 8
        txtBrigade.Text = grdTableInfo.Text
        grdTableInfo.Col = 9
        txtSite.Text = grdTableInfo.Text
        grdTableInfo.Col = 10
        txtCategory.Text = grdTableInfo.Text
        grdTableInfo.Col = 11
        txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
        grdTableInfo.Col = intColNumCorr
    End If

End Sub

            'Выбрана опция - "PersonID"
Private Sub optPersonID_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 0
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtPersonID.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "CardID"
Private Sub optCardID_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 1
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtCardID.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtCardID.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "TNumber"
Private Sub optTNumber_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 4
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtTNumber.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtTNumber.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Name"
Private Sub optName_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 5
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtName.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtName.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "SurName"
Private Sub optSurName_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 6
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtSurName.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtSurName.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Department"
Private Sub optDepartment_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 7
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtDepartment.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtDepartment.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Brigade"
Private Sub optBrigade_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 8
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtBrigade.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtBrigade.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Site"
Private Sub optSite_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 9
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtSite.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtSite.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Category"
Private Sub optCategory_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 10
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtCategory.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtCategory.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtRemark.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            'Выбрана опция - "Remark"
Private Sub optRemark_Click()
            'Номер корректируемого столбца "Таблицы информации"
    intColNumCorr = 11
    grdTableInfo.Col = intColNumCorr
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtRemark.Enabled = True
            'Установить фокус на текстовом поле для Коррекции
    txtRemark.SetFocus
            'Сделать (не)доступными некоторые элементы управл. Коррекцией "Таблицы информации"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
            'Копирование ячеек "Таблицы информации" в текстовое поле
            '  для Отображения
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            'Восстановить номер корректируемого столбца "Таблицы информации"
    grdTableInfo.Col = intColNumCorr

End Sub

            'Процедура ввода и анализа Корректируемого имени "PersonID"
Private Sub txtPersonID_KeyPress(KeyAscii As Integer)
            'Имя введено
    If KeyAscii = vbKeyReturn Then
            'Имя в допустимом диапазоне
        If Len(Trim(txtPersonID.Text)) < 17 Then
            'Изменение имени "Person or Terminal" в "Таблице информации"
        grdTableInfo.Text = Trim(txtPersonID.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optCardID"
        optCardID.Value = True
            Exit Sub
            'Имя в недопустимом диапазоне
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "CardID"
Private Sub txtCardID_KeyPress(KeyAscii As Integer)
            'Код введен
    If KeyAscii = vbKeyReturn Then
            'Переход по ошибке преобразования данных
        On Error GoTo CardIDError
            'Персональный код в допустимом диапазоне
        If Len(Trim(txtCardID.Text)) = 16 Then
            'Изменение ячейки "CardID" в "Таблице информации"
            grdTableInfo.Text = Trim(txtCardID.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
            gChangesTableInfo = True
            'Включить опцию "optTNumber"
            optTNumber.Value = True
            Exit Sub
            'Персональный код в недопустимом диапазоне
CardIDError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "TNumber"
Private Sub txtTNumber_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtTNumber.Text)) > 8 Then
            txtTNumber.Text = Left(Trim(txtTNumber.Text), 8)
        End If
            'Изменение ячейки "TNumber" в "Таблице информации"
        grdTableInfo.Text = Trim(txtTNumber.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optName"
        optName.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Name"
Private Sub txtName_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtName.Text)) > 32 Then
            txtName.Text = Left(Trim(txtName.Text), 32)
        End If
            'Изменение ячейки "Name" в "Таблице информации"
        grdTableInfo.Text = Trim(txtName.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optSurName"
        optSurName.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "SurName"
Private Sub txtSurName_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtSurName.Text)) > 32 Then
            txtSurName.Text = Left(Trim(txtSurName.Text), 32)
        End If
            'Изменение ячейки "SurName" в "Таблице информации"
        grdTableInfo.Text = Trim(txtSurName.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optDeparttment"
        optDepartment.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Department"
Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtDepartment.Text)) > 32 Then
            txtDepartment.Text = Left(Trim(txtDepartment.Text), 32)
        End If
            'Изменение ячейки "Department" в "Таблице информации"
        grdTableInfo.Text = Trim(txtDepartment.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optRemark"
        optBrigade.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Brigade"
Private Sub txtBrigade_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtBrigade.Text)) > 16 Then
            txtBrigade.Text = Left(Trim(txtBrigade.Text), 16)
        End If
            'Изменение ячейки "Brigade" в "Таблице информации"
        grdTableInfo.Text = Trim(txtBrigade.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optRemark"
        optSite.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Site"
Private Sub txtSite_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtSite.Text)) > 16 Then
            txtSite.Text = Left(Trim(txtSite.Text), 16)
        End If
            'Изменение ячейки "Site" в "Таблице информации"
        grdTableInfo.Text = Trim(txtSite.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optRemark"
        optCategory.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Category"
Private Sub txtCategory_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtCategory.Text)) > 16 Then
            txtCategory.Text = Left(Trim(txtCategory.Text), 16)
        End If
            'Изменение ячейки "Category" в "Таблице информации"
        grdTableInfo.Text = Trim(txtCategory.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optRemark"
        optRemark.Value = True
        Exit Sub
    End If

End Sub
            
            'Процедура ввода и анализа Корректируемого "Remark"
Private Sub txtRemark_KeyPress(KeyAscii As Integer)
            
            'Информация введена
    If KeyAscii = vbKeyReturn Then
            'Удаление лишних символов из текстового поля
        If Len(Trim(txtRemark.Text)) > 64 Then
            txtRemark.Text = Left(Trim(txtRemark.Text), 64)
        End If
            'Изменение ячейки "Remark" в "Таблице информации"
        grdTableInfo.Text = Trim(txtRemark.Text)
            'Установить признак  внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
            'Включить опцию "optPersonID"
        optPersonID.Value = True
        Exit Sub
    End If

End Sub
            
            'Добавление строки в "Таблицу информации"
Private Sub cmdAdd_Click()
            'Имя и Персональный код в "Таблице информации"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить список имен
    lstPersonID.Clear
    
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя Имя персоны
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Add ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            'Получить от пользователя Персональный код
    strCardID = InputBox("CardID: 16 Characters !!!", "Add ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strCardID)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            'Имя или Персональный код не выбраны
    If strPersonID = "" Or strCardID = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            'Имя и Персональный код выбраны
    Else
            'Регистрация персонального кода в "Таблице информации"
        Call Reg(strCardID, strPersonID, "", "", "", "", "", "", "", "")
    End If
            'Установить фокус на кнопке "Add"
    If frmTableInfo.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            'Поиск строки в "Таблице информации"
Private Sub cmdFind_Click()
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Имя и Персональный код в "Таблице информации"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить список имен
    lstPersonID.Clear
    
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя - код жителя
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Find ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            'Получить от пользователя - код карты
    strCardID = InputBox("CardID: 16 Characters !!!", "Find ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strCardID)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            'Имя или Персональный код не выбраны
    If strPersonID = "" Or strCardID = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            'Имя и Персональный код выбраны
    Else
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
        grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = strCardID Then
            'Текущий столбец "Таблицы информации" = 0 (Имя)
                grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Отобразить текстовые поля "Таблицы информации"
                    grdTableInfo.Col = 0
                    txtPersonID.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 1
                    txtCardID.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 2
                    txtRegistration.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 3
                    txtDeletion.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 4
                    txtTNumber.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 5
                    txtName.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 6
                    txtSurName.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 7
                    txtDepartment.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 8
                    txtBrigade.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 9
                    txtSite.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 10
                    txtCategory.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 11
                    txtRemark.Text = Trim(grdTableInfo.Text)
            'Досрочный выход из цикла
                    Exit For
                End If
            End If
        Next
            'ИМЕНИ или ПЕРСОНАЛЬНОГО КОДА нет в "Таблице информации"
        If intRowNum = grdTableInfo.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
            frmDemo.BeepSound
            MsgBox ("Unexistent PersonID Or CardID")
        End If
            'Установить фокус на кнопке "Find"
        If frmTableInfo.Visible = True Then cmdFind.SetFocus
    End If
            
    
End Sub

            'Регистрация персонального кода в "Таблице информации"
            '   Код возврата: 0 - Регистрация выполнена успешно;
            '                 1 - Коррекция выполнена успешно;
            '                 2 - в Регистрации отказано.
Public Function Reg(ByVal vntCardID As Variant, ByVal strPersonID As String, _
ByVal strTNumber As String, ByVal strName As String, ByVal strSurName As String, _
ByVal strDepartment As String, ByVal strBrigade As String, _
ByVal strSite As String, ByVal strCategory As String, ByVal strRemark As String)
            'Номер текущей строки в "Таблицы информации"
Dim intRowNum As Integer
            
            'В Регистрации отказано
    Reg = 2
            
            'Если это ГОСТЬ - нет информации в "Таблице информации"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Exit Function
    End If

            'Удаление признака входа/выхода из поля ИМЕНИ
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
    
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
    grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
        If Trim(grdTableInfo.Text) = vntCardID Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
            grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Досрочный выход из цикла
                Exit For
            End If
        End If
    Next
        

            'Введенные ПЕРСОНАЛЬНЫЙ КОД & ИМЯ уже есть в "Таблице информации"
    If intRowNum < grdTableInfo.Rows Then
            'Коррекция
        Reg = 1
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Correction 'TableInfo'")
        Else
            MsgBox ("Korekcija 'TableInfo'")
        End If
            
            'Коррекция информации, хранящейся в "Таблице информации"
        Call frmTableInfo.Corr(strTNumber, strName, strSurName, _
        strDepartment, strBrigade, strSite, strCategory, strRemark)
            
            'Введенного ПЕРСОНАЛЬНОГО КОДА & ИМЕНИ нет в "Таблице информации"
    Else
            'Регистрация
        Reg = 0
            'Добавление строки в конец "Таблицы информации"
        grdTableInfo.AddItem strPersonID
        grdTableInfo.Row = grdTableInfo.Rows - 1
            'Изменение ячейки "Person or Terminal" в "Таблице информации"
        grdTableInfo.Col = 0
        grdTableInfo.Text = Trim(strPersonID)
            'Изменение ячейки "CardID" в "Таблице информации"
        grdTableInfo.Col = 1
        grdTableInfo.Text = Trim(vntCardID)
            'Изменение ячейки "TNumber" в "Таблице информации"
        grdTableInfo.Col = 4
            'Удаление лишних символов из текстового поля
        If Len(Trim(strTNumber)) > 8 Then
            strTNumber = Left(Trim(strTNumber), 8)
        End If
        grdTableInfo.Text = Trim(strTNumber)
            'Изменение ячейки "Name" в "Таблице информации"
        grdTableInfo.Col = 5
            'Удаление лишних символов из текстового поля
        If Len(Trim(strName)) > 32 Then
            strName = Left(Trim(strName), 32)
        End If
        grdTableInfo.Text = Trim(strName)
            'Изменение ячейки "SurName" в "Таблице информации"
        grdTableInfo.Col = 6
            'Удаление лишних символов из текстового поля
        If Len(Trim(strSurName)) > 32 Then
            strSurName = Left(Trim(strSurName), 32)
        End If
        grdTableInfo.Text = Trim(strSurName)
            'Изменение ячейки "Department" в "Таблице информации"
        grdTableInfo.Col = 7
            'Удаление лишних символов из текстового поля
        If Len(Trim(strDepartment)) > 32 Then
            strDepartment = Left(Trim(strDepartment), 32)
        End If
        grdTableInfo.Text = Trim(strDepartment)
            'Изменение ячейки "Brigade" в "Таблице информации"
        grdTableInfo.Col = 8
            'Удаление лишних символов из текстового поля
        If Len(Trim(strBrigade)) > 16 Then
            strBrigade = Left(Trim(strBrigade), 16)
        End If
        grdTableInfo.Text = Trim(strBrigade)
            'Изменение ячейки "Site" в "Таблице информации"
        grdTableInfo.Col = 9
            'Удаление лишних символов из текстового поля
        If Len(Trim(strSite)) > 16 Then
            strSite = Left(Trim(strSite), 16)
        End If
        grdTableInfo.Text = Trim(strSite)
            'Изменение ячейки "Category" в "Таблице информации"
        grdTableInfo.Col = 10
            'Удаление лишних символов из текстового поля
        If Len(Trim(strCategory)) > 16 Then
            strCategory = Left(Trim(strCategory), 16)
        End If
        grdTableInfo.Text = Trim(strCategory)
            'Изменение ячейки "Remark" в "Таблице информации"
        grdTableInfo.Col = 11
            'Удаление лишних символов из текстового поля
        If Len(Trim(strRemark)) > 64 Then
            strRemark = Left(Trim(strRemark), 64)
        End If
        grdTableInfo.Text = Trim(strRemark)
            
            'Время
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Текущий столбец "Таблицы информации" = 2 (Время и Дата Регистрации)
        grdTableInfo.Col = 2
        grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
        "  ||  " + Trim(gProtocol.strProtocDate)
            
            'Строка передачи сообщения
        strMessage = "RegInfo " + strPersonID + Chr(7) + _
        Trim(vntCardID) + Chr(7) + grdTableInfo.Text + Chr(7) + " " + Chr(7) + _
        strTNumber + Chr(7) + strName + Chr(7) + strSurName + Chr(7) + _
        strDepartment + Chr(7) + strBrigade + Chr(7) + strSite + Chr(7) + _
        strCategory + Chr(7) + strRemark + Chr(7) + " "
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Количество удалений/добавлений строк в "Таблице информации"
        grdTableInfo.Tag = grdTableInfo.Tag + 1
            'Установить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
    
            'Протоколировать данное событие
        gProtocol.strProtocName = strPersonID
        gProtocol.strProtocPersonCode = vntCardID
        gProtocol.strProtocStatus = "?? - TableInfo"
            'Примечания
        gProtocol.strProtocReserve = "Registration"
            'Записать строку в файл "Таблицы протокола"
        frmDemo.WriteProtocol
    End If
            
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения
            '  внимания - возможно скорое переполнение "Таблицы информации"
    If grdTableInfo.Rows > 32000 Then
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("'TableInfo' > 32000 rows")
        Else
            MsgBox ("'TableInfo' > 32000 rin.")
        End If
    End If
    
End Function
            
            'Поиск & Персонального кода & Имени в "Таблице информации"
            '   Код возврата: 0 - Поиск выполнен успешно;
            '                 1 - Поиск неуспешный.
Public Function Find(ByVal vntCardID As Variant, ByVal strPersonID As String, _
ByRef strTNumber As String, ByRef strName As String, ByRef strSurName As String, _
ByRef strDepartment As String, ByRef strBrigade As String, _
ByRef strSite As String, ByRef strCategory As String, ByRef strRemark As String)
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            
            
            'Если это ГОСТЬ - нет информации в "Таблице информации"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Find = 1
        Exit Function
    End If
            
            'Удаление признака входа/выхода из поля ИМЕНИ
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
            
            'Поиск неуспешный
    Find = 1
        
        'Поиск по Персональномуй коду & Имени
    If vntCardID <> "" And strPersonID <> "" Then
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
        grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = vntCardID Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
                grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Поиск выполнен успешно
                    Find = 0
            'Текущий столбец "Таблицы информации" = 4
                    grdTableInfo.Col = 4
                    strTNumber = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 5
                    grdTableInfo.Col = 5
                    strName = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 6
                    grdTableInfo.Col = 6
                    strSurName = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 7
                    grdTableInfo.Col = 7
                    strDepartment = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 8
                    grdTableInfo.Col = 8
                    strBrigade = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 9
                    grdTableInfo.Col = 9
                    strSite = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 10
                    grdTableInfo.Col = 10
                    strCategory = Trim(grdTableInfo.Text)
            'Текущий столбец "Таблицы информации" = 11
                    grdTableInfo.Col = 11
                    strRemark = Trim(grdTableInfo.Text)
            'Досрочный выход из цикла
                    Exit For
                End If
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
                grdTableInfo.Col = 1
            End If
        Next
    End If
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице информации"
    If Find = 1 Or intRowNum = grdTableInfo.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            'Поиск & Персонального кода & Имени в "Таблице информации"
            '   Код возврата: 0 - Поиск выполнен успешно;
            '                 1 - Поиск неуспешный.
Public Function Corr(ByVal strTNumber As String, ByVal strName As String, _
ByVal strSurName As String, ByVal strDepartment As String, _
ByVal strBrigade As String, ByVal strSite As String, ByVal strCategory As String, _
ByVal strRemark As String)
            'Персональный идентификатор - персональный код жителя
Dim strPersonID As String
            'Персональный код в системе - номер карточки
Dim strCardID As String
            'Время и Дата регистрации
Dim strTimeDateReg As String
            'Время и Дата исключения
Dim strTimeDateDel As String
            'Резерв
Dim strReserve As String
            
            'Извлечение информации из ячейки "PersonID" в "Таблице информации"
    grdTableInfo.Col = 0
    strPersonID = grdTableInfo.Text
            'Извлечение информации из ячейки "CardID" в "Таблице информации"
    grdTableInfo.Col = 1
    strCardID = grdTableInfo.Text
            'Извлечение информации из ячейки "TimeDateReg" в "Таблице информации"
    grdTableInfo.Col = 2
    strTimeDateReg = grdTableInfo.Text
            'Извлечение информации из ячейки "TimeDateDel" в "Таблице информации"
    grdTableInfo.Col = 3
    strTimeDateDel = grdTableInfo.Text
            'Извлечение информации из ячейки "Reserve" в "Таблице информации"
    grdTableInfo.Col = 12
    strReserve = grdTableInfo.Text
            
            'Изменение ячейки "TNumber" в "Таблице информации"
    grdTableInfo.Col = 4
            'Удаление лишних символов из текстового поля
    If Len(Trim(strTNumber)) > 8 Then
        strTNumber = Left(Trim(strTNumber), 8)
    End If
    grdTableInfo.Text = Trim(strTNumber)
            'Изменение ячейки "Name" в "Таблице информации"
    grdTableInfo.Col = 5
            'Удаление лишних символов из текстового поля
    If Len(Trim(strName)) > 32 Then
        strName = Left(Trim(strName), 32)
    End If
    grdTableInfo.Text = Trim(strName)
            'Изменение ячейки "SurName" в "Таблице информации"
    grdTableInfo.Col = 6
            'Удаление лишних символов из текстового поля
    If Len(Trim(strSurName)) > 32 Then
        strSurName = Left(Trim(strSurName), 32)
    End If
    grdTableInfo.Text = Trim(strSurName)
            'Изменение ячейки "Department" в "Таблице информации"
    grdTableInfo.Col = 7
            'Удаление лишних символов из текстового поля
    If Len(Trim(strDepartment)) > 32 Then
        strDepartment = Left(Trim(strDepartment), 32)
    End If
    grdTableInfo.Text = Trim(strDepartment)
            'Изменение ячейки "Brigade" в "Таблице информации"
    grdTableInfo.Col = 8
            'Удаление лишних символов из текстового поля
    If Len(Trim(strBrigade)) > 16 Then
        strBrigade = Left(Trim(strBrigade), 16)
    End If
    grdTableInfo.Text = Trim(strBrigade)
            'Изменение ячейки "Site" в "Таблице информации"
    grdTableInfo.Col = 9
            'Удаление лишних символов из текстового поля
    If Len(Trim(strSite)) > 16 Then
        strSite = Left(Trim(strSite), 16)
    End If
    grdTableInfo.Text = Trim(strSite)
            'Изменение ячейки "Category" в "Таблице информации"
    grdTableInfo.Col = 10
            'Удаление лишних символов из текстового поля
    If Len(Trim(strCategory)) > 16 Then
        strCategory = Left(Trim(strCategory), 16)
    End If
    grdTableInfo.Text = Trim(strCategory)
            'Изменение ячейки "Remark" в "Таблице информации"
    grdTableInfo.Col = 11
            'Удаление лишних символов из текстового поля
    If Len(Trim(strRemark)) > 64 Then
        strRemark = Left(Trim(strRemark), 64)
    End If
    grdTableInfo.Text = Trim(strRemark)
            
            'Время и Дату Коррекции в "Таблицу информации"
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
        
    If Len(Trim(strRemark)) = 64 Then
        grdTableInfo.Text = grdTableInfo.Text + "  " + _
        Trim(gProtocol.strProtocTime) + " || " + _
        Trim(gProtocol.strProtocDate)
    Else
        grdTableInfo.Text = grdTableInfo.Text + _
        Left("                                                                  ", _
        66 - Len(Trim(strRemark))) + _
        Trim(gProtocol.strProtocTime) + " || " + _
        Trim(gProtocol.strProtocDate)
    End If
            
            'Строка передачи сообщения
    strMessage = "CorInfo " + strPersonID + Chr(7) + strCardID + Chr(7) + _
    strTimeDateReg + Chr(7) + strTimeDateDel + Chr(7) + _
    Trim(strTNumber) + Chr(7) + Trim(strName) + Chr(7) + _
    Trim(strSurName) + Chr(7) + Trim(strDepartment) + Chr(7) + _
    Trim(strBrigade) + Chr(7) + Trim(strSite) + Chr(7) + _
    Trim(strCategory) + Chr(7) + Trim(strRemark) + Chr(7) + _
    Trim(strReserve) + Chr(7)
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
    Call frmDemo.SendMessage(strMessage)
            
            'Установить признак внесенных изменений в "Таблицу информации"
    gChangesTableInfo = True
            
            'Протоколировать данное событие
    grdTableInfo.Col = 0
    gProtocol.strProtocName = grdTableInfo.Text
    grdTableInfo.Col = 1
    gProtocol.strProtocPersonCode = grdTableInfo.Text
    gProtocol.strProtocStatus = "?? - TableInfo"
            'Примечания
    gProtocol.strProtocReserve = "Correction"
            'Записать строку в файл "Таблицы протокола"
    frmDemo.WriteProtocol
    
End Function
            'Удаление (ЛОГИЧЕСКОЕ) строки из "Таблицы информации"
            '   Код возврата: 0 - Удаление выполнено успешно;
            '                 1 - в удалении отказано.
Public Function Del(ByVal vntCardID As Variant, ByVal strPersonID As String)
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Время и Дата регистрации
Dim strTimeDateReg As String
            'Время и Дата исключения
Dim strTimeDateDel As String
            
            'Если это ГОСТЬ - нет информации в "Таблице информации"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Del = 1
        Exit Function
    End If
            
            'Удаление признака входа/выхода из поля ИМЕНИ
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
    
            'Удаление неуспешное
    Del = 1
        
        'Поиск по Персональномуй коду & Имени
    If vntCardID <> "" And strPersonID <> "" Then
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
        grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = vntCardID Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
                grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Поиск выполнен успешно
                    Del = 0
            'Логически удалить строку из "Таблицы информации" и
            '  Протоколировать данное событие
                    gProtocol.strProtocName = strPersonID
                    gProtocol.strProtocPersonCode = vntCardID
                    gProtocol.strProtocStatus = "?? - TableInfo"
            'Время
                    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                    gProtocol.strProtocReserve = "Logical Deletion"
            'Записать строку в файл "Таблицы протокола"
                    frmDemo.WriteProtocol
            
            'Текущий столбец "Таблицы информации" = 3 (Время и Дата Исключения)
                    grdTableInfo.Col = 3
                    grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
                    "  ||  " + Trim(gProtocol.strProtocDate)
                    
            'Извлечение информации из ячейки "TimeDateDel" в "Таблице информации"
                    strTimeDateDel = grdTableInfo.Text
            'Извлечение информации из ячейки "TimeDateReg" в "Таблице информации"
                    grdTableInfo.Col = 2
                    strTimeDateReg = grdTableInfo.Text

            'Строка передачи сообщения
                    strMessage = "DelInfo " + Trim(strPersonID) + Chr(7) + _
                    vntCardID + Chr(7) + strTimeDateReg + Chr(7) + _
                    strTimeDateDel + Chr(7)
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            'Установить признак внесенных изменений в "Таблицу информации"
                    gChangesTableInfo = True
            'Досрочный выход из цикла
                    Exit For
                End If
            End If
        Next
    End If
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице информации"
    If Del = 1 Or intRowNum = grdTableInfo.Rows Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            'Удаление (ФИЗИЧЕСКОЕ) строки из "Таблицы информации"
            '   Код возврата: 0 - Удаление выполнено успешно;
            '                 1 - в удалении отказано.
Public Function RealDel(ByVal vntCardID As Variant, ByVal strPersonID As String)
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Переменная "Кнопки + Иконки" в окне сообщений
Dim intButtonsAndIcons  As Integer
            'Строка ответа пользователя на вывод окна сообщения
Dim strResponse As String
            
            'Последняя строка не может быть удалена - выход из процедуры
    If grdTableInfo.Rows = 2 Then Exit Function
    
            'Удаление неуспешное
    RealDel = 1
        
        'Поиск по Персональномуй коду & Имени
    If vntCardID <> "" And strPersonID <> "" Then
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
        grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNum
            'Полученный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = vntCardID Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
                grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Поиск выполнен успешно
                    RealDel = 0
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
                    frmDemo.BeepSound
            'Окно собщения с повторным запросом удаления
            '   ПЕРСОНАЛЬНОГО КОДА - на экран
                    intButtonsAndIcons = vbYesNo + vbQuestion
                    If frmDemo.optEnglish = True Then
                        strResponse = MsgBox("Deletion Information ?", intButtonsAndIcons, "Cancel")
                    Else
                        strResponse = MsgBox("Izslegt Info ?", intButtonsAndIcons, "Cancel")
                    End If
            'Нажата кнопка "Да"
                    If strResponse = vbYes Then
            'Физическое удаление строки из "Таблицы информации"
                        grdTableInfo.RemoveItem grdTableInfo.Row
            'Количество удалений/добавлений строк в "Таблице информации"
                        grdTableInfo.Tag = grdTableInfo.Tag - 1
            
            'Протоколировать данное событие
                        gProtocol.strProtocName = strPersonID
                        gProtocol.strProtocPersonCode = vntCardID
                        gProtocol.strProtocStatus = "?? - TableInfo"
            'Время
                        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                        gProtocol.strProtocReserve = "Real Deletion"
            'Записать строку в файл "Таблицы протокола"
                        frmDemo.WriteProtocol
            
            'Установить признак внесенных изменений в "Таблицу информации"
                        gChangesTableInfo = True
                    End If
            'Досрочный выход из цикла
                    Exit For
                End If
            End If
        Next
    End If
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА нет в "Таблице информации"
    If RealDel = 1 Then
            'Подача ДЛИТЕЛЬНОГО звукового сигнала для привлечения внимания
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            'Удаление строки из "Таблицы информации"
Private Sub cmdDelete_Click()
            'Имя и Персональный код в "Таблице информации"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить список имен
    lstPersonID.Clear
    
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Получить от пользователя Имя персоны
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Delete ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            'Получить от пользователя Персональный код
    strCardID = InputBox("CardID: 16 Characters !!!", "Delete ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            'Длина персонального кода меньше 16-и символов
    If Len(Trim(strCardID)) < 16 Then
            'Добавить необходимое количество лидирующих нулей
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            'Имя или Персональный код не выбраны
    If strPersonID = "" Or strCardID = "" Then
            'Издать звуковой сигнал
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            'Имя и Персональный код выбраны
    Else
            'Реальное удаление персонального кода из "Таблицы информации"
        Call RealDel(strCardID, strPersonID)
    End If
            'Установить фокус на кнопке "Add"
    If frmTableInfo.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            'Сохранение "Таблицы информации" в файле по умолчанию
Public Function SaveTableInfo()
    Call cmdSave_Click
    SaveTableInfo = 0
    
End Function
            
            'Сохранение "Таблицы информации" в файле по умолчанию
Private Sub cmdSave_Click()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы информации"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы информации"
Dim intColNum As Integer
            
            'Если курсор мыши = "Песочные часы", то выйти
    If Me.MousePointer = vbHourglass Then Exit Sub
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    Me.MousePointer = vbHourglass
            
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить список имен
    lstPersonID.Clear
            
            'Вычислить длину записи (строки) "Таблицы информации"
    lngRecordLen = Len(gInfo)
            'Получить свободный номер файла
    intFileNum = FreeFile
    
    
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableInfo.dat"
    
            'Строк, удаленных из "Таблицы информации" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
    If grdTableInfo.Tag < 0 Then
        On Error Resume Next
            'Удалить "старый" умалчиваемый файл
        Kill strPathFileName
        On Error GoTo 0
    End If
    
            'Открыть умалчиваемый файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'По всем столбцам "Таблицы информации"
        For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            'Текущий столбец "Таблицы информации"
            grdTableInfo.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы информации"
            '  в файл
            Select Case intColNum
                Case 0
                gInfo.strPersonID = grdTableInfo.Text
                Case 1
                gInfo.strCardID = grdTableInfo.Text
                Case 2
                gInfo.strTimeDateReg = grdTableInfo.Text
                Case 3
                gInfo.strTimeDateDel = grdTableInfo.Text
                Case 4
                gInfo.strTNumber = grdTableInfo.Text
                Case 5
                gInfo.strName = grdTableInfo.Text
                Case 6
                gInfo.strSurName = grdTableInfo.Text
                Case 7
                gInfo.strDepartment = grdTableInfo.Text
                Case 8
                gInfo.strBrigade = grdTableInfo.Text
                Case 9
                gInfo.strSite = grdTableInfo.Text
                Case 10
                gInfo.strCategory = grdTableInfo.Text
                Case 11
                gInfo.strRemark = grdTableInfo.Text
                Case 12
                gInfo.strReserve = " "
            End Select
        Next
            'Записать строку "Таблицы информации" в файл
        Put intFileNum, intRowNum, gInfo
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице информации"
    grdTableInfo.Tag = 0
            'Сбросить признак внесенных изменений в "Таблицу информации"
    gChangesTableInfo = False
            'Восстановить стандартный курсор мыши
    Me.MousePointer = 0
            'Установить фокус на кнопке "Cancel"
    If frmTableInfo.Visible = True Then cmdCancel.SetFocus
            
End Sub
            
            'Сохранение "Таблицы информации" в выбираемом файле
Private Sub cmdSaveAs_Click()
            'Полное имя файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы информации"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы информации"
Dim intColNum As Integer

            'Если курсор мыши = "Песочные часы", то выйти
    If Me.MousePointer = vbHourglass Then Exit Sub
            
            'Сделать недоступными элементы управления Коррекцией
            '  "Таблицы информации"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            'Очистить список имен
    lstPersonID.Clear
            
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
            'Запись "Таблицы информации" в выбранный файл
    Else
            'Изменить стандартный курсор мыши  на "Песочные часы"
        Me.MousePointer = vbHourglass
            
            'Полное имя файла (с указанием "пути" к нему)
        strPathFileName = frmGetFile.Tag
            'Вычислить длину записи (строки) "Таблицы информации"
        lngRecordLen = Len(gInfo)
            'Получить свободный номер файла
        intFileNum = FreeFile
    
            'Строк, удаленных из "Таблицы информации" больше количества добавленных,
            ' т.е. умалчиваемый файл станет короче
        If grdTableInfo.Tag < 0 Then
            'Удалить "старый" файл, если он существует
            If Dir(strPathFileName) = strPathFileName Then
                On Error Resume Next
            'Удалить "старый" умалчиваемый файл
                Kill strPathFileName
                On Error GoTo 0
            End If
        End If

            'Открыть выбранный файл для произвольного доступа или
            '   создать его, если он не существует
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы информации"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
            grdTableInfo.Row = intRowNum
            'По всем столбцам "Таблицы информации"
            For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            'Текущий столбец "Таблицы информации"
                grdTableInfo.Col = intColNum
            'Заполнение буфера для записи текущей строки "Таблицы информации"
            '  в файл
                Select Case intColNum
                    Case 0
                    gInfo.strPersonID = grdTableInfo.Text
                    Case 1
                    gInfo.strCardID = grdTableInfo.Text
                    Case 2
                    gInfo.strTimeDateReg = grdTableInfo.Text
                    Case 3
                    gInfo.strTimeDateDel = grdTableInfo.Text
                    Case 4
                    gInfo.strTNumber = grdTableInfo.Text
                    Case 5
                    gInfo.strName = grdTableInfo.Text
                    Case 6
                    gInfo.strSurName = grdTableInfo.Text
                    Case 7
                    gInfo.strDepartment = grdTableInfo.Text
                    Case 8
                    gInfo.strBrigade = grdTableInfo.Text
                    Case 9
                    gInfo.strSite = grdTableInfo.Text
                    Case 10
                    gInfo.strCategory = grdTableInfo.Text
                    Case 11
                    gInfo.strRemark = grdTableInfo.Text
                    Case 12
                    gInfo.strReserve = " "
                End Select
            Next
            'Записать строку "Таблицы информации" в файл
            Put intFileNum, intRowNum, gInfo
        Next
            'Закрыть выбранный файл
        Close intFileNum
             'Количество удалений/добавлений строк в "Таблице информации"
        grdTableInfo.Tag = 0
               'Сбросить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = False
    End If
    
            'Выгрузить форму "frmGetFile"
    UnLoad frmGetFile
            'Освободить память, занимаемую выгруженной формой
    Set frmGetFile = Nothing
            'Восстановить стандартный курсор мыши
    Me.MousePointer = 0
            'Установить фокус на кнопке "Cancel"
    If frmTableInfo.Visible = True Then cmdCancel.SetFocus
    
End Sub

            'Загрузка формы "Таблица информации"
Private Sub Form_Load()
            'Полное имя умалчиваемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Длина строки "Таблицы информации"
Dim lngRecordLen As Long
            'Текущий номер нефиксированной строки "Таблицы информации"
Dim intRowNum As Integer
            'Текущий номер столбца "Таблицы информации"
Dim intColNum As Integer

            'Установить ширину столбцов
    SetColWidth
            'Текущая строка = 0 (Заголовки столбцов)
    grdTableInfo.Row = 0
    grdTableInfo.Col = 0
    grdTableInfo.Text = "PersonID"
            'Записать в ячейку (строка 0, столбец 1)
    grdTableInfo.Col = 1
    grdTableInfo.Text = "CardID"
            'Записать в ячейку (строка 0, столбец 2)
    grdTableInfo.Col = 2
    grdTableInfo.Text = "Time & Date Registration"
            'Записать в ячейку (строка 0, столбец 3)
    grdTableInfo.Col = 3
    grdTableInfo.Text = "Time & Date Deletion"
            'Записать в ячейку (строка 0, столбец 4)
    grdTableInfo.Col = 4
    grdTableInfo.Text = "TNumber"
            'Записать в ячейку (строка 0, столбец 5)
    grdTableInfo.Col = 5
    grdTableInfo.Text = "Name"
            'Записать в ячейку (строка 0, столбец 6)
    grdTableInfo.Col = 6
    grdTableInfo.Text = "SurName"
            'Записать в ячейку (строка 0, столбец 7)
    grdTableInfo.Col = 7
    grdTableInfo.Text = "Department"
            'Записать в ячейку (строка 0, столбец 8)
    grdTableInfo.Col = 8
    grdTableInfo.Text = "Brigade"
            'Записать в ячейку (строка 0, столбец 9)
    grdTableInfo.Col = 9
    grdTableInfo.Text = "Site"
            'Записать в ячейку (строка 0, столбец 10)
    grdTableInfo.Col = 10
    grdTableInfo.Text = "Category"
            'Записать в ячейку (строка 0, столбец 11)
    grdTableInfo.Col = 11
    grdTableInfo.Text = "Remark"
            'Записать в ячейку (строка 0, столбец 12)
    grdTableInfo.Col = 12
    grdTableInfo.Text = "Reserve"
            
            
            'Заполнение "Таблицы информации" из файла по умолчанию
            
            'Вычислить длину записи (строки) "Таблицы информации"
    lngRecordLen = Len(gInfo)
            'Получить свободный номер файла
    intFileNum = FreeFile
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableInfo.dat"
                
            'Файл отсутствует - ?
    On Error GoTo ErrorTableInfo
                'Количество строк "Таблицы информации"
    grdTableInfo.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            'Открыть умалчиваемый файл для произвольного доступа
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'Читать строку "Таблицы информации" из файла в буфер
        Get intFileNum, intRowNum, gInfo
            'По всем столбцам "Таблицы информации"
        For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            'Текущий столбец "Таблицы информации"
            grdTableInfo.Col = intColNum
            'Заполнение текущей строки "Таблицы информации" из буфера
            Select Case intColNum
                Case 0
                grdTableInfo.Text = gInfo.strPersonID
                Case 1
                grdTableInfo.Text = gInfo.strCardID
                Case 2
                grdTableInfo.Text = gInfo.strTimeDateReg
                Case 3
                grdTableInfo.Text = gInfo.strTimeDateDel
                Case 4
                grdTableInfo.Text = gInfo.strTNumber
                Case 5
                grdTableInfo.Text = gInfo.strName
                Case 6
                grdTableInfo.Text = gInfo.strSurName
                Case 7
                grdTableInfo.Text = gInfo.strDepartment
                Case 8
                grdTableInfo.Text = gInfo.strBrigade
                Case 9
                grdTableInfo.Text = gInfo.strSite
                Case 10
                grdTableInfo.Text = gInfo.strCategory
                Case 11
                grdTableInfo.Text = gInfo.strRemark
                Case 12
                grdTableInfo.Text = gInfo.strReserve
            End Select
        Next
    Next
            'Закрыть умалчиваемый файл
    Close intFileNum
            'Количество удалений/добавлений строк в "Таблице информации"
    grdTableInfo.Tag = 0
            'Сбросить признак внесенных изменений в "Таблицу информации"
    gChangesTableInfo = False
    
    Exit Sub
    
ErrorTableInfo:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox ("TableInfo Error !")
            'Количество удалений/добавлений строк в "Таблице информации"
    grdTableInfo.Tag = 0
    
End Sub
            
            'Добавление строки с заданным полями в "Таблицу информации"
            '  по сообщению MSMQ, полученому по сети
Public Function MSMQReg(ByVal strMessage As String)

            'Номер текущей строки в "Таблице информации"
Dim intRowNum As Integer
            'Номер анализируемой записи в строке сообщения
Dim intNumber As Integer
            'Пустая строка "Таблицы информации"
Dim strInfo As String
            
            'Персональный идентификатор - персональный код жителя
Dim strPersonID As String
            'Персональный код в системе - номер карточки
Dim strCardID As String
            'Время и Дата регистрации
Dim strTimeDateReg As String
            'Время и Дата исключения
Dim strTimeDateDel As String
            'Табельный номер
Dim strTNumber As String
            'Имя
Dim strName As String
            'Фамилия
Dim strSurName As String
            'Отдел/Цех
Dim strDepartment As String
            'Бригада/Подразделение
Dim strBrigade As String
            'Участок/Рабочее Место
Dim strSite As String
            'Категория/Должность/Профессия
Dim strCategory As String
            'Дополнительная информация
Dim strRemark As String
            'Резерв
Dim strReserve As String
            
            'Номер анализируемой записи в строке сообщения
    intNumber = 1
            'Искать в строке сообщения символы "07H" - разделители записей
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            'Подготовка ячейки "PersonID" для "Таблицы информации"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            'Подготовка ячейки "CardID" для "Таблицы информации"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            'Подготовка ячейки "TimeDateReg" для "Таблицы информации"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            'Подготовка ячейки "TimeDateDel" для "Таблицы информации"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            'Подготовка ячейки "TNumber" для "Таблицы информации"
            strTNumber = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 6 Then
            'Подготовка ячейки "Name" для "Таблицы информации"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 7 Then
            'Подготовка ячейки "SurName" для "Таблицы информации"
            strSurName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 8 Then
            'Подготовка ячейки "Department" для "Таблицы информации"
            strDepartment = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 9 Then
            'Подготовка ячейки "Brigade" для "Таблицы информации"
            strBrigade = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 10 Then
            'Подготовка ячейки "Site" для "Таблицы информации"
            strSite = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 11 Then
            'Подготовка ячейки "Category" для "Таблицы информации"
            strCategory = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 12 Then
            'Подготовка ячейки "Remark" для "Таблицы информации"
            strRemark = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 13 Then
            'Подготовка ячейки "Reserve" для "Таблицы информации"
            strReserve = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            'Принудительный выход
            Exit Do
        End If
    Loop
        
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
    grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
            grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Досрочный выход из цикла
                Exit For
            End If
        End If
    Next

            'Полученные ПЕРСОНАЛЬНЫЙ КОД & ИМЯ уже есть в "Таблице информации" -
            '   Коррекция
    If intRowNum < grdTableInfo.Rows Then
            'Изменение ячейки "PersonID" в "Таблице информации"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            'Изменение ячейки "CardID" в "Таблице информации"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            'Изменение ячейки "TimeDateReg" в "Таблице информации"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            'Изменение ячейки "TimeDateDel" в "Таблице информации"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            'Изменение ячейки "TNumber" в "Таблице информации"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            'Изменение ячейки "Name" в "Таблице информации"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            'Изменение ячейки "SurName" в "Таблице информации"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            'Изменение ячейки "Department" в "Таблице информации"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            'Изменение ячейки "Brigade" в "Таблице информации"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            'Изменение ячейки "Site" в "Таблице информации"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            'Изменение ячейки "Category" в "Таблице информации"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            'Изменение ячейки "Remark" в "Таблице информации"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            'Изменение ячейки "Reserve" в "Таблице информации"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
            
            'Установить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
    
            'Полученного ПЕРСОНАЛЬНОГО КОДА & ИМЕНИ нет в "Таблице информации" -
            '   Регистрация (Добавление)
    Else
            'Добавление строки в конец "Таблицы информации"
        grdTableInfo.AddItem strInfo
        grdTableInfo.Row = grdTableInfo.Rows - 1
            'Изменение ячейки "PersonID" в "Таблице информации"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            'Изменение ячейки "CardID" в "Таблице информации"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            'Изменение ячейки "TimeDateReg" в "Таблице информации"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            'Изменение ячейки "TimeDateDel" в "Таблице информации"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            'Изменение ячейки "TNumber" в "Таблице информации"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            'Изменение ячейки "Name" в "Таблице информации"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            'Изменение ячейки "SurName" в "Таблице информации"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            'Изменение ячейки "Department" в "Таблице информации"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            'Изменение ячейки "Brigade" в "Таблице информации"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            'Изменение ячейки "Site" в "Таблице информации"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            'Изменение ячейки "Category" в "Таблице информации"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            'Изменение ячейки "Remark" в "Таблице информации"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            'Изменение ячейки "Reserve" в "Таблице информации"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
    
            'Количество удалений/добавлений строк в "Таблице информации"
        grdTableInfo.Tag = grdTableInfo.Tag + 1
    
            'Установить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
    
    End If
    
End Function
            
            'Удаление (ЛОГИЧЕСКОЕ) строки с заданным персональным кодом
            '  из "Таблицы информации" по сообщению MSMQ, полученому по сети
Public Function MSMQDel(ByVal strMessage As String)

            'Номер текущей строки в "Таблице информации"
Dim intRowNum As Integer
            'Номер анализируемой записи в строке сообщения
Dim intNumber As Integer
            
            'Персональный идентификатор - персональный код жителя
Dim strPersonID As String
            'Персональный код в системе - номер карточки
Dim strCardID As String
            'Время и Дата регистрации
Dim strTimeDateReg As String
            'Время и Дата исключения
Dim strTimeDateDel As String
            
            'Номер анализируемой записи в строке сообщения
    intNumber = 1
            'Искать в строке сообщения символы "07H" - разделители записей
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            'Подготовка ячейки "PersonID" для "Таблицы информации"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            'Подготовка ячейки "CardID" для "Таблицы информации"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            'Подготовка ячейки "TimeDateReg" для "Таблицы информации"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            'Подготовка ячейки "TimeDateDel" для "Таблицы информации"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            'Принудительный выход
            Exit Do
        End If
    Loop
        
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
    grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
            grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Досрочный выход из цикла
                Exit For
            End If
        End If
    Next

            'Полученные ПЕРСОНАЛЬНЫЙ КОД & ИМЯ уже есть в "Таблице информации" -
            '   Коррекция
    If intRowNum < grdTableInfo.Rows Then
            'Изменение ячейки "TimeDateDel" в "Таблице информации"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            
            'Установить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
    
    End If

End Function
            
            'Коррекция заданных полей строки "Таблицы персон"
            '  по сообщению MSMQ, полученому по сети
Public Function MSMQCor(ByVal strMessage As String)

            'Номер текущей строки в "Таблице информации"
Dim intRowNum As Integer
            'Номер анализируемой записи в строке сообщения
Dim intNumber As Integer
            
            'Персональный идентификатор - персональный код жителя
Dim strPersonID As String
            'Персональный код в системе - номер карточки
Dim strCardID As String
            'Время и Дата регистрации
Dim strTimeDateReg As String
            'Время и Дата исключения
Dim strTimeDateDel As String
            'Табельный номер
Dim strTNumber As String
            'Имя
Dim strName As String
            'Фамилия
Dim strSurName As String
            'Отдел/Цех
Dim strDepartment As String
            'Бригада/Подразделение
Dim strBrigade As String
            'Участок/Рабочее Место
Dim strSite As String
            'Категория/Должность/Профессия
Dim strCategory As String
            'Дополнительная информация
Dim strRemark As String
            'Резерв
Dim strReserve As String
            
            'Номер анализируемой записи в строке сообщения
    intNumber = 1
            'Искать в строке сообщения символы "07H" - разделители записей
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            'Подготовка ячейки "PersonID" для "Таблицы информации"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            'Подготовка ячейки "CardID" для "Таблицы информации"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            'Подготовка ячейки "TimeDateReg" для "Таблицы информации"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            'Подготовка ячейки "TimeDateDel" для "Таблицы информации"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            'Подготовка ячейки "TNumber" для "Таблицы информации"
            strTNumber = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 6 Then
            'Подготовка ячейки "Name" для "Таблицы информации"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 7 Then
            'Подготовка ячейки "SurName" для "Таблицы информации"
            strSurName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 8 Then
            'Подготовка ячейки "Department" для "Таблицы информации"
            strDepartment = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 9 Then
            'Подготовка ячейки "Brigade" для "Таблицы информации"
            strBrigade = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 10 Then
            'Подготовка ячейки "Site" для "Таблицы информации"
            strSite = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 11 Then
            'Подготовка ячейки "Category" для "Таблицы информации"
            strCategory = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 12 Then
            'Подготовка ячейки "Remark" для "Таблицы информации"
            strRemark = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 13 Then
            'Подготовка ячейки "Reserve" для "Таблицы информации"
            strReserve = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            'Принудительный выход
            Exit Do
        End If
    Loop
        
        'Текущий столбец "Таблицы информации" = 1 (Персональный код)
    grdTableInfo.Col = 1
            'Цикл по всем нефиксированным строкам "Таблицы информации"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            'Текущая строка "Таблицы информации"
        grdTableInfo.Row = intRowNum
            'Введенный ПЕРСОНАЛЬНЫЙ КОД есть в "Таблице информации"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            'Текущий столбец "Таблицы информации" = 0 (Персона или Терминал)
            grdTableInfo.Col = 0
            'Полученное ИМЯ есть в "Таблице информации"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            'Досрочный выход из цикла
                Exit For
            End If
        End If
    Next

            'Полученные ПЕРСОНАЛЬНЫЙ КОД & ИМЯ уже есть в "Таблице информации" -
            '   Коррекция
    If intRowNum < grdTableInfo.Rows Then
            'Изменение ячейки "PersonID" в "Таблице информации"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            'Изменение ячейки "CardID" в "Таблице информации"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            'Изменение ячейки "TimeDateReg" в "Таблице информации"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            'Изменение ячейки "TimeDateDel" в "Таблице информации"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            'Изменение ячейки "TNumber" в "Таблице информации"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            'Изменение ячейки "Name" в "Таблице информации"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            'Изменение ячейки "SurName" в "Таблице информации"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            'Изменение ячейки "Department" в "Таблице информации"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            'Изменение ячейки "Brigade" в "Таблице информации"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            'Изменение ячейки "Site" в "Таблице информации"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            'Изменение ячейки "Category" в "Таблице информации"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            'Изменение ячейки "Remark" в "Таблице информации"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            'Изменение ячейки "Reserve" в "Таблице информации"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
            
            'Установить признак внесенных изменений в "Таблицу информации"
        gChangesTableInfo = True
    
    End If
    
End Function

            'Процедура установки ширины и выравнивания столбцов "Таблицы информации"
Public Sub SetColWidth()
            'Объявление переменной - текущий номер столбца
Dim intColNumber As Integer
            'Цикл по всем столбцам
    For intColNumber = 0 To grdTableInfo.Cols - 1 Step 1
        grdTableInfo.ColWidth(intColNumber) = 2500
        grdTableInfo.ColAlignment(intColNumber) = 0
    Next
            'Уменьшение размера 2-го и 3-го столбцов (Время и Дата Рег/Искл)
    intColNumber = 2
    grdTableInfo.ColWidth(intColNumber) = 1850
    intColNumber = 3
    grdTableInfo.ColWidth(intColNumber) = 1850
            'Уменьшение размера 4-го столбцa (Табельный номер)
    intColNumber = 4
    grdTableInfo.ColWidth(intColNumber) = 1300
            'Увеличение размера 5-го и 6-го столбцов (Имя, Фамилия и Цех)
    intColNumber = 5
    grdTableInfo.ColWidth(intColNumber) = 4300
    intColNumber = 6
    grdTableInfo.ColWidth(intColNumber) = 4300
    intColNumber = 7
    grdTableInfo.ColWidth(intColNumber) = 4300
            'Увеличение размера 11-го столбца (Информация)
    intColNumber = 11
    grdTableInfo.ColWidth(intColNumber) = 4950

End Sub
