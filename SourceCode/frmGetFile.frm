VERSION 5.00
Begin VB.Form frmGetFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmFileSelection"
   ClientHeight    =   4665
   ClientLeft      =   3375
   ClientTop       =   3495
   ClientWidth     =   4815
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
   ScaleHeight     =   4665
   ScaleWidth      =   4815
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
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
      Left            =   3000
      TabIndex        =   11
      Top             =   4080
      Width           =   1212
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
      Left            =   600
      TabIndex        =   10
      Top             =   4080
      Width           =   1212
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   330
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   2172
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   1290
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   2172
   End
   Begin VB.ComboBox cboFileType 
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3600
      Width           =   2172
   End
   Begin VB.FileListBox filFiles 
      Height          =   1350
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2172
   End
   Begin VB.TextBox txtFileName 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3240
      Width           =   1572
   End
   Begin VB.Label lblDirName 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label lblDitectories 
      Caption         =   "Directories"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label lblFileType 
      Caption         =   "File type"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Label lblFileName 
      Caption         =   "File name"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "frmGetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            
            'Изменение типа файла
Private Sub cboFileType_Click()
Dim intPatternPos1 As Integer
Dim intPatternPos2 As Integer
Dim intPatternLen As Integer
Dim strPattern As String
            'Найти в списке комбинированного поля "cboFileType" начало маски
    intPatternPos1 = InStr(1, cboFileType.Text, "(") + 1
            'Найти в списке комбинированного поля "cboFileType" конец маски
    intPatternPos2 = InStr(1, cboFileType.Text, ")") - 1
            'Вычислить длину маски
    intPatternLen = intPatternPos2 - intPatternPos1 + 1
            'Извлечь строку маски из списка комбинированного поля "cboFileType"
    strPattern = Mid(cboFileType.Text, intPatternPos1, intPatternLen)
            'Передать маску в список файлов
    filFiles.Pattern = strPattern

End Sub

            'Обработка события "Cancel"
Private Sub cmdCancel_Click()
            'Очистить своство "Tag" формы
    frmGetFile.Tag = ""
            'Убрать с экрана форму
    frmGetFile.Hide

End Sub
            
            'Обработка события "OK"
Private Sub cmdOK_Click()
Dim strPathAndName As String
Dim strPath As String
            'Если файл не выбран, то выйти
    If txtFileName = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The file isn't selected !", vbExclamation, "Error"
        Exit Sub
    End If
            'Строка пути к файлу должна заканчиваться символом "\"
    If Right(filFiles.Path, 1) <> "\" Then
        strPath = filFiles.Path + "\"
    Else
        strPath = filFiles.Path
    End If
            'Извлечь имя выбранного файла и путь к нему
    If txtFileName.Text = filFiles.FileName Then
        strPathAndName = strPath + filFiles.FileName
    Else
        strPathAndName = strPath + txtFileName.Text
    End If
            'Сохранить полное имя файла в свойстве "Tag" формы
    frmGetFile.Tag = strPathAndName
            'Убрать с экрана форму
    frmGetFile.Hide

End Sub

            'Изменение каталога
Private Sub dirDirectory_Change()
            'Изменить путь в списке файлов
    filFiles.Path = dirDirectory.Path
            'Обновить этикетку "lblDirName"
    lblDirName.Caption = dirDirectory.Path
    
End Sub

            'Изменение списка дисководов
Private Sub drvDrive_Change()
            'Устанавливаем метку перехода по ошибке
    On Error GoTo DriveError
            'Изменить путь в списке каталогов на новое устройство
    dirDirectory.Path = drvDrive.Drive
            'Ошибок нет, выйти
    Exit Sub
            'Обработка ошибки
DriveError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
            'Произошла ошибка, сообщить об этом пользователю
            ' и восстановить состояние списка устройств
    MsgBox "The drive selection Error !", vbExclamation, "Error"
    drvDrive.Drive = dirDirectory.Path
    Exit Sub
    
End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Процедура ввода и анализа имени файла
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
            'Имя файла введено
    If KeyAscii = vbKeyReturn Then
            'Выполнить процедуру "cmdOK_Click()"
        cmdOK_Click
    End If
    
End Sub
            
            'Выбран файл
Private Sub filFiles_Click()
            'Обновить имя файла в поле "txtFileName
    txtFileName.Text = filFiles.FileName
    
End Sub
            'Выбран файл
Private Sub filFiles_DblClick()
            'Обновить имя файла в поле "txtFileName
    txtFileName.Text = filFiles.FileName
            'Выполнить процедуру "cmdOK_Click()"
    cmdOK_Click

End Sub

            'Загрузка формы
Private Sub Form_Load()
            'Инициализировать этикетку "lblDirName"
    lblDirName.Caption = dirDirectory.Path
   
End Sub
