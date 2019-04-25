VERSION 5.00
Begin VB.Form frmPreprocessors 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preprocessors"
   ClientHeight    =   3525
   ClientLeft      =   6705
   ClientTop       =   2745
   ClientWidth     =   2925
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
   ScaleHeight     =   3525
   ScaleWidth      =   2925
   Visible         =   0   'False
   Begin VB.CommandButton cmdBookKeeperBase 
      Caption         =   "BookKeeper Base"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CommandButton cmdProtocolBase 
      Caption         =   "Protocol Base"
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1212
   End
   Begin VB.CommandButton cmdArchives 
      Caption         =   "DownLoad Archives"
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
      Left            =   840
      TabIndex        =   7
      Top             =   1560
      Width           =   1212
   End
   Begin VB.CommandButton cmdStopWorkStation 
      Caption         =   "Stop WorkStation"
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
      TabIndex        =   6
      Top             =   960
      Width           =   1212
   End
   Begin VB.ComboBox cboPreprocessors 
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRestartWorkStation 
      Caption         =   "Restart WorkStation"
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
      TabIndex        =   3
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton cmdProtocol 
      Caption         =   "DownLoad Protocol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox cboFileName 
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label lblPreprocessors 
      Caption         =   "Preproc. =>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPreprocessors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            'Текущий номер файла
Dim intFileNum As Integer
            'Строка "Таблицы протокола"
Dim gSystem As SystemInfo



            'Обработка "щелчка" мыши на спискe "cboPreprocessors"
Private Sub cboPreprocessors_Click()
            
            'Выбрать указанный элемент списка
    cboFileName.ListIndex = _
    cboPreprocessors.ListIndex

End Sub

            'Обработка события "RestartWorkStation"
Private Sub cmdRestartWorkStation_Click()

            'Строка передачи сообщения на "Preprocessor"
Dim strMessage As String

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If

            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Строка передачи сообщения на "Preprocessor"
    strMessage = "StartApp"
            'Выбраны все Препроцессоры: "Whole"
    If Trim(cboFileName.Text) = "Whole" Then
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Выбран один Препроцессор
    Else
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = strMessage
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        Trim(cboPreprocessors.Text) + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
        qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
        qQueueOutput.Close
    End If
            
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Обработка события "StopWorkStation"
Private Sub cmdStopWorkStation_Click()

            'Строка передачи сообщения на "Preprocessor"
Dim strMessage As String

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If

            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Строка передачи сообщения на "Preprocessor"
    strMessage = "StopApp"
            'Выбраны все Препроцессоры: "Whole"
    If Trim(cboFileName.Text) = "Whole" Then
            'Процедура передачи сообщения
            '  средствами сервиса "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            'Выбран один Препроцессор
    Else
            'Формирование ОТСЫЛАЕМОГО СООБЩЕНИЯ
        qMsgOutput.Body = strMessage
            'Установить путь к очереди ПЕРЕДАВАЕМЫХ СООБЩЕНИЙ
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        Trim(cboPreprocessors.Text) + "\Private$\GeneralQueue"
            'Открыть очередь сообщений с параметрами (для передачи
            '  сообщений, доступ к очереди разрешен всем)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            'Отослать СООБЩЕНИЕ
        qMsgOutput.Send qQueueOutput
            'Закрыть очередь СООБЩЕНИЙ
        qQueueOutput.Close
    End If
            
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Обработка события "Cancel"
Private Sub cmdCancel_Click()

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Обработка события "DownLoad Protocol"
Private Sub cmdProtocol_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Код возврата при сохранении "Системной таблицы"
Dim intSaveTableSystem As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Полное имя выбираемой папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Длина строки копируемой из Препроцессора "Таблицы протокола"
Dim lngRecordLen As Long
            'Количество строк в копируемой таблице
            '   "TableProtocol" Препроцессора
Dim intRowQuanP As Integer
            'Текущий номер строки копируемой таблицы
            '   "TableProtocol" Препроцессора
Dim intRowNumP As Integer

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If

            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Полное имя папки-файла Препроцессора
            '   (с указанием "пути" к ней) или "Whole"
    strPathFolderName = cboFileName.Text
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Выбраны все Препроцессоры: "Whole"
    If strPathFolderName = "Whole" Then
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                frmTableSystem.grdTableSystem.Col = 0
            'Выбрать элемент из списка комбинированного поля "cboFileName"
            '   - папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            
            'Проверка существования папки-файла Препроцессора
                On Error GoTo UnAccessable
                If (FSO.FolderExists(strPathFolderName)) Then
            'Папка-файл имеется - продолжить
                    On Error GoTo CopyingMistake
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
                    If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование ранее не скопированной части файла
            '   таблицы "TableProtocol" из папки-файла Препроцессора в конец
            '   файла таблицы "TableProtocol" для "Host Computer'a"
                
            'Вычислить длину записи (строки)
            '   "Таблицы протокола" Препроцессора
                        lngRecordLen = Len(gProtocol)
            'Количество строк в "Таблице протокола" Препроцессора
                        intRowQuanP = FileLen(strPathFileName) / lngRecordLen
            'Текущий столбец "Системной таблицы" = 4 (Аппендикс)
                        frmTableSystem.grdTableSystem.Col = 4
            'Ни одна строка "Таблицы протокола" Препроцессора ранее не
            '   копировалась в "Host Computer" или "Таблица протокола"
            '   Препроцессора начата Препроцессором заново (например,
            '   при Архивировании) - копировать с первой строки
                        If Trim(frmTableSystem.grdTableSystem.Text) = "" Or _
                        intRowQuanP < _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = 1
            'Все строки "Таблицы протокола" Препроцессора ранее были
            '   скопированы в "Host Computer" или "Таблица протокола"
            '   еще не подготовлена Препроцессором к копированию -
            '   к следующему Препроцессору
                        ElseIf intRowQuanP = 0 Or intRowQuanP = _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            GoTo EndCycle
            'Не все строки "Таблицы протокола" Препроцессора ранее были
            '   скопированы в "Host Computer" - копировать оставшиеся строки
                        ElseIf intRowQuanP > _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = Trim(frmTableSystem.grdTableSystem.Text) _
                            + 1
                        End If
            
            'Получить свободный номер файла
            '   "Таблицы протокола" Препроцессора
                        intFileNum = FreeFile
            'Открыть файл "TableProtocol" Препроцессора для
            '   произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Сформировать разделительную полосу
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
            'Протоколирование события - "Сгрузить Протокол Препроцессора"
                        gProtocol.strProtocName = strPathFolderName
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                        gProtocol.strProtocReserve = "DownLoad Protocol"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                        
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                        For intRowNumP = intRowNumP To intRowQuanP Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                            Get intFileNum, intRowNumP, gProtocol
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                        Close intFileNum
            'Запомнить номер последней строки "Таблицы протокола"
            '   Препроцессора, скопированной в "Host Computer"
                        frmTableSystem.grdTableSystem.Text = intRowQuanP
            'Установить признак внесенных в "Системную таблицу" изменений
                        gChangesTableSystem = True
            'Сформировать разделительную полосу
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                
                    Else
            'Файл отсутствует - выход из процедуры с сообщением
                        GoTo CopyingMistake
                    End If
                
                Else
            'Папка-файл отсутствует - выход из процедуры с сообщением
                    GoTo UnAccessable
                End If
                
                
                GoTo EndCycle
UnAccessable:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndCycle
CopyingMistake:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
            End If
EndCycle:
        Next
            'Выбран один Препроцессор
    Else
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 0 (Объект)
            frmTableSystem.grdTableSystem.Col = 0
            'Требуемый Препроцессор
            If Trim(frmTableSystem.grdTableSystem.Text) = Trim(cboFileName.Text) Then
            'Выбрать элемент из списка комбинированного поля "cboFileName"
            '   - папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(cboFileName.Text)
            
            'Проверка существования папки-файла Препроцессора
                On Error GoTo UnExist
                If (FSO.FolderExists(strPathFolderName)) Then
            'Папка-файл имеется - продолжить
                    On Error GoTo CopyingError
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
                    If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование ранее не скопированной части файла
            '   таблицы "TableProtocol" из папки-файла Препроцессора в конец
            '   файла таблицы "TableProtocol" для "Host Computer'a"
                
            'Вычислить длину записи (строки)
            '   "Таблицы протокола" Препроцессора
                        lngRecordLen = Len(gProtocol)
            'Количество строк в "Таблице протокола" Препроцессора
                        intRowQuanP = FileLen(strPathFileName) / lngRecordLen
            'Текущий столбец "Системной таблицы" = 4 (Аппендикс)
                        frmTableSystem.grdTableSystem.Col = 4
            'Ни одна строка "Таблицы протокола" Препроцессора ранее не
            '   копировалась в "Host Computer" или "Таблица протокола"
            '   Препроцессора начата Препроцессором заново (например,
            '   при Архивировании) - копировать с первой строки
                        If Trim(frmTableSystem.grdTableSystem.Text) = "" Or _
                        intRowQuanP < _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = 1
            'Все строки "Таблицы протокола" Препроцессора ранее были
            '   скопированы в "Host Computer" или "Таблица протокола"
            '   еще не подготовлена Препроцессором к копированию -
            '   к следующему Препроцессору
                        ElseIf intRowQuanP = 0 Or intRowQuanP = _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            GoTo EndProcedure
            'Не все строки "Таблицы протокола" Препроцессора ранее были
            '   скопированы в "Host Computer" - копировать оставшиеся строки
                        ElseIf intRowQuanP > _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = Trim(frmTableSystem.grdTableSystem.Text) _
                            + 1
                        End If
            'Получить свободный номер файла
            '   "Таблицы протокола" Препроцессора
                        intFileNum = FreeFile
            'Открыть файл "TableProtocol" Препроцессора для
            '   произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Сформировать разделительную полосу
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
            'Протоколирование события - "Сгрузить Протокол Препроцессора"
                        gProtocol.strProtocName = strPathFolderName
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                        gProtocol.strProtocReserve = "DownLoad Protocol"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                        
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                        For intRowNumP = intRowNumP To intRowQuanP Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                            Get intFileNum, intRowNumP, gProtocol
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                        Close intFileNum
            'Запомнить номер последней строки "Таблицы протокола"
            '   Препроцессора, скопированной в "Host Computer"
                        frmTableSystem.grdTableSystem.Text = intRowQuanP
            'Установить признак внесенных в "Системную таблицу" изменений
                        gChangesTableSystem = True
            'Сформировать разделительную полосу
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                
                    Else
            'Файл отсутствует - выход из процедуры с сообщением
                        GoTo CopyingError
                    End If
                
                Else
            'Папка-файл отсутствует - выход из процедуры с сообщением
                    GoTo UnExist
                End If
                
                GoTo EndProcedure
UnExist:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndProcedure
CopyingError:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
                GoTo EndProcedure
            End If
        Next
EndProcedure:
    End If
            
            'Установлен признак внесенных в "Системную таблицу"
            '   изменений - сохранить таблицу в умалчиваемом файле
    If gChangesTableSystem = True Then
            'Сохранить измененную 'Системную таблицу"
        intSaveTableSystem = frmTableSystem.SaveTableSystem()
    End If
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Обработка события "DownLoad Archives"
Private Sub cmdArchives_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Полное имя файла-копии (с указанием "пути" к нему)
Dim strHostFileName As String
            'Полное имя выбираемой папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Номер дня (обратный отсчет, начиная с текущего дня),
            '  который просматривается системой при копировании
            '  Архивов Препроцессоа в "Host Computer"
Dim intDayArchiveCopy As Integer

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If

            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Полное имя папки-файла Препроцессора
            '   (с указанием "пути" к ней) или "Whole"
    strPathFolderName = cboFileName.Text
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strHostFileName = App.Path
    If Right(strHostFileName, 1) <> "\" Then
            'Полное имя папки "Host Computera" для файла-копии
            '  Препроцессора(с указанием "пути" к ней)
        strHostFileName = strHostFileName + "\"
    End If
    
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Выбраны все Препроцессоры: "Whole"
    If strPathFolderName = "Whole" Then
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                frmTableSystem.grdTableSystem.Col = 0
            'Выбрать элемент из списка комбинированного поля "cboFileName"
            '   - папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            
            'Проверка существования папки-файла Препроцессора
                On Error GoTo UnAccessable
            'Папка-файл имеется - продолжить
                If (FSO.FolderExists(strPathFolderName)) Then
            'Протоколирование события - "Сгрузить Архивы Препроцессора"
                    gProtocol.strProtocName = strPathFolderName
            'Системный пароль
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                    gProtocol.strProtocStatus = "04 - Manager"
            'Время
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                    gProtocol.strProtocReserve = "DownLoad Archives"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                    On Error GoTo CopyingMistake
                    
            'Установка "Календаря" на Текущую дату
                    frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                    For intDayArchiveCopy = 1 To gDayNum Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                        frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива Препроцессора
            '  (с указанием "пути" к нему)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            'Формирование Примечания
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

                        If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование Архива в "Host Computer"
                            FSO.CopyFile strPathFileName, _
                            strHostFileName
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        
            'Протоколирование события - "Копирование Архива Препроцессора"
                            gProtocol.strProtocName = "Copy Archive"
            'Системный пароль
                            gProtocol.strProtocPersonCode = _
                            frmDemo.txtPassword.Tag
            'Статус
                            gProtocol.strProtocStatus = "04 - Manager"
            'Время
                            gProtocol.strProtocTime = _
                            Format(Now, "h:mm:ss")
            'Дата
                            gProtocol.strProtocDate = _
                            Format(Now, "dd/mm/yyyy")
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            'Установка "Календаря" на Предыдущую дату
                        frmTableCalendar.comCalendar.PreviousDay
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
                
            'Папка-файл отсутствует - выход из процедуры с сообщением
                Else
                    GoTo UnAccessable
                End If
                
                
                GoTo EndCycle
UnAccessable:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndCycle
CopyingMistake:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
            End If
EndCycle:
        Next
            'Выбран один Препроцессор
    Else
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 0 (Объект)
            frmTableSystem.grdTableSystem.Col = 0
            'Требуемый Препроцессор
            If Trim(frmTableSystem.grdTableSystem.Text) = Trim(cboFileName.Text) Then
            'Выбрать элемент из списка комбинированного поля "cboFileName"
            '   - папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(cboFileName.Text)
            
            'Проверка существования папки-файла Препроцессора
                On Error GoTo UnExist
            'Папка-файл имеется - продолжить
                If (FSO.FolderExists(strPathFolderName)) Then
            'Протоколирование события - "Сгрузить Архивы Препроцессора"
                    gProtocol.strProtocName = strPathFolderName
            'Системный пароль
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                    gProtocol.strProtocStatus = "04 - Manager"
            'Время
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
                    gProtocol.strProtocReserve = "DownLoad Archives"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                    On Error GoTo CopyingError
                    
            'Установка "Календаря" на Текущую дату
                    frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                    For intDayArchiveCopy = 1 To gDayNum Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                        frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива Препроцессора
            '  (с указанием "пути" к нему)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            'Формирование Примечания
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

                        If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование Архива в "Host Computer"
                            FSO.CopyFile strPathFileName, _
                            strHostFileName
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        
            'Протоколирование события - "Копирование Архива Препроцессора"
                            gProtocol.strProtocName = "Copy Archive"
            'Системный пароль
                            gProtocol.strProtocPersonCode = _
                            frmDemo.txtPassword.Tag
            'Статус
                            gProtocol.strProtocStatus = "04 - Manager"
            'Время
                            gProtocol.strProtocTime = _
                            Format(Now, "h:mm:ss")
            'Дата
                            gProtocol.strProtocDate = _
                            Format(Now, "dd/mm/yyyy")
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            'Установка "Календаря" на Предыдущую дату
                        frmTableCalendar.comCalendar.PreviousDay
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
                
            'Папка-файл отсутствует - выход из процедуры с сообщением
                Else
                    GoTo UnExist
                End If
                
                GoTo EndProcedure
UnExist:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndProcedure
CopyingError:
            'Издать звуковой сигнал
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
                GoTo EndProcedure
            End If
        Next
EndProcedure:
    End If
            
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub
            
            'Формирование баз Протокола и Бухгалтерии в формате ACCESS"
Public Sub BasesConvert()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            
            'Установить стандартный курсор мыши над формой "Preprocessors"
            '  - это приводит к невидимой загрузке формы
    frmPreprocessors.MousePointer = 0
            
            'Вызов процедуры обработки события "cmdProtocolBase_Click"
    Call cmdProtocolBase_Click
            'Вызов процедуры обработки события "cmdBookKeeperBase_Click"
    Call cmdBookKeeperBase_Click

End Sub
            
            'Обработка события "Protocol Base"
Private Sub cmdProtocolBase_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Количество строк в "Базе Протокола"
Dim lngProtocolBaseCount As Long
            'Номер файла Архива
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола" и DUMMY файла
Dim lngRecordLen As Long
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer
            'Полное имя DUMMY файла (с указанием "пути" к нему)
Dim strDummyFileName As String
            'Текущий номер строки таблицы DUMMY файла
Dim lngRowDummy As Long
            'Полное имя папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Номер дня (обратный отсчет, начиная с текущего дня),
            '  который просматривается системой при копировании
            '  Архивов Препроцессоа в DUMMY файл
Dim intDayArchive As Integer
            'Количество строк в копируемом файле (Архиве или "TableProtocol")
Dim intRowQuan As Integer
            'Текущий номер строки копируемого Архива
            '   или таблицы "TableProtocol"
Dim intRowNumArchive As Integer

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            'Полное имя файла "Таблица протокола "(с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
    gFileDummy = FreeFile
            'Полное имя DUMMY файла (с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            'Начальная позиция в полном имени DUMMY файла(за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить "старый" DUMMY файл, если он существует
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
    
            'Обработка ошибок
    On Error GoTo UnDefError
            'Открыть DUMMY файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            'Текущий номер  свободной строки DUMMY файла
    gDummyRowNum = 1
            
            'Выбраны все Препроцессоры: "All"
    If Trim(cboPreprocessors.Text) = "All" Then
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                frmTableSystem.grdTableSystem.Col = 1
            
            'Установка "Календаря" на Текущую дату
                frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
                    frmTableCalendar.comCalendar.PreviousDay
                Next
            'Цикл по всем датам, начиная с Начальной даты
                For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                    frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\" + _
                    Trim(frmTableSystem.grdTableSystem.Text)
                    If frmTableCalendar.comCalendar.Day < 10 Then
                        strPathFileName = strPathFileName + "_0" + _
                        CStr(frmTableCalendar.comCalendar.Day)
                    Else
                        strPathFileName = strPathFileName + "_" + _
                        CStr(frmTableCalendar.comCalendar.Day)
                    End If
                    If frmTableCalendar.comCalendar.Month < 10 Then
                        strPathFileName = strPathFileName + "_0" + _
                        CStr(frmTableCalendar.comCalendar.Month)
                    Else
                        strPathFileName = strPathFileName + "_" + _
                        CStr(frmTableCalendar.comCalendar.Month)
                    End If
                    strPathFileName = strPathFileName + "_" + _
                    Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
                    If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                        intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Цикл по всем строкам Архива
                        For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                            Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                            WriteDummy
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл Архива
                        Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                        gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                    End If
            'Установка "Календаря" на Следующую дату
                    frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                    
                Next
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                frmTableSystem.grdTableSystem.Col = 0
            'Папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла Препроцессора в конец DUMMY файла
                If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" Препроцессора
                    intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                    intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" Препроцессора для
            '   произвольного доступа
                    Open strPathFileName For Random As intFileNum _
                    Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                    For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                        Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                        WriteDummy
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                    Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол Препроцессора"
                    gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                    gProtocol.strProtocStatus = "04 - Manager"
            'Время
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
                    gProtocol.strProtocReserve = "Protocol From Preproc"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                End If
            
            'Сформировать разделительную полосу
                gProtocol.strProtocName = "=========="
                gProtocol.strProtocPersonCode = "=========="
                gProtocol.strProtocStatus = "=========="
                gProtocol.strProtocTime = "=========="
                gProtocol.strProtocDate = "=========="
                gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
                WriteDummy
            
            'Определить действительный "путь" к каталогу выполняемой программы
                strPathFolderName = App.Path
                If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
                    strPathFolderName = Left(strPathFolderName, _
                    Len(strPathFolderName) - 1)
                End If
            
            End If
        Next
        
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   "Host Computer'a" (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла "Host Computer'a" DUMMY файла
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" "Host Computer'a" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" "Host Computer'a" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола" "Host Computer'a"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "Protocol From Host"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
            
            'Выбран один Препроцессор
    Else
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                frmTableSystem.grdTableSystem.Col = 1
            
            'Выбранный Препроцессор найден - продолжить
                If Trim(cboPreprocessors.Text) = _
                Trim(frmTableSystem.grdTableSystem.Text) Then
            
            'Установка "Календаря" на Текущую дату
                    frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                    For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
                        frmTableCalendar.comCalendar.PreviousDay
                    Next
            'Цикл по всем датам, начиная с Начальной даты
                    For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                        frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
                        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                            intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                            Open strPathFileName For Random As intFileNum _
                            Len = lngRecordLen
            'Цикл по всем строкам Архива
                            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                                WriteDummy
            'Разрешить прерывания для обработки различных событий
                                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                                frmPreprocessors.MousePointer = vbHourglass
                            Next
            'Закрыть файл Архива
                            Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                            gProtocol.strProtocStatus = "04 - Manager"
            'Время
                            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(frmTableSystem.grdTableSystem.Text)
                            If frmTableCalendar.comCalendar.Day < 10 Then
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_0" + _
                                CStr(frmTableCalendar.comCalendar.Day)
                            Else
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_" + _
                                CStr(frmTableCalendar.comCalendar.Day)
                            End If
                            If frmTableCalendar.comCalendar.Month < 10 Then
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_0" + _
                                CStr(frmTableCalendar.comCalendar.Month)
                            Else
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_" + _
                                CStr(frmTableCalendar.comCalendar.Month)
                            End If
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            'Установка "Календаря" на Следующую дату
                        frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                    frmTableSystem.grdTableSystem.Col = 0
            'Папка-файл Препроцессора (с полным путем к ней)
                    strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла Препроцессора в конец DUMMY файла
                    If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" Препроцессора
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                        intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" Препроцессора для
            '   произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                        For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                            Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                            WriteDummy
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                        Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол Препроцессора"
                        gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
                        gProtocol.strProtocReserve = "Protocol From Preproc"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                    
                    End If
            'Завершить цикл
                    Exit For
                End If
            End If
        Next
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
            
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFolderName = App.Path
        If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
            strPathFolderName = Left(strPathFolderName, _
            Len(strPathFolderName) - 1)
        End If
    
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   "Host Computer'a" (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла "Host Computer'a" DUMMY файла
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" "Host Computer'a" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" "Host Computer'a" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола" "Host Computer'a"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "Protocol From Host"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
    
    End If
            
            'Определить действительный "путь" к каталогу
            '  выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            'Установка свойств элемента "Data" доступа к "Базе Протокола"
    frmDemo.datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
    frmDemo.datBase.RecordSource = "Protocol"
            
            'Определить количество записей в "Базе Протокола"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngProtocolBaseCount = frmDemo.datBase.Recordset.RecordCount
            'Обновить "Базу Протокола"
    frmDemo.datBase.Recordset.MoveFirst
            'Цикл по всем строкам DUMMY файла
    For lngRowDummy = 1 To gDummyRowNum - 1 Step 1
            'Разрешить прерывания для обработки различных событий
        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
        frmPreprocessors.MousePointer = vbHourglass
            'Читать строку DUMMY файла в буфер
        Get gFileDummy, lngRowDummy, gProtocol
            'Обновить текущую запись "Базы Протокола"
        frmDemo.datBase.Recordset.Edit
        frmDemo.datBase.Recordset.Fields("Name").Value = gProtocol.strProtocName
        frmDemo.datBase.Recordset.Fields("CodeOrPassword").Value = _
        gProtocol.strProtocPersonCode
        frmDemo.datBase.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
        frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
        frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
        frmDemo.datBase.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
        frmDemo.datBase.Recordset.Update
            'Не последняя запись старой "Базы Протокола"
        If lngRowDummy < lngProtocolBaseCount Then
            frmDemo.datBase.Recordset.MoveNext
            'Последняя запись старой "Базы Протокола"
        Else
            frmDemo.datBase.Recordset.AddNew
            frmDemo.datBase.Recordset.Update
            frmDemo.datBase.Recordset.MoveNext
        End If
    Next
            'Удаление одной лишней записи из  "Базы Протокола"
    If lngRowDummy > lngProtocolBaseCount Then
        frmDemo.datBase.Recordset.Delete
            'Удаление лишних записей из  "Базы Протокола"
    Else
        For lngRowDummy = lngRowDummy To lngProtocolBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
        Next
    End If
            
            'Протоколирование события - "Формирование Базы Протокола"
    gProtocol.strProtocName = "ProtocolBase"
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
    gProtocol.strProtocReserve = "Creation"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol
    
    GoTo EndProcedure
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            'Закрыть DUMMY файл
    Close gFileDummy
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Обработка события "BookKeeper Base"
Private Sub cmdBookKeeperBase_Click()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Номер файла Архива
Dim intFileNum As Integer
            'Длина строки "Таблицы протокола" и DUMMY файла
Dim lngRecordLen As Long
            'Позиция символа "\" в полном имени файла
Dim intSymbPos As Integer
            'Полное имя DUMMY файла (с указанием "пути" к нему)
Dim strDummyFileName As String
            'Текущий номер строки таблицы DUMMY файла
Dim lngRowDummy As Long
            'Полное имя папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Номер дня (обратный отсчет, начиная с текущего дня),
            '  который просматривается системой при копировании
            '  Архивов Препроцессоа в DUMMY файл
Dim intDayArchive As Integer
            'Количество строк в копируемом файле (Архиве или "TableProtocol")
Dim intRowQuan As Integer
            'Текущий номер строки копируемого Архива
            '   или таблицы "TableProtocol"
Dim intRowNumArchive As Integer
            'Количество строк в "Базе Бухгалтерии"
Dim lngBookKeepingBaseCount As Long
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
Dim lngBookKeepingRowNum As Long

            'Если курсор мыши = "Песочные часы", то выйти
    If frmPreprocessors.MousePointer = vbHourglass Then
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If

            'Если Препроцессоры отсутствуют, то выйти
    If cboFileName.Text = "" Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            'Убрать с экрана форму
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            'Полное имя файла "Таблица протокола "(с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Вычислить длину записи (строки) "Таблицы протокола"
    lngRecordLen = Len(gProtocol)
            'Получить свободный номер файла
    gFileDummy = FreeFile
            'Полное имя DUMMY файла (с указанием "пути" к нему)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            'Начальная позиция в полном имени DUMMY файла(за символами "C:\")
    intSymbPos = 4
            'Найти начальную позицию собственно имени файла
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            'Удалить "старый" DUMMY файл, если он существует
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            'Обработка ошибок
    On Error GoTo UnDefError
            'Открыть DUMMY файл для произвольного доступа или
            '   создать его, если он не существует
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            'Текущий номер  свободной строки DUMMY файла
    gDummyRowNum = 1
            
            'Выбраны все Препроцессоры: "All"
    If Trim(cboPreprocessors.Text) = "All" Then
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                frmTableSystem.grdTableSystem.Col = 1
            
            'Установка "Календаря" на Текущую дату
                frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
                    frmTableCalendar.comCalendar.PreviousDay
                Next
            'Цикл по всем датам, начиная с Начальной даты
                For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                    frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\" + _
                    Trim(frmTableSystem.grdTableSystem.Text)
                    If frmTableCalendar.comCalendar.Day < 10 Then
                        strPathFileName = strPathFileName + "_0" + _
                        CStr(frmTableCalendar.comCalendar.Day)
                    Else
                        strPathFileName = strPathFileName + "_" + _
                        CStr(frmTableCalendar.comCalendar.Day)
                    End If
                    If frmTableCalendar.comCalendar.Month < 10 Then
                        strPathFileName = strPathFileName + "_0" + _
                        CStr(frmTableCalendar.comCalendar.Month)
                    Else
                        strPathFileName = strPathFileName + "_" + _
                        CStr(frmTableCalendar.comCalendar.Month)
                    End If
                    strPathFileName = strPathFileName + "_" + _
                    Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
                    If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                        intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Цикл по всем строкам Архива
                        For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                            Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                            WriteDummy
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл Архива
                        Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                        gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                    End If
            'Установка "Календаря" на Следующую дату
                    frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                    
                Next
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                frmTableSystem.grdTableSystem.Col = 0
            'Папка-файл Препроцессора (с полным путем к ней)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла Препроцессора в конец DUMMY файла
                If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" Препроцессора
                    intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                    intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" Препроцессора для
            '   произвольного доступа
                    Open strPathFileName For Random As intFileNum _
                    Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                    For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                        Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                        WriteDummy
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                    Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол Препроцессора"
                    gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                    gProtocol.strProtocStatus = "04 - Manager"
            'Время
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
                    gProtocol.strProtocReserve = "Protocol From Preproc"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                End If
            
            'Сформировать разделительную полосу
                gProtocol.strProtocName = "=========="
                gProtocol.strProtocPersonCode = "=========="
                gProtocol.strProtocStatus = "=========="
                gProtocol.strProtocTime = "=========="
                gProtocol.strProtocDate = "=========="
                gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
                WriteDummy
            
            'Определить действительный "путь" к каталогу выполняемой программы
                strPathFolderName = App.Path
                If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
                    strPathFolderName = Left(strPathFolderName, _
                    Len(strPathFolderName) - 1)
                End If
            
            End If
        Next
        
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   "Host Computer'a" (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла "Host Computer'a" DUMMY файла
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" "Host Computer'a" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" "Host Computer'a" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола" "Host Computer'a"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "Protocol From Host"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
            
            'Выбран один Препроцессор
    Else
            'Цикл по всем нефиксированным строкам "Системной таблицы"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            'Текущая строка "Системной таблицы"
            frmTableSystem.grdTableSystem.Row = intRowNum
            'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
            'Тип="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                frmTableSystem.grdTableSystem.Col = 1
            
            'Выбранный Препроцессор найден - продолжить
                If Trim(cboPreprocessors.Text) = _
                Trim(frmTableSystem.grdTableSystem.Text) Then
            
            'Установка "Календаря" на Текущую дату
                    frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
                    For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
                        frmTableCalendar.comCalendar.PreviousDay
                    Next
            'Цикл по всем датам, начиная с Начальной даты
                    For intDayArchive = 1 To gDayNum + 1 Step 1
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
                        frmTableSystem.grdTableSystem.Col = 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
                        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                            intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                            Open strPathFileName For Random As intFileNum _
                            Len = lngRecordLen
            'Цикл по всем строкам Архива
                            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                                WriteDummy
            'Разрешить прерывания для обработки различных событий
                                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                                frmPreprocessors.MousePointer = vbHourglass
                            Next
            'Закрыть файл Архива
                            Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                            gProtocol.strProtocStatus = "04 - Manager"
            'Время
                            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                            gProtocol.strProtocReserve = _
                            Trim(frmTableSystem.grdTableSystem.Text)
                            If frmTableCalendar.comCalendar.Day < 10 Then
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_0" + _
                                CStr(frmTableCalendar.comCalendar.Day)
                            Else
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_" + _
                                CStr(frmTableCalendar.comCalendar.Day)
                            End If
                            If frmTableCalendar.comCalendar.Month < 10 Then
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_0" + _
                                CStr(frmTableCalendar.comCalendar.Month)
                            Else
                                gProtocol.strProtocReserve = _
                                Trim(gProtocol.strProtocReserve) + "_" + _
                                CStr(frmTableCalendar.comCalendar.Month)
                            End If
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            'Установка "Календаря" на Следующую дату
                        frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
                        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                    frmTableSystem.grdTableSystem.Col = 0
            'Папка-файл Препроцессора (с полным путем к ней)
                    strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   Препроцессора (с указанием "пути" к нему)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла Препроцессора в конец DUMMY файла
                    If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" Препроцессора
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                        intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" Препроцессора для
            '   произвольного доступа
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" Препроцессора
                        For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" Препроцессора из файла в буфер
                            Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                            WriteDummy
            'Разрешить прерывания для обработки различных событий
                            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            'Закрыть файл "Таблицы протокола" Препроцессора
                        Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол Препроцессора"
                        gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                        gProtocol.strProtocStatus = "04 - Manager"
            'Время
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
                        gProtocol.strProtocReserve = "Protocol From Preproc"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                        frmDemo.WriteProtocol
                    
                    End If
            'Завершить цикл
                    Exit For
                End If
            End If
        Next
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
            
            'Определить действительный "путь" к каталогу выполняемой программы
        strPathFolderName = App.Path
        If Right(strPathFolderName, 1) = "\" Then
            'Полное имя папки "Host Computera" для DUMMY файла
            '  (с указанием "пути" к ней)
            strPathFolderName = Left(strPathFolderName, _
            Len(strPathFolderName) - 1)
        End If
    
            'Установка "Календаря" на Текущую дату
        frmTableCalendar.comCalendar.Today
            
            'Цикл по всем датам, начиная с Текущей даты
        For intDayArchive = 1 To gDayNum Step 1
            'Установка "Календаря" на Предыдущую дату
            frmTableCalendar.comCalendar.PreviousDay
        Next
            'Цикл по всем датам, начиная с Начальной даты
        For intDayArchive = 1 To gDayNum + 1 Step 1
            'Полное имя копируемого Архива (с указанием "пути" к нему)
            strPathFileName = strPathFolderName + "\" + Trim(gHost)
            If frmTableCalendar.comCalendar.Day < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Day)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Day)
            End If
            If frmTableCalendar.comCalendar.Month < 10 Then
                strPathFileName = strPathFileName + "_0" + _
                CStr(frmTableCalendar.comCalendar.Month)
            Else
                strPathFileName = strPathFileName + "_" + _
                CStr(frmTableCalendar.comCalendar.Month)
            End If
            strPathFileName = strPathFileName + "_" + _
            Right(CStr(frmTableCalendar.comCalendar.Year), 2)

            'Файл Архива имеется
            If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в Архиве
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
                intFileNum = FreeFile
            'Открыть файл Архива для произвольного доступа
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам Архива
                For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку Архива из файла в буфер
                    Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                    WriteDummy
            'Разрешить прерывания для обработки различных событий
                    DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            'Закрыть файл Архива
                Close intFileNum
                            
            'Протоколирование события - "Копирование Архива в DUMMY файл"
                gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
                gProtocol.strProtocStatus = "04 - Manager"
            'Время
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Формирование Примечания
                gProtocol.strProtocReserve = Trim(gHost)
                If frmTableCalendar.comCalendar.Day < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Day)
                End If
                If frmTableCalendar.comCalendar.Month < 10 Then
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_0" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                Else
                    gProtocol.strProtocReserve = _
                    Trim(gProtocol.strProtocReserve) + "_" + _
                    CStr(frmTableCalendar.comCalendar.Month)
                End If
                gProtocol.strProtocReserve = _
                Trim(gProtocol.strProtocReserve) + "_" + _
                Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            'Установка "Календаря" на Следующую дату
            frmTableCalendar.comCalendar.NextDay
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            'Полное имя копируемого файла таблицы "TableProtocol"
            '   "Host Computer'a" (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            'Файл имеется - копирование файла таблицы "TableProtocol"
            '    из папки-файла "Host Computer'a" DUMMY файла
        If (FSO.FileExists(strPathFileName)) Then
            'Количество строк в "Таблице протокола" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            'Получить свободный номер файла
            intFileNum = FreeFile
            'Открыть файл "Таблицы протокола" "Host Computer'a" для
            '   произвольного доступа
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            'Цикл по всем строкам "Таблицы протокола" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            'Читать строку "Таблицы протокола" "Host Computer'a" из файла в буфер
                Get intFileNum, intRowNumArchive, gProtocol
            'Записать строку в DUMMY файл
                WriteDummy
            'Разрешить прерывания для обработки различных событий
                DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            'Закрыть файл "Таблицы протокола" "Host Computer'a"
            Close intFileNum
                        
            'Протоколирование события - "Копировать Протокол "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
            gProtocol.strProtocReserve = "Protocol From Host"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            'Сформировать разделительную полосу
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            'Записать строку в DUMMY файл
        WriteDummy
    
    End If
        
            'Определить действительный "путь" к каталогу
            '  выполняемой программы
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            'Установка свойств элемента "Data" доступа к "Базе Бухгалтерии"
    frmDemo.datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
    frmDemo.datBase.RecordSource = "BookKeeping"
            
            'Определить количество записей в "Базе Бухгалтерии"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngBookKeepingBaseCount = frmDemo.datBase.Recordset.RecordCount
            'Обновить "Базу Бухгалтерии"
    frmDemo.datBase.Recordset.MoveFirst
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
    lngBookKeepingRowNum = 0
    For lngRowDummy = 0 To gDummyRowNum - 1 Step 1
            'Разрешить прерывания для обработки различных событий
        DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
        frmPreprocessors.MousePointer = vbHourglass
            'Создать первую обязательную "фиктивную" запись
        If lngRowDummy = 0 Then
            'Отредактировать текущую запись "Базы Бухгалтерии"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = "Fiktive Record"
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = "0000000000000000"
            frmDemo.datBase.Recordset.Fields("Status").Value = "00"
            frmDemo.datBase.Recordset.Fields("Time").Value = "00:00:00AM"
            frmDemo.datBase.Recordset.Fields("Date").Value = "01.01.2000"
            'Обновление записи в "Базе Бухгалтерии"
            frmDemo.datBase.Recordset.Update
        Else
            'Читать строку DUMMY файла в буфер
            Get gFileDummy, lngRowDummy, gProtocol
            'Отредактировать текущую запись "Базы Бухгалтерии"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = gProtocol.strProtocName
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = gProtocol.strProtocPersonCode
            frmDemo.datBase.Recordset.Fields("Status").Value = Left(Trim(gProtocol.strProtocStatus), 2)
            frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            'Событие протокола:
            '                                  - Вход/Выход ("18"/"19") Служащего или
            '                                  - АвтоРегистрация ("16) Служащего или
            '                                  - АвтоУдаление ("17") Служащего или
            '                                  - Регистрация ("12") платного Клиента Автостоянки или
            '                                  - Исключение ("13") платного Клиента Автостоянки
            '                                  - Регистрация ("14") платного Посетителя Предприятия или
            '                                  - Исключение ("15") платного Посетителя Предприятия
            If ((frmDemo.datBase.Recordset.Fields("Status").Value = "00" Or _
            frmDemo.datBase.Recordset.Fields("Status").Value = "01") And _
            (Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Or _
                Right(Trim(gProtocol.strProtocReserve), 6) = "Output") Or _
            (frmDemo.datBase.Recordset.Fields("Status").Value = "01") And _
            (Trim(gProtocol.strProtocReserve) = "AutoRegistration" Or _
                Trim(gProtocol.strProtocReserve) = "AutoDelete") Or _
                (frmDemo.datBase.Recordset.Fields("Status").Value = "05" Or _
                frmDemo.datBase.Recordset.Fields("Status").Value = "06") And _
                (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Or _
                Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark") Or _
                (frmDemo.datBase.Recordset.Fields("Status").Value = "08" Or _
                frmDemo.datBase.Recordset.Fields("Status").Value = "09") And _
                (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Or _
                Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce")) And _
            Left(gProtocol.strProtocName, 1) <> "@" Then
            'Событие  - АвтоРегистрация Служащего (Установить признак
            '  в "Базе Бухгалтерии")
                If Trim(gProtocol.strProtocReserve) = "AutoRegistration" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "16"
            'Событие  - АвтоУдаление Служащего (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Trim(gProtocol.strProtocReserve) = "AutoDelete" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "17"
            'Событие  - Вход Служащего на Предприятие (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "18"
            'Событие  - Выход Служащего с Предприятия (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 6) = "Output" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "19"
            'Событие  - Регистрация Клиента Автостоянки (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "12"
            'Событие  - Исключение Клиента Автостоянки (Установить признак
            '  в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "13"
            'Событие  - Регистрация Посетителя Предприятия (Установить
            '  признак в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "14"
            'Событие  - Исключение Посетителя Предприятия (Установить
            '  признак в "Базе Бухгалтерии")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "15"
                End If
            'Обновление записи в "Базе Бухгалтерии"
                frmDemo.datBase.Recordset.Update
            'Текущий номер отредактированной записи "Базы Бухгалтерии"
                lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            'Не последняя запись старой "Базы Бухгалтерии"
                If lngBookKeepingRowNum < lngBookKeepingBaseCount Then
                    frmDemo.datBase.Recordset.MoveNext
            'Последняя запись старой "Базы Бухгалтерии"
                Else
                    frmDemo.datBase.Recordset.AddNew
                    frmDemo.datBase.Recordset.Update
                    frmDemo.datBase.Recordset.MoveNext
                End If
            End If
        End If
    Next
            'Текущий номер записи "Базы Бухгалтерии"
    lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            'Удаление одной лишней записи из  "Базы Бухгалтерии"
    If lngBookKeepingRowNum > lngBookKeepingBaseCount Then
        frmDemo.datBase.Recordset.Delete
            'Удаление лишних записей из  "Базы Бухгалтерии",
            '  кроме единственной
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum = 1 Then
        frmDemo.datBase.Recordset.MoveFirst
        frmDemo.datBase.Recordset.MoveNext
        For lngBookKeepingRowNum = 2 To lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
        Next
            'Удаление лишних записей из  "Базы Бухгалтерии"
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum <> 1 Then
        For lngBookKeepingRowNum = lngBookKeepingRowNum To _
        lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            'Разрешить прерывания для обработки различных событий
            DoEvents
            'Изменить стандартный курсор мыши  на "Песочные часы"
            frmPreprocessors.MousePointer = vbHourglass
        Next
    End If
            
            'Протоколирование события - "Формирование Базы Протокола"
    gProtocol.strProtocName = "BookKeeperBase"
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечание
    gProtocol.strProtocReserve = "Creation"

            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol
            
    GoTo EndProcedure
            'Неопределенная ошибка
UnDefError:
            'Издать звуковой сигнал
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            'Закрыть DUMMY файл
    Close gFileDummy
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            'Убрать с экрана форму
    frmPreprocessors.Hide

End Sub

            'Копирование архива из Препроцессора в "Host Computer"
Public Function ArchiveCopy(ByVal strMessage As String)
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer
            'Полное имя копируемого файла (с указанием "пути" к нему)
Dim strPathFileName As String
            'Полное имя файла-копии (с указанием "пути" к нему)
Dim strHostFileName As String
            'Полное имя выбираемой папки-файла (с указанием "пути" к ней)
Dim strPathFolderName As String
            'Объект "FileSystemObject" - "Файловая Система"
Dim FSO As Variant
            'Позиция символа "_" в имени файла
Dim intSymbPos As Integer

            'Начальная позиция в полном имени файла (за символами "Archive ")
    intSymbPos = 9
            'Найти конечную позицию собственно имени файла
    If InStr(intSymbPos, strMessage, "_") <> 0 Then
        intSymbPos = InStr(intSymbPos, strMessage, "_")
    Else
        intSymbPos = Len(strMessage)
    End If
            
            'Текущий столбец "Системной таблицы" = 2 (Тип)
    frmTableSystem.grdTableSystem.Col = 2
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = intRowNum
            'Тип="03" - Preprocessor (Препроцессор)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 1 (Имя Препроцессора)
            frmTableSystem.grdTableSystem.Col = 1
            'Требуемое имя Препроцессора
            If Trim(frmTableSystem.grdTableSystem.Text) = _
            Mid(strMessage, 9, intSymbPos - 9) Then
            'Текущий столбец "Системной таблицы" = 0 (Объект)
                frmTableSystem.grdTableSystem.Col = 0
            'Полное имя папки-файла Препроцессора
            '   (с указанием "пути" к ней) или "Whole"
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
                Exit For
            End If
            frmTableSystem.grdTableSystem.Col = 2
        End If
    Next
            
            'Создать объект "FSO" - "Файловая система"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            'Определить действительный "путь" к каталогу выполняемой программы
    strHostFileName = App.Path
    If Right(strHostFileName, 1) <> "\" Then
            'Полное имя папки "Host Computera" для файла-копии
            '  Препроцессора(с указанием "пути" к ней)
        strHostFileName = strHostFileName + "\"
    End If
    
            'Изменить стандартный курсор мыши  на "Песочные часы"
    frmPreprocessors.MousePointer = vbHourglass
            
            'Проверка существования папки-файла Препроцессора
    On Error GoTo UnExist
            'Папка-файл имеется - продолжить
    If (FSO.FolderExists(strPathFolderName)) Then
            'Протоколирование события - "Сгрузить Архивы Препроцессора"
        gProtocol.strProtocName = strPathFolderName
            'Системный пароль
        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
        gProtocol.strProtocStatus = "04 - Manager"
            'Время
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
        gProtocol.strProtocReserve = "DownLoad Archives"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
        frmDemo.WriteProtocol
                    
        On Error GoTo CopyingError
                    
            'Полное имя копируемого Архива Препроцессора
            '  (с указанием "пути" к нему)
        strPathFileName = strPathFolderName + "\" + _
        Mid(strMessage, 9)

        If (FSO.FileExists(strPathFileName)) Then
            'Файл имеется - копирование Архива в "Host Computer"
            FSO.CopyFile strPathFileName, strHostFileName
                        
            'Протоколирование события - "Копирование Архива Препроцессора"
            gProtocol.strProtocName = "Copy Archive"
            'Системный пароль
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
            gProtocol.strProtocStatus = "04 - Manager"
            'Время
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
            gProtocol.strProtocReserve = Mid(strMessage, 9, intSymbPos - 9)
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
            frmDemo.WriteProtocol
        End If
                
            'Папка-файл отсутствует - выход из процедуры с сообщением
    Else
        GoTo UnExist
    End If
                
    GoTo EndProcedure

UnExist:
            'Протоколирование события - "Ошибка при копировании Архива"
    gProtocol.strProtocName = strPathFolderName
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = "DownLoad Archives Err"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol
    GoTo EndProcedure

CopyingError:
            'Протоколирование события - "Ошибка при копировании Архива"
    gProtocol.strProtocName = "Copy Archive"
            'Системный пароль
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            'Статус
    gProtocol.strProtocStatus = "04 - Manager"
            'Время
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            'Дата
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            'Примечания
    gProtocol.strProtocReserve = _
    Trim(gProtocol.strProtocReserve) + " Err"
            'Записать строку в файл "Таблицы протокола" "Host Computer'a"
    frmDemo.WriteProtocol

EndProcedure:
            'Восстановить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0

End Function

            'Загрузка формы
Private Sub Form_Load()
            'Текущий номер строки таблицы "TableSystem"
            '   в "Host Computer'e"
Dim intRowNum As Integer

            'Установить стандартный курсор мыши
    frmPreprocessors.MousePointer = 0
            
            'Текущий столбец "Системной таблицы" = 2 (Тип)
    frmTableSystem.grdTableSystem.Col = 2
            'Цикл по всем нефиксированным строкам "Системной таблицы"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            'Текущая строка "Системной таблицы"
        frmTableSystem.grdTableSystem.Row = intRowNum
            'Тип="03" - Preprocessor (Препроцессор)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            'Текущий столбец "Системной таблицы" = 0 (Объект)
            frmTableSystem.grdTableSystem.Col = 0
            'Заполнить список комбинированного поля "cboFileName"
            If cboFileName.ListCount = 0 Then
                cboFileName.AddItem "Whole"
                cboPreprocessors.AddItem "All"
            End If
            cboFileName.AddItem _
            frmTableSystem.grdTableSystem.Text
            'Имя Препроцессора локальной сети
            cboPreprocessors.AddItem _
            gSocketNet(cboPreprocessors.ListCount)
        'Текущий столбец "Системной таблицы" = 2 (Тип)
            frmTableSystem.grdTableSystem.Col = 2
        End If
    Next
            'Блокировать возможную ошибку
    On Error GoTo UnLoad
            'Выбрать первый элемент списка: "Все Препроцессоры"
    cboFileName.ListIndex = 0
    cboPreprocessors.ListIndex = 0
UnLoad:

End Sub

            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            'Процедура записи строки в DUMMY файл
Public Sub WriteDummy()

            'Записать строку в файл "Таблицы протокола"
    Put gFileDummy, gDummyRowNum, gProtocol
            'Номер следующей свободной строки DUMMY файла
    gDummyRowNum = gDummyRowNum + 1
    
End Sub

