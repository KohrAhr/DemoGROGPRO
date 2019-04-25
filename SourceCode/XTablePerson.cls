VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTablePerson"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements ITablePerson
            
            'Код возврата при обращении к функциям объекта "Mutex"
Dim lngRetCode As Long

            'Создание (Конструктор) объекта ActiveX.EXE
Private Sub Class_Initialize()
            
            ' Создание объекта "Mutex"
    gMutex = CreateMutex(0&, False, "Mutex")
    If gMutex = 0 Then
        Class_Terminate
    End If

End Sub

            'Разрушение объекта ActiveX.EXE
Private Sub Class_Terminate()
            
            'Выгрузить все формы
    UnLoad frmDataAccessIn
    UnLoad frmDataAccessInfo
    UnLoad frmDataAccessOut
    UnLoad frmDataAccessServ
    UnLoad frmDataEmployeIn
    UnLoad frmDataEmployeInfo
    UnLoad frmDataEmployeOut
    UnLoad frmDataParkingIn
    UnLoad frmDataParkingInfo
    UnLoad frmDataParkingOut
    UnLoad frmDataParkingServ
    
    UnLoad frmDemo
    UnLoad frmGetFile
    UnLoad frmLease
    UnLoad frmPreprocessors
    UnLoad frmPrintPreview
    
    UnLoad frmTableCalendar
    UnLoad frmTablePerson
    UnLoad frmTableSystem
    UnLoad frmTableTerminal
    UnLoad frmTableTime
    
            ' Освободить ссылку на объект "FlexGrid" ("Таблица Персон")
    Set gTablePerson = Nothing
            ' Освободить ссылку на объект ActiveX.EXE
    Set objTablePerson = Nothing

            ' Освободить объект "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            
            'The decrement of the accessable attribute
Private Sub ITablePerson_AccessMinus()
            
            'Ждать освобождения объекта "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            'Уменьшить текущее значение атрибута
            '  доступности "Таблицы персон"
    frmTablePerson.lblAccess.Caption = _
    frmTablePerson.lblAccess.Caption - 1
    If frmTablePerson.lblAccess.Caption < 0 Then _
    frmTablePerson.lblAccess.Caption = 0

            'Освободить объект "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            
            'The increment of the accessable attribute
Private Sub ITablePerson_AccessPlus()
            
            'Ждать освобождения объекта "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            'Увеличить текущее значение атрибута
            '  доступности "Таблицы персон"
    frmTablePerson.lblAccess.Caption = _
    frmTablePerson.lblAccess.Caption + 1

            'Освободить объект "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            'Adds a new row to a FlexGrid control at run time
Private Sub ITablePerson_AddItem(ByVal Item As String)
            
            'Добавление строки в конец "Таблицы персон"
    frmTablePerson.grdTablePerson.AddItem Item

End Sub

            'Removes a row from a FlexGrid control at run time
Private Sub ITablePerson_RemoveItem(ByVal Index As Long)

            'Физическое удаление строки из "Таблицы персон"
    frmTablePerson.grdTablePerson.RemoveItem Index

End Sub

            'Returns attribute the accessable your object
Private Property Get ITablePerson_Access() As String
            
            'Ждать освобождения объекта "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            'Определить текущее значение атрибута
            '  доступности "Таблицы персон"
    ITablePerson_Access = frmTablePerson.lblAccess.Caption

            'Освободить объект "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Property

            'Returns the active cell in a FlexGrid
Private Property Get ITablePerson_Col() As Long
            
            'Определить номер текущего столбца "Таблицы персон"
    ITablePerson_Col = frmTablePerson.grdTablePerson.Col

End Property

            'Sets the active cell in a FlexGrid
Private Property Let ITablePerson_Col(ByVal RHs As Long)
            
            'Установить номер текущего столбца "Таблицы персон"
    frmTablePerson.grdTablePerson.Col = RHs

End Property

            'Sets the alignment of data in a column
Private Property Let ITablePerson_ColAlignment(ByVal RHs As Long)

End Property

            'Determines the total number of columns in a FlexGrid
Private Property Get ITablePerson_Cols() As Long
            
            'Определить общее количество столбцов в "Таблице персон"
    ITablePerson_Cols = frmTablePerson.grdTablePerson.Cols

End Property

            'Determines the total number of columns in a FlexGrid
Private Property Let ITablePerson_Cols(ByVal RHs As Long)
            
            'Установить общее количество столбцов в "Таблице персон"
    frmTablePerson.grdTablePerson.Cols = RHs

End Property

            'Determines the starting or ending column for a range of cells
Private Property Get ITablePerson_ColSel() As Long

            'Определить номер корректируемого столбца "Таблицы персон"
    ITablePerson_ColSel = frmTablePerson.grdTablePerson.ColSel

End Property

            'Determines the starting or ending column for a range of cells
Private Property Let ITablePerson_ColSel(ByVal RHs As Long)
            
            'Установить номер корректируемого столбца "Таблицы персон"
    frmTablePerson.grdTablePerson.ColSel = RHs

End Property

            'Determines the width of the specified column in Twips
Private Property Let ITablePerson_ColWidth(ByVal RHs As Long)

End Property

            'Returns the active cell in a FlexGrid
Private Property Get ITablePerson_Row() As Long
            
            'Определить номер текущей строки "Таблицы персон"
    ITablePerson_Row = frmTablePerson.grdTablePerson.Row

End Property

            'Sets the active cell in a FlexGrid
Private Property Let ITablePerson_Row(ByVal RHs As Long)

            'Установить номер текущей строки "Таблицы персон"
    frmTablePerson.grdTablePerson.Row = RHs

End Property

            'Determines the total number of rows in a FlexGrid
Private Property Get ITablePerson_Rows() As Long

            'Определить количество строк в "Таблице персон"
    ITablePerson_Rows = frmTablePerson.grdTablePerson.Rows

End Property
            
            'Determines the total number of rows in a FlexGrid
Private Property Let ITablePerson_Rows(ByVal RHs As Long)

            'Установить количество строк в "Таблице персон"
    frmTablePerson.grdTablePerson.Rows = RHs

End Property

            'Determines the starting or ending row for a range of cells
Private Property Get ITablePerson_RowSel() As Long

            'Определить номер корректируемой строки "Таблицы персон"
    ITablePerson_RowSel = frmTablePerson.grdTablePerson.RowSel

End Property

            'Determines the starting or ending row for a range of cells
Private Property Let ITablePerson_RowSel(ByVal RHs As Long)

            'Установить номер корректируемой строки "Таблицы персон"
    frmTablePerson.grdTablePerson.RowSel = RHs

End Property

            'Returns any extra data needed for your program
Private Property Get ITablePerson_Tag() As String

    ITablePerson_Tag = frmTablePerson.grdTablePerson.Tag
    
End Property

            'Stores any extra data needed for your program
Private Property Let ITablePerson_Tag(ByVal RHs As String)

    frmTablePerson.grdTablePerson.Tag = RHs
    
End Property

            'Returns the text contents of a cell
Private Property Get ITablePerson_Text() As String

    ITablePerson_Text = frmTablePerson.grdTablePerson.Text

End Property

            'Sets the text contents of a cell
Private Property Let ITablePerson_Text(ByVal RHs As String)

    frmTablePerson.grdTablePerson.Text = RHs

End Property

