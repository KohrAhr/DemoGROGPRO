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
            
            '��� �������� ��� ��������� � �������� ������� "Mutex"
Dim lngRetCode As Long

            '�������� (�����������) ������� ActiveX.EXE
Private Sub Class_Initialize()
            
            ' �������� ������� "Mutex"
    gMutex = CreateMutex(0&, False, "Mutex")
    If gMutex = 0 Then
        Class_Terminate
    End If

End Sub

            '���������� ������� ActiveX.EXE
Private Sub Class_Terminate()
            
            '��������� ��� �����
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
    
            ' ���������� ������ �� ������ "FlexGrid" ("������� ������")
    Set gTablePerson = Nothing
            ' ���������� ������ �� ������ ActiveX.EXE
    Set objTablePerson = Nothing

            ' ���������� ������ "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            
            'The decrement of the accessable attribute
Private Sub ITablePerson_AccessMinus()
            
            '����� ������������ ������� "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    frmTablePerson.lblAccess.Caption = _
    frmTablePerson.lblAccess.Caption - 1
    If frmTablePerson.lblAccess.Caption < 0 Then _
    frmTablePerson.lblAccess.Caption = 0

            '���������� ������ "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            
            'The increment of the accessable attribute
Private Sub ITablePerson_AccessPlus()
            
            '����� ������������ ������� "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    frmTablePerson.lblAccess.Caption = _
    frmTablePerson.lblAccess.Caption + 1

            '���������� ������ "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Sub
            'Adds a new row to a FlexGrid control at run time
Private Sub ITablePerson_AddItem(ByVal Item As String)
            
            '���������� ������ � ����� "������� ������"
    frmTablePerson.grdTablePerson.AddItem Item

End Sub

            'Removes a row from a FlexGrid control at run time
Private Sub ITablePerson_RemoveItem(ByVal Index As Long)

            '���������� �������� ������ �� "������� ������"
    frmTablePerson.grdTablePerson.RemoveItem Index

End Sub

            'Returns attribute the accessable your object
Private Property Get ITablePerson_Access() As String
            
            '����� ������������ ������� "Mutex"
    lngRetCode = WaitForSingleObject(gMutex, 15000)
            
            '���������� ������� �������� ��������
            '  ����������� "������� ������"
    ITablePerson_Access = frmTablePerson.lblAccess.Caption

            '���������� ������ "Mutex"
    lngRetCode = ReleaseMutex(gMutex)

End Property

            'Returns the active cell in a FlexGrid
Private Property Get ITablePerson_Col() As Long
            
            '���������� ����� �������� ������� "������� ������"
    ITablePerson_Col = frmTablePerson.grdTablePerson.Col

End Property

            'Sets the active cell in a FlexGrid
Private Property Let ITablePerson_Col(ByVal RHs As Long)
            
            '���������� ����� �������� ������� "������� ������"
    frmTablePerson.grdTablePerson.Col = RHs

End Property

            'Sets the alignment of data in a column
Private Property Let ITablePerson_ColAlignment(ByVal RHs As Long)

End Property

            'Determines the total number of columns in a FlexGrid
Private Property Get ITablePerson_Cols() As Long
            
            '���������� ����� ���������� �������� � "������� ������"
    ITablePerson_Cols = frmTablePerson.grdTablePerson.Cols

End Property

            'Determines the total number of columns in a FlexGrid
Private Property Let ITablePerson_Cols(ByVal RHs As Long)
            
            '���������� ����� ���������� �������� � "������� ������"
    frmTablePerson.grdTablePerson.Cols = RHs

End Property

            'Determines the starting or ending column for a range of cells
Private Property Get ITablePerson_ColSel() As Long

            '���������� ����� ��������������� ������� "������� ������"
    ITablePerson_ColSel = frmTablePerson.grdTablePerson.ColSel

End Property

            'Determines the starting or ending column for a range of cells
Private Property Let ITablePerson_ColSel(ByVal RHs As Long)
            
            '���������� ����� ��������������� ������� "������� ������"
    frmTablePerson.grdTablePerson.ColSel = RHs

End Property

            'Determines the width of the specified column in Twips
Private Property Let ITablePerson_ColWidth(ByVal RHs As Long)

End Property

            'Returns the active cell in a FlexGrid
Private Property Get ITablePerson_Row() As Long
            
            '���������� ����� ������� ������ "������� ������"
    ITablePerson_Row = frmTablePerson.grdTablePerson.Row

End Property

            'Sets the active cell in a FlexGrid
Private Property Let ITablePerson_Row(ByVal RHs As Long)

            '���������� ����� ������� ������ "������� ������"
    frmTablePerson.grdTablePerson.Row = RHs

End Property

            'Determines the total number of rows in a FlexGrid
Private Property Get ITablePerson_Rows() As Long

            '���������� ���������� ����� � "������� ������"
    ITablePerson_Rows = frmTablePerson.grdTablePerson.Rows

End Property
            
            'Determines the total number of rows in a FlexGrid
Private Property Let ITablePerson_Rows(ByVal RHs As Long)

            '���������� ���������� ����� � "������� ������"
    frmTablePerson.grdTablePerson.Rows = RHs

End Property

            'Determines the starting or ending row for a range of cells
Private Property Get ITablePerson_RowSel() As Long

            '���������� ����� �������������� ������ "������� ������"
    ITablePerson_RowSel = frmTablePerson.grdTablePerson.RowSel

End Property

            'Determines the starting or ending row for a range of cells
Private Property Let ITablePerson_RowSel(ByVal RHs As Long)

            '���������� ����� �������������� ������ "������� ������"
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

