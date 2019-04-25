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
            '������� ����� ������ ����� "frmPrintPreview"
Dim intRowPrintNum As Integer
            '���������� ����� �� ����� �������� ����� "frmPrintPreview"
Dim intRowPrintQuan As Integer
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "����� ���������")
Dim lngRowNum As Long
            '������� ����� ������� ������ ("TablePerson", "TableCalendar","TableProtocol,
            '  "TableSystem", "TableTime", "TableTerminal")
Dim intColNum As Integer
            '��������� ���� � "������� ���������"
Dim strDateFrom As String
            '���������� ������� � "������ ���������"
Dim lngArchivesRowNum As Long
             '����� ����� "����� ���������"
Dim intFileNum As Integer
           '���������� ����� � "���� ���������"
Dim intProtocolBaseCount As Integer
           '����� ������ ������ "��������� �������"
Dim strTableSystem(5) As String
           '����� ������ ������ "������� ������"
Dim strTablePerson(6) As String
            '����� ������ ������ "������� ���������"
Dim strTableCalendar(8) As String
           '����� ������ ������ "������� �������"
Dim strTableTime(3) As String
           '����� ������ ������ "������� ����������"
Dim strTableTerminal(4) As String
            
            '��������� �������� ����� "���� ���������"
Private Sub chkProtocol_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
            '������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
    
            '�������� �����
    frmPrintPreview.Cls
            '����� "���� ���������" ���������
    If chkProtocol.Value = 0 Then
            '������ "Find" �������� ��� �������
        cmdFind.MousePointer = 0
            '������� ��������� ������ "First"
        cmdFirst.Enabled = True
            '������� ��������� ������ "Last"
        cmdLast.Enabled = True
            '������� ��������� ������ "Previous"
        cmdPrevious.Enabled = True
            '������� ����������� ������ "Next"
        cmdNext.Enabled = False
            '������� ��������� ������ "PrintPage"
        cmdPrintPage.Enabled = True
            '������� ��������� ������ "Cancel"
        cmdCancel.Enabled = True
            '������� ��������� ������ "Find"
        cmdFind.Visible = False
            '������� ���������  ����� "PrintPage"
        chkPrintPage.Visible = False
            '������� ��������� ���� ����� "Name"
        txtName.Visible = False
            '������� ��������� ���� ����� "CodeOrPassword"
        txtCodeOrPassword.Visible = False
            '������� ���������� ���� ����� "Date"
        txtDateFrom.Visible = False
        txtDateTo.Visible = False
            '������� ���������� ����� ����� ����� "Date"
        lblDateFrom.Visible = False
        lblDateTo.Visible = False
            '������� ��������� ���� ����� "ReservOrNote"
        txtReservOrNote.Visible = False

            
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
            
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
                If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
        
            '����� "���� ���������" ��������
    Else
            
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
            '��������� ������� �������� "Data" ������� � "���� ���������"
        datProtocol.DatabaseName = strPathFileName + "ProtocolBase.mdb"
        datProtocol.RecordSource = "Protocol"
            
            '�������� ����������� ������ ����  �� "�������� ����"
        frmPrintPreview.MousePointer = vbHourglass
            '������ "Find" �������� ��� �������
        cmdFind.MousePointer = 0
            '������� ����������� ������ "First"
        cmdFirst.Enabled = False
            '������� ����������� ������ "Last"
        cmdLast.Enabled = False
            '������� ����������� ������ "Previous"
        cmdPrevious.Enabled = False
            '������� ����������� ������ "Next"
        cmdNext.Enabled = False
            '������� ����������� ������ "PrintPage"
        cmdPrintPage.Enabled = False
            '������� ����������� ������ "Cancel"
        cmdCancel.Enabled = False
        
            '���������� ���������� ������� � "���� ���������"
        datProtocol.Refresh
        datProtocol.Recordset.MoveLast
        intProtocolBaseCount = datProtocol.Recordset.RecordCount
            '�������� "���� ���������"
        datProtocol.Recordset.MoveFirst
            '���� �� ���� ������� "������� ���������"
        For lngRowNum = 1 To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, lngRowNum, gProtocol
            '���� "From"
            If lngRowNum = 1 Then
                txtDateFrom.Text = gProtocol.strProtocDate
                strDateFrom = txtDateFrom.Text
            End If
            '�������� ������� ������ "���� ���������"
            datProtocol.Recordset.Edit
            datProtocol.Recordset.Fields("Name").Value = gProtocol.strProtocName
            datProtocol.Recordset.Fields("CodeOrPassword").Value = _
            gProtocol.strProtocPersonCode
            datProtocol.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
            datProtocol.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            datProtocol.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            datProtocol.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
            datProtocol.Recordset.Update
            '�� ��������� ������ ������ "���� ���������"
            If lngRowNum < intProtocolBaseCount Then
                datProtocol.Recordset.MoveNext
            '��������� ������ ������ "���� ���������"
            Else
                datProtocol.Recordset.AddNew
                datProtocol.Recordset.Update
                datProtocol.Recordset.MoveNext
            End If
        Next
            '�������� ����� ������ ������ ��  "���� ���������"
        If lngRowNum > intProtocolBaseCount Then
            datProtocol.Recordset.Delete
            '�������� ������ ������� ��  "���� ���������"
        Else
            For lngRowNum = lngRowNum To intProtocolBaseCount Step 1
                datProtocol.Recordset.Delete
                datProtocol.Recordset.MoveNext
            Next
        End If
            '������������ ������� ����� ������ ������� "TableProtocol"
        lngRowNum = 1
            '������������ ����������� ������ ����
        frmPrintPreview.MousePointer = 0
            '���� "To"
        txtDateTo.Text = Format(Now, "dd/mm/yyyy")
            '������� ��������� ����� �����
        txtName.Text = ""
        txtCodeOrPassword.Text = ""
        txtReservOrNote.Text = ""
            '������� ������� ������ "Find"
        cmdFind.Visible = True
            '����������  ����� "PrintPage"
        chkPrintPage.Value = 0
            '������� �������  ����� "PrintPage"
        chkPrintPage.Visible = True
            '������� ������� ���� ����� "Name"
        txtName.Visible = True
            '������� ������� ���� ����� "CodeOrPassword"
        txtCodeOrPassword.Visible = True
            '������� �������� ���� ����� "Date"
        txtDateFrom.Visible = True
        txtDateTo.Visible = True
            '������� �������� ����� ����� ����� "Date"
        lblDateFrom.Visible = True
        lblDateTo.Visible = True
            '������� ������� ���� ����� "ReservOrNote"
        txtReservOrNote.Visible = True
    
            '���������� �������� ��� ������ ������
        frmPrintPreview.CurrentY = 1350
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = 7
             '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
   End If
    
End Sub

            '��������� ������� "Cancel"
Private Sub cmdCancel_Click()
            '������� ���������  ���� "����� ���������", ���� �� ������
    If intFileNum <> Empty Then Close intFileNum
            '������ � ������ �����
    frmPrintPreview.Hide

End Sub
            
            '��������� ������� ������� ������ "Find"
Private Sub cmdFind_Click()
            '����-���� ���������
Dim lngDateProtocol As Long
            '����-���� ���������
Dim lngDateFrom As Long
            '����-���� ��������
Dim lngDateTo As Long
            '������ "Find" �������� ��� �������
    If cmdFind.MousePointer = 0 Then
            '������� ����������� ������� �� ������ "Find"
        cmdFind.MousePointer = vbNoDrop
                '�������� ������ ��������� ���� "From"
        If Mid(Trim(txtDateFrom.Text), 3, 1) <> "." Or Mid(Trim(txtDateFrom.Text), 6, 1) <> "." _
        Or Len(Trim(txtDateFrom.Text)) <> 10 Then
            '��������� ������ �� ��������� ���� "DateFrom"
            txtDateFrom.SetFocus
            '�������� ������
            frmDemo.BeepSound
            '��������� ����� �� ���������
            Exit Sub
        End If
                '�������� ������ ��������� ���� "To"
        If Mid(Trim(txtDateTo.Text), 3, 1) <> "." Or Mid(Trim(txtDateTo.Text), 6, 1) <> "." _
        Or Len(Trim(txtDateTo.Text)) <> 10 Then
            '��������� ������ �� ��������� ���� "DateTo"
            txtDateTo.SetFocus
            '�������� ������
            frmDemo.BeepSound
            '��������� ����� �� ���������
            Exit Sub
        End If
            '�������� ����� � �������
        frmPrintPreview.Cls
        Printer.EndDoc
            '���������� �������� ��� ������ ������
        frmPrintPreview.CurrentY = 1350
        If chkPrintPage.Value = 1 Then Printer.CurrentY = 5
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = 7
            '������ ���������� �������� �� �����
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
             '������ ���������� �������� �� �������
        If chkPrintPage.Value = 1 Then Printer.Print Tab(3); "Name"; Tab(25); _
        "Code or Password"; Tab(55); "Status"; Tab(70); "Time"; Tab(95); "Date"; _
        Tab(100); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
        If chkPrintPage.Value = 1 Then Printer.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '����-���� ���������
        lngDateFrom = CLng(Mid(Trim(txtDateFrom.Text), 7, 4)) * 365 + _
        CLng(Mid(Trim(txtDateFrom.Text), 4, 2)) * 31 + _
        CLng(Mid(Trim(txtDateFrom.Text), 1, 2))
            '����-���� ��������
        lngDateTo = CLng(Mid(Trim(txtDateTo.Text), 7, 4)) * 365 + _
        CLng(Mid(Trim(txtDateTo.Text), 4, 2)) * 31 + _
        CLng(Mid(Trim(txtDateTo.Text), 1, 2))
            '���� �� ���� ������� "������� ���������"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, lngRowNum, gProtocol
            '��������� ������ - �������� ������ ���� � ���������
            If Val(Mid(gProtocol.strProtocDate, 7, 4)) = 0 Or _
            Val(Mid(gProtocol.strProtocDate, 4, 2)) = 0 Or _
            Val(Mid(gProtocol.strProtocDate, 1, 2)) = 0 Then
            '����-���� ��������� - ��������
                lngDateProtocol = lngDateFrom
            Else
            '����-���� ���������
            lngDateProtocol = CLng(Mid(gProtocol.strProtocDate, 7, 4)) * 365 + _
                CLng(Mid(gProtocol.strProtocDate, 4, 2)) * 31 + _
                CLng(Mid(gProtocol.strProtocDate, 1, 2))
            End If
            
            '������ "������� ���������" ������������� �������� ������
            If Len(Trim(txtName.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocName), _
            Trim(txtName.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            '������� ������ "������� ���������" �� �����
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            '������� ������ "������� ���������" �� �������
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            '������ "������� ���������" ������������� �������� ������
            ElseIf Len(Trim(txtCodeOrPassword.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocPersonCode), _
            Trim(txtCodeOrPassword.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            '������� ������ "������� ���������" �� �����
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            '������� ������ "������� ���������" �� �������
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            '������ "������� ���������" ������������� �������� ������
            ElseIf Len(Trim(txtReservOrNote.Text)) > 0 And _
            InStr(1, Trim(gProtocol.strProtocReserve), _
            Trim(txtReservOrNote.Text)) <> 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            '������� ������ "������� ���������"
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            '������� ������ "������� ���������" �� �������
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            '������ "������� ���������" ������������� �������� ������
            ElseIf Len(Trim(txtName.Text)) = 0 And Len(Trim(txtCodeOrPassword.Text)) = 0 _
            And Len(Trim(txtReservOrNote.Text)) = 0 And _
            lngDateProtocol >= lngDateFrom And lngDateProtocol <= lngDateTo Then
            '������� ������ "������� ���������"
                frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(115); gProtocol.strProtocReserve
            '������� ������ "������� ���������" �� �������
                If chkPrintPage.Value = 1 Then Printer.Print Tab(3); gProtocol.strProtocName; _
                Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
                Tab(70); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
                Tab(100); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
                intRowPrintNum = intRowPrintNum + 1
            End If
            
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������� �� ������ "Find"
                If lngRowNum < gProtocRowNum - 1 Then cmdFind.MousePointer = 0
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
            '������� ������ �� ������
        If chkPrintPage.Value = 1 Then Printer.EndDoc
    End If
    
End Sub

            '������ ������ "Next"
Private Sub cmdNext_Click()
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = lngRowNum + 1
    
            '�������� �����
    frmPrintPreview.Cls
            '������� ��������� ������ "Previous"
    cmdPrevious.Enabled = True
            '������� ����������� ������ "Next"
    cmdNext.Enabled = False

            
            '��������������� ������ "��������� �������" �� ����� "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "��������� �������"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            '������� ������ "��������� �������"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ������"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            '������ - ������ �����������
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            '������� ������ "������� ������"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(95); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ���������"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(95); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
        Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������ ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������ ���������"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            '������ ������ "������ ���������" �� ����� � �����
            Get intFileNum, lngRowNum, gProtocol
            '������� ������ "������ ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
                
                '��������������� ������ "������� �������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� �������"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� ������"
            frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            '������� ������ "������� �������"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ����������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ����������"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            '������� ������ "������� ����������"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
End If

End Sub
            
            '������ ������ "Previous"
Private Sub cmdPrevious_Click()
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "������ ���������")
    lngRowNum = lngRowNum - (intRowPrintQuan + intRowPrintNum - 5) + 1
    If intRowPrintNum <= intRowPrintQuan Then lngRowNum = lngRowNum - 1
    If lngRowNum < 1 Then lngRowNum = 1
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
    
            '�������� �����
    frmPrintPreview.Cls
            '������� ����������� ������ "Previous"
    If lngRowNum = 1 Then cmdPrevious.Enabled = False
            '������� ��������� ������ "Next"
    cmdNext.Enabled = True

            
            '��������������� ������ "��������� �������" �� ����� "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "��������� �������"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "��������� �������"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            '������� ������ "��������� �������"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ������"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            '������ - ������ �����������
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            '������� ������ "������� ������"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ���������"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
        Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������ ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������ ���������"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            '������ ������ "������ ���������" �� ����� � �����
            Get intFileNum, lngRowNum, gProtocol
            '������� ������ "������ ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
                
                '��������������� ������ "������� �������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� �������"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� ������"
            frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            '������� ������ "������� �������"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ����������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ����������"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            '������� ������ "������� ����������"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
    End If

End Sub
            
            '������ ������ "First"
Private Sub cmdFirst_Click()
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
            '����� ������ ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = 1
    
            '�������� �����
    frmPrintPreview.Cls
            '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
            '������� ����������� ������ "Next"
    cmdNext.Enabled = False

            
            '��������������� ������ "��������� �������" �� ����� "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "��������� �������"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            '������� ������ "��������� �������"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ������"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            '������ - ������ �����������
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            '������� ������ "������� ������"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ���������"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
        Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������ ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������ ���������"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            '������ ������ "������ ���������" �� ����� � �����
            Get intFileNum, lngRowNum, gProtocol
            '������� ������ "������ ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
                
                '��������������� ������ "������� �������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� �������"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� ������"
            frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            '������� ������ "������� �������"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ����������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ����������"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            '������� ������ "������� ����������"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
End If

End Sub
            
            '������ ������ "Last"
Private Sub cmdLast_Click()
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
    
            '�������� �����
    frmPrintPreview.Cls
            '������� ��������� ������ "Previous"
    cmdPrevious.Enabled = True
            '������� ����������� ������ "Next"
    cmdNext.Enabled = False
            
            
            '��������������� ������ "��������� �������" �� ����� "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "��������� �������"
    lngRowNum = frmTableSystem.grdTableSystem.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            '���� �� ���� ��������������� ������� "��������� �������"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            '������� ������ "��������� �������"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            '��������������� ������ "������� ������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "������� ������"
    lngRowNum = gTablePerson.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            '���� �� ���� ��������������� ������� "������� ������"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            '������ - ������ �����������
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            '������� ������ "������� ������"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "������� ���������"
    lngRowNum = frmTableCalendar.grdTableCalendar.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            '���� �� ���� ��������������� ������� "������� ���������"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "������� ���������"
        lngRowNum = gProtocRowNum + 3 - intRowPrintQuan
        If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
            cmdPrevious.Enabled = False
            lngRowNum = 1
        End If
            '���� �� ���� ������� ��������� �������� "������� ���������"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
        Next
    
            '��������������� ������ "������ ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            
            '����� ������ ������ ��������� �������� "������� ���������"
        lngRowNum = lngArchivesRowNum + 4 - intRowPrintQuan
        If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
            cmdPrevious.Enabled = False
            lngRowNum = 1
        End If
            '���� �� ���� ������� "������ ���������"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            '������ ������ "������ ���������" �� ����� � �����
            Get intFileNum, lngRowNum, gProtocol
            '������� ������ "������ ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
        Next
                
                '��������������� ������ "������� �������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "������� �������"
    lngRowNum = frmTableTime.grdTableTime.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            '���� �� ���� ��������������� ������� "������� �������"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� ������"
            frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            '������� ������ "������� �������"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
            '��������������� ������ "������� ����������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '����� ������ ������ ��������� �������� "������� ����������"
    lngRowNum = frmTableTerminal.grdTableTerminal.Rows + 3 - intRowPrintQuan
    If lngRowNum < 0 Then
                '������� ����������� ������ "Previous"
    cmdPrevious.Enabled = False
    lngRowNum = 1
    End If
            '���� �� ���� ��������������� ������� "������� ����������"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            '������� ������ "������� ����������"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
    Next
    
End If

End Sub
            '������ ������ "PrintPage"
Private Sub cmdPrintPage_Click()
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal")
    lngRowNum = lngRowNum - (intRowPrintNum - 3) + 1
    If intRowPrintNum <= intRowPrintQuan Then lngRowNum = lngRowNum - 1
    If lngRowNum < 1 Then lngRowNum = 1
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
    
            '�������� ������� �� "��������" ���������� ������
    Printer.EndDoc
    
            '������ "��������� �������"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
        Tab(70); "Index"; Tab(95); "Appendix"
            '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "��������� �������"
        For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "��������� �������"
            For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
                frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
                strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
            Next
            '������� �� ������ ������ "��������� �������"
            Printer.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
            Tab(55); strTableSystem(2); Tab(70); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            '������ "������� ������"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
        Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ������"
        For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
            For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
                gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
                strTablePerson(intColNum) = gTablePerson.Text
            Next
            '������ - ������ �����������
            If Left(Trim(strTablePerson(2)), 2) = "07" Or _
            Left(Trim(strTablePerson(2)), 2) = "05" Or _
            Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
                strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
            ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
            Left(Trim(strTablePerson(2)), 2) = "08" Or _
            Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
                strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
            End If
            '������� �� ������ ������ "������� ������"
            Printer.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
            Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
            Tab(115); strTablePerson(5)
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            '������ "������� ���������"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
        Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
        Tab(115); "Sunday"
            '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ���������"
        For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
            frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
            For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
                frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
                strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
            Next
            '������� �� ������ ������ "������� ���������"
            Printer.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
            Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
            Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
                '������ "������� ���������"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(75); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
           '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
        For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
            Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
            Printer.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(75); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
        
            '������ "������� �������"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� �������"
        For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
            frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
            For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
                frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
                strTableTime(intColNum) = frmTableTime.grdTableTime.Text
            Next
            '������� �� ������ ������ "������� �������"
            Printer.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); Tab(55); strTableTime(2)
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
            '������ "������� ����������"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
        Printer.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
        Tab(70); "Expander"
            '������� ������ ������
        Printer.Print
            '������� ����� ������ �� �������� ������
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ����������"
        For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
            frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
            For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
                frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
                strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
            Next
            '������� �� ������ ������ "������� ����������"
            Printer.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
            Tab(55); strTableTerminal(2); Tab(70); strTableTerminal(3)
            '������� ����� ������ �� �������� ������
            intRowPrintNum = intRowPrintNum + 1
            '�������� ������ ��������� - ��������� ������
            If intRowPrintNum > intRowPrintQuan Then Exit For
        Next
    
    End If
    
            '������ ��� ������ ������ ���
    Printer.EndDoc

End Sub

            '����� "������������"
Private Sub Form_Activate()
             '������ ��� ����� "����� ���������" (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������ ���������"
Dim lngRecordLen As Long
            '����� � ������ "������ ���������"
Dim lngFileLength As Long

            '������� ������� "�� ��������� - �������"
    Set Printer = Printers(0)

            '������� "������" ����� ����� "����� ���������"
    intFileNum = Empty
           '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = 1
            '���������� ����� �� ����� �������� ����� "frmPrintPreview"
    intRowPrintQuan = gRowPrintQuan
            '������� ����� ������ ������ ("TablePerson", "TableCalendar", "TableProtocol",
            '  "TableSystem", "TableTime", "TableTerminal", "����� ���������")
    lngRowNum = 1
            
            '�������� �����
    frmPrintPreview.Cls
            '��������������� ������ "��������� �������" �� ����� "frmPrintPreview"
    If frmPrintPreview.Tag = "TableSystem" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Objects"; Tab(25); "Cons.,Addr.,Term."; Tab(55); "Type"; _
    Tab(75); "Index"; Tab(95); "Appendix"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "��������� �������"
    For lngRowNum = lngRowNum To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = lngRowNum
            '�� ���� �������� "��������� �������"
        For intColNum = 0 To frmTableSystem.grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            frmTableSystem.grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������ "��������� �������"
            strTableSystem(intColNum) = frmTableSystem.grdTableSystem.Text
        Next
            '������� ������ "��������� �������"
        frmPrintPreview.Print Tab(3); strTableSystem(0); Tab(25); strTableSystem(1); _
        Tab(55); strTableSystem(2); Tab(75); strTableSystem(3); Tab(95); strTableSystem(4)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableSystem.grdTableSystem.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TablePerson" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "PersonCode"; Tab(55); "Status"; _
    Tab(75); "Time"; Tab(95); "Calendar"; Tab(115); "Reservation"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ������"
    For lngRowNum = lngRowNum To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = lngRowNum
            '�� ���� �������� "������� ������"
        For intColNum = 0 To gTablePerson.Cols - 1 Step 1
            '������� ������� "������� ������"
            gTablePerson.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ������"
            strTablePerson(intColNum) = gTablePerson.Text
        Next
            '������ - ������ �����������
        If Left(Trim(strTablePerson(2)), 2) = "07" Or _
        Left(Trim(strTablePerson(2)), 2) = "05" Or _
        Left(Trim(strTablePerson(2)), 2) = "06" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoParking || " + Right(Trim(strTablePerson(5)), 2)
            '������ - ���������� �����������
        ElseIf Left(Trim(strTablePerson(2)), 2) = "10" Or _
        Left(Trim(strTablePerson(2)), 2) = "08" Or _
        Left(Trim(strTablePerson(2)), 2) = "09" Then
            '������������ ������������ �������� � ���� ������ (����������)
            strTablePerson(5) = "AutoAccess || " + Right(Trim(strTablePerson(5)), 2)
        End If
            '������� ������ "������� ������"
        frmPrintPreview.Print Tab(3); strTablePerson(0); Tab(25); strTablePerson(1); _
        Tab(55); strTablePerson(2); Tab(75); strTablePerson(3); Tab(95); strTablePerson(4); _
        Tab(115); strTablePerson(5)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gTablePerson.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableCalendar*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Week Number"; Tab(25); "Monday"; Tab(40); "Tuesday"; _
    Tab(55); "Wednesday"; Tab(70); "Thursday"; Tab(85); "Friday"; Tab(100); "Saturday"; _
    Tab(115); "Sunday"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ���������"
    For lngRowNum = lngRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = lngRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ���������"
            strTableCalendar(intColNum) = frmTableCalendar.grdTableCalendar.Text
        Next
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); strTableCalendar(0); Tab(25); strTableCalendar(1); _
        Tab(40); strTableCalendar(2); Tab(55); strTableCalendar(3); Tab(70); strTableCalendar(4); _
        Tab(85); strTableCalendar(5); Tab(100); strTableCalendar(6); Tab(115); strTableCalendar(7)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableCalendar.grdTableCalendar.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
                '��������������� ������ "������� ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableProtocol" Then
            '������� ������� ����� "Connect Protocol Base"
    chkProtocol.Visible = True
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
    Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������� ���������"
    For lngRowNum = lngRowNum To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
        Get gProtocFileNum, lngRowNum, gProtocol
            '������� ������ "������� ���������"
        frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
        Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
        Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
        Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < gProtocRowNum - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������ ���������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "ProtocolFromArchives" Then

                '������ ��� ����� "����� ���������" (� ��������� "����" � ����)
        strPathFileName = gPathFileName
            '��������� ����� ������ (������) "������ ���������"
        lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '������� ��������� �������� ���� ��� ������������� �������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���������� ����� � ������ ���������� ����� "����� ���������"
        lngFileLength = LOF(intFileNum)
            '��������� ���������� ������� � ��������� ����� "����� ���������"
        lngArchivesRowNum = lngFileLength / lngRecordLen
        
            '������ ���������� ��������
        frmPrintPreview.Print Tab(3); "Name"; Tab(25); "Code or Password"; Tab(55); "Status"; _
        Tab(80); "Time"; Tab(95); "Date"; Tab(115); "Reserv. or Note"
            '������� ������ ������
        frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ������� "������ ���������"
        For lngRowNum = lngRowNum To lngArchivesRowNum Step 1
            '������ ������ "������ ���������" �� ����� � �����
            Get intFileNum, lngRowNum, gProtocol
            '������� ������ "������ ���������"
            frmPrintPreview.Print Tab(3); gProtocol.strProtocName; _
            Tab(25); gProtocol.strProtocPersonCode; Tab(55); gProtocol.strProtocStatus; _
            Tab(80); gProtocol.strProtocTime; Tab(95); gProtocol.strProtocDate; _
            Tab(115); gProtocol.strProtocReserve
            '������� ����� ������ ����� "frmPrintPreview"
            intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
            If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
                If lngRowNum < lngArchivesRowNum Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
                Exit For
            End If
        Next
    
            '��������������� ������ "������� �������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTime*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Intervals"; Tab(25); "Time"; Tab(55); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� �������"
    For lngRowNum = lngRowNum To frmTableTime.grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        frmTableTime.grdTableTime.Row = lngRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To frmTableTime.grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
            frmTableTime.grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������ "������� �������"
            strTableTime(intColNum) = frmTableTime.grdTableTime.Text
        Next
            '������� ������ "������� �������"
        frmPrintPreview.Print Tab(3); strTableTime(0); Tab(25); strTableTime(1); _
        Tab(55); strTableTime(2)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTime.grdTableTime.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
            '��������������� ������ "������� ����������" �� ����� "frmPrintPreview"
    ElseIf frmPrintPreview.Tag = "TableTerminal*" Then
            '������ ���������� ��������
    frmPrintPreview.Print Tab(3); "Terminal"; Tab(25); "Address and Port"; Tab(55); "Description"; _
    Tab(75); "Expander"
            '������� ������ ������
    frmPrintPreview.Print
            '������� ����� ������ ����� "frmPrintPreview"
    intRowPrintNum = intRowPrintNum + 2
            
            '���� �� ���� ��������������� ������� "������� ����������"
    For lngRowNum = lngRowNum To frmTableTerminal.grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        frmTableTerminal.grdTableTerminal.Row = lngRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To frmTableTerminal.grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            frmTableTerminal.grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������ "������� ����������"
            strTableTerminal(intColNum) = frmTableTerminal.grdTableTerminal.Text
        Next
            '������� ������ "������� ����������"
        frmPrintPreview.Print Tab(3); strTableTerminal(0); Tab(25); strTableTerminal(1); _
        Tab(55); strTableTerminal(2); Tab(75); strTableTerminal(3)
            '������� ����� ������ ����� "frmPrintPreview"
        intRowPrintNum = intRowPrintNum + 1
            '�������� ����� "frmPrintPreview" ���������
        If intRowPrintNum > intRowPrintQuan Then
            '������� ��������� ������ "Next"
            If lngRowNum < frmTableTerminal.grdTableTerminal.Rows - 1 Then cmdNext.Enabled = True
            '������������� ����� ����� �� ����� "frmPrintPreview"
            Exit For
        End If
    Next
    
    End If

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '�������� ����� ���������� ���������� � ��������� ���� "DateFrom"
Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
            '������ ���������� ������
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            '������ �������
        KeyAscii = 0
            '�������� ������
        frmDemo.BeepSound
    End If
                '�������� ������ ��������� ����
    If (KeyAscii = 0 Or KeyAscii = vbKeyReturn) And _
    (Mid(Trim(txtDateFrom.Text), 3, 1) <> "." Or Mid(Trim(txtDateFrom.Text), 6, 1) <> "." _
    Or Len(Trim(txtDateFrom.Text)) <> 10) Then
            '�������������� ������� ���� "From"
        txtDateFrom.Text = Trim(strDateFrom)
            '�������� ������
        frmDemo.BeepSound
    End If
            '������������ ������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
            '������ "Find" �������� ��� �������
    cmdFind.MousePointer = 0
    
End Sub

            '�������� ����� ���������� ���������� � ��������� ���� "DateTo"
Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
            '������ ���������� ������
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            '������ �������
        KeyAscii = 0
            '�������� ������
        frmDemo.BeepSound
    End If
            '�������� ������ ��������� ����
    If (KeyAscii = 0 Or KeyAscii = vbKeyReturn) And _
    (Mid(Trim(txtDateTo.Text), 3, 1) <> "." Or Mid(Trim(txtDateTo.Text), 6, 1) <> "." _
    Or Len(Trim(txtDateTo.Text)) <> 10) Then
            '�������������� ������� ���� "To"
        txtDateTo.Text = Format(Now, "dd/mm/yyyy")
            '�������� ������
        frmDemo.BeepSound
    End If
            '������������ ������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
            '������ "Find" �������� ��� �������
    cmdFind.MousePointer = 0
    
End Sub

            '�������� ����� ��������� ���������� � ��������� ���� "CodeOrPassword"
Private Sub txtCodeOrPassword_Change()
            '����� ��������� ������ ������ ����������
    If Len(Trim(txtCodeOrPassword.Text)) > 16 Then
            '������� ���������� ����
        txtCodeOrPassword.Text = ""
            '�������� ������
        frmDemo.BeepSound
    End If
            '������������ ������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
            '������ "Find" �������� ��� �������
    cmdFind.MousePointer = 0

End Sub

            '�������� ����� ��������� ���������� � ��������� ���� "Name"
Private Sub txtName_Change()
            '����� ��������� ������ ������ ����������
    If Len(Trim(txtName.Text)) > 16 Then
            '������� ���������� ����
        txtName.Text = ""
            '�������� ������
        frmDemo.BeepSound
    End If
            '������������ ������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
            '������ "Find" �������� ��� �������
    cmdFind.MousePointer = 0

End Sub

            '�������� ����� ��������� ���������� � ��������� ���� "ReservOrNote"
Private Sub txtReservOrNote_Change()
            '����� ��������� ������ ������ ����������
    If Len(Trim(txtReservOrNote.Text)) > 22 Then
            '������� ���������� ����
        txtReservOrNote.Text = ""
            '�������� ������
        frmDemo.BeepSound
    End If
            '������������ ������� ����� ������ ������� "TableProtocol"
    lngRowNum = 1
            '������ "Find" �������� ��� �������
    cmdFind.MousePointer = 0

End Sub
