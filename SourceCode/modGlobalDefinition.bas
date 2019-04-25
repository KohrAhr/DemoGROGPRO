Attribute VB_Name = "modGlobalDefinition"
Option Explicit
            '�������� ������������ ������������� ��� ������ - ������ "��������� �������"
Type SystemInfo
            '��� (������, �����������, ������, ... � �.�.)
    strObject As String * 16
            '��������� ��� ����� � ��������
    strConsAddrTerm As String * 16
            '��� (���������, �����, ������, ������������)
    strType As String * 2
            '������ - ������ �� ����� ������ "��������� �������"
    strIndex As String * 6
            '������
    strAppendix As String * 8
End Type
            '�������� ������������ ������������� ��� ������ - ������ "������� ������"
Type PersonInfo
            '������� (���)
    strName As String * 16
            '������������ ���
    strPersonCode As String * 16
            '������
    strStatus As String * 2
            '�����
    strTime As String * 8
            '���������
    strCalendar As String * 2
            '������
    strReserve As String * 8
End Type


            '�������� ����������-������ ��� ��� ����������
            '   c ActiveX.EXE
Global objTablePerson As Sel_2Server.XTablePerson
            '�������� ����������-������ ��� ��� ���������� ������
            '   c ����������� ������������ "FlexGrid"
            '   ("������� ������")
Global gTablePerson As Sel_2Server.ITablePerson

'�������
'Global objTablePerson As XTablePerson
'Global gTablePerson As ITablePerson
            
            '�������� ����������-������ MSMQQueueInfo -
            '   ��������� ��� �������� � ����������
            '   �������� ����������� ���������
Global qInfoInput As MSMQQueueInfo
            '����������-������ ������� ����������� ���������
Global qQueueInput As MSMQQueue
            '����������-������ �������-�������
            ' ����������� ���������
Global evQueue As MSMQQueue
            '����������-������ ����������� ��������E
Global qMsgInput As MSMQMessage
            '������ ����������� ����� � ������ ���������
Global strMsgInput As String
            '�������� ����������-������ MSMQQueueInfo -
            '   ��������� ��� ����������
            '   �������� ������������ ���������
Global qInfoOutput As MSMQQueueInfo
            '����������-������ ������� ������������ ���������
Global qQueueOutput As MSMQQueue
            '����������-������ ������������ ��������E
Global qMsgOutput As MSMQMessage
            
            '�������� ������������ ������������� ��� ������ - ������ "������� ���������"
Type CalendarInfo
            '����� ������
    strWeekNum As String * 20
            '����������� - ����
    strMonday As String * 4
            '������� - ����
    strTuesday As String * 4
            '����� - ����
    strWednesday As String * 4
            '������� - ����
    strThursday As String * 4
            '������� - ����
    strFriday As String * 4
            '������� - ����
    strSaturday As String * 4
            '�����������
    strSunday As String * 4
End Type

            '�������� ������������ ������������� ��� ������ - ������ "������� �������"
Type TimeInfo
            '����� ��������� �������
    strIntervalNum As String * 12
            '�����
    strTime As String * 8
            '����������
    strExpander As String * 8
End Type
            
            '�������� ������������ ������������� ��� ������ - ������ "������� ����������"
Type TerminalInfo
            '�������� (��� �� ������ "Object" "��������� �������")
    strTerminal As String * 16
            '����� � ����
    strAddrPort As String * 4
            '�������� ���������
    strDescription As String * 16
            '����������
    strExpander As String * 8
End Type

            '�������� ������������ ������������� ��� ������ - ������ "������� ���������"
Type ProtocolInfo
            '������� (���) ��� ��������
    strProtocName As String * 16
            '������������ ��� ��� ������
    strProtocPersonCode As String * 16
            '������
    strProtocStatus As String * 22
            '�����
    strProtocTime As String * 10
            '����
    strProtocDate As String * 10
            '������
    strProtocReserve As String * 22
End Type
            
            '������ "������� ���������"
Global gProtocol As ProtocolInfo
            '���� ������� ������ ����������
Global gTermContr As Integer
            '���� ���� ����������
Global gBus As Integer
            '������ ������� "Controller'��" ��� ������� ����������
            '  ����������� "�������"
Global gAddrManual(3) As String * 2
            '����� "����" ��� ���������� �����
Global gSluice(3) As Integer
            '��������� ������ � ����� ������� - ������� ����� (�.�. �� ���������)
Global gVisitor As String
            '���� ������� �������������� ���������� "������� ���������" � ����
Global gMSBase As Integer
            '���� ��������� ��� "������ ���������" (1 - ������� ������� �������,
            '  2 - ������� ���������, 4 - �������� ������� �������;
            '  �������� ����������: 0 - ��� ���������, 1+2, 1+4, 2+4, 1+2+4)
Global gDocument As Integer
            '������ - (�������� ��������)
Global gPrintSIAName As String * 16
            '����� (������) ���������/�������� ������
            '  �� ����� ������/������
Global gTalonLength As Integer
            '����� ��������������� ������������ ����� "PersonCode" � "Info"
            '  0 - ������ ������������; 1 - �������������� ��� �����������;
            '  2 - �������������� ��� �����������. �������� ����������: 1+2
Global gCreatePersonCode As Integer
            '���� � ����� ������������ ���������� "Z_������"
Global gZ_Report As String
            '������ ��������� (��������) ���������
Global gTermInp As Integer
            '������ ��������� (���������) ���������
Global gTermOut As Integer
            
            '������������ ��������� �������� ��� ��������������� ����������
Global gDefaultTime As String
            '������������ �������� ������ "Status"
            '  � "������� ������" ��� ��������������� ����������
Global gDefaultStatus As String
            '������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� ����������
Global gDefaultCalendar As String
            '������� ������������� ��������������� ������������
            '  ����� �������� ���� � "������� ���������" ���
            '  ������ ����
Global gHolidays As Integer
            '������ ������� "Controller'��" ��� �������������
            ' ����������� "�������" ����� (������) ��������
Global gEmplAddrTerm(3) As String * 2
            
            
            '����� ����������� ������������ (������������) ��� �����������
            '  �������� (����� ����������� �� ����������� - ����������� ������)
Global gTariffConst As Integer
            
            '������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� �� �����������
Global gDefaultParkCale As String
            '������ ������� "Controller'��" ����������� ��� �������������
            ' ����������� "�������" ������ (������) ��������
Global gParkAddrTerm(3) As String * 2
            '������� ����� ����������� (��� ���������� ��������)
Global gParkingD As Integer
            '������ ����� ����������� (��� ���������� ��������)
Global gParkingN As Integer
            '�������� ����� ����������� (��� ���������� ��������)
Global gParkingDN As Integer
            '������� ������� ����� ����������� (��� ��������� ��������)
Global gParkingHourD As Integer
            '������� ������ ����� ����������� (��� ��������� ��������)
Global gParkingHourN As Integer
            '������� ����� ������� � ����������� (��� ����������� ��������)
Global gParkingTimeD As String
            '������������ ����� ������� � ����������� (��� ���� ��������)
Global gDefaultParkTime As String
            '����� ������������ (��������) ����� ������� ��������
            '  ��� �����������/���������� ���������� ������� �����������
Global gParkingTimeCell As Integer
            '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
            '  (���������� ����������� ������ ��� ��� ������ ���������� �����
            '   ���������� gParkTimeLimit � gParkingTimeCell)
Global gParkingCellLimit As Integer
            '���������� ������������ �������������� ����� �������
            '  ��� ����������� ��������� ��������
Global gParkInpCellNumb As Integer
            '�������� ����� ����������� (��� ��������� ��������)
Global gParkingMoneyCell As Integer
            '����� "����������/���������� ��������" ��� ����������
            '  ������� ����������� �� "������� ������"
Global gParkingDeletion As Integer
            '����� ����������� "PersonCode"� ���� "Info"
            '  ��� ����������� ���������� ������� �����������
Global gParkingCodeInfo As Integer
            '����� ������� ������������� �������� ���������
            '  ��� �����������/���������� ���������� ������� �����������
Global gParkingPresButton As Integer
            '����� (� ���.) ������� ������������ ����������
            '  ����������� ������� �� �����������
Global gParkTimeLimit As Integer
            '���������� ���� �� �����������
Global gParkingPlaceNum As Integer
            '���������� ��������� ���� �� �����������
Global gParkFreePlaces As Integer

            '������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� �����������
Global gDefaultAcceCale As String
            '������ ������� "Controller'��" ����������� ��� �������������
            ' ����������� "�������" ����� (������) �����������
Global gAcceAddrTerm(3) As String * 2
            '������� ����� ����������� (��� ���������� ��������)
Global gAccessD As Integer
            '������ ����� ����������� (��� ���������� ��������)
Global gAccessN As Integer
            '�������� ����� ����������� (��� ���������� ��������)
Global gAccessDN As Integer
            '������� ������� ����� ����������� (��� ��������� ��������)
Global gAccessHourD As Integer
            '������� ������ ����� ����������� (��� ��������� ��������)
Global gAccessHourN As Integer
            '������� ����� ������� � ����������� (��� ����������� ��������)
Global gAccessTimeD As String
            '������������ ����� ������� �� ����������� (��� ���� ��������)
Global gDefaultAcceTime As String
            '����� ������������ (��������) ����� ������� ���������
            '  ��� �����������/���������� ���������� �������
Global gAccessTimeCell As Integer
            '  ���������� ����� �������, � ������� �������� �����������
            '   ����������� ���������� ���������� ���������� �� �����������
            '  (���������� ����������� ������ ��� ��� ������ ���������� �����
            '   ���������� gAcceTimeLimit � gAccessTimeCell)
Global gAccessCellLimit As Integer
            '���������� ������������ �������������� ����� �������
            '  ��� ����������� ��������� ��������
Global gAcceInpCellNumb As Integer
            '������� ����� ����������� (��� ��������� ��������)
Global gAccessMoneyCell As Integer
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
Global gAccessMoneyCellHuman As Integer
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
Global gAccessMoneyCellBaby As Integer
            '������� ����� ����������� ��� ������ (��� ��������� ��������)
Global gAccessMoneyCellConvoy As Integer
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
Global gAccessMoneyCellFamily As Integer
            '����� "����������/���������� ��������" ��� ����������
            '  ������� ����������� �� "������� ������"
Global gAccessDeletion As Integer
            '����� ����������� "PersonCode"� ���� "Info"
            '  ��� ����������� ���������� ������� �����������
Global gAccessCodeInfo As Integer
            '����� ������� ������������� �������� ���������
            '  ��� �����������/���������� ���������� ������� �����������
Global gAccessPresButton As Integer
            '��������� ������� ��������� # 1, 2, 3 � 4
Global gLease1 As Integer
Global gLease2 As Integer
Global gLease3 As Integer
Global gLease4 As Integer
            '����� (� ���.) ������� ������������ ����������
            '  ����������� ������� �� �����������
Global gAcceTimeLimit As Integer
            '���������� ���� �� �����������
Global gAccessPlaceNum As Integer
            '���������� ��������� ���� �� �����������
Global gAcceFreePlaces As Integer
            
            '������ ��� ����� "����� ���������" (� ��������� "����" � ����)
Global gPathFileName As String
            '������ �� "������� ���������" - (������� ����)
Global gToday() As String * 4
            '����� ������ "������� ���������", ��� ����������� ������ �������� ���
Global gRowNum As Integer
            '����� ������� "������� ���������", ��� ����������� ������ �������� ���
Global gColNum As Integer
            '������ ���������� ������� ��� ���� ��������� "������� �������"
Global gInterval() As String * 11
            '������ �������������� ���������� � ���������� ���
            '  ���� ��������� "������� �������"
Global gTerCal() As String * 12
            '������ ���������� ������� ��� ���� ��������� "������� ����������"
Global gAddrPort() As String * 4
            '������� �����  ��������� ������ "������� ���������"
Global gProtocRowNum As Integer
            '������� ����� ����� "������� ���������"
Global gProtocFileNum As Integer
            '������� �����  ��������� ������ DUMMY �����
Global gDummyRowNum As Long
            '������� ����� DUMMY �����
Global gFileDummy As Integer
            '��� ������������ ������ (��� ��������� StartUp)
Global gModuleStartUp As String
            '������� ����� ����������� � ������������ �����
Global gPreprocFileNum As Integer
            '��� ������������� � ��������� ����
Global gPreprocName As String
            '��� "Host'a" � ��������� ����
Global gHost As String
            '���������� ����������� � ��������� ���� (��������
            '  ����������� ������������)
Global gNetPreprocNum As Integer
            '������ ���� ����������� ��������� ���� (��������
            '  ����������� ������������)
Global gSocketNet() As String
            '������ ������������� (����� ������ "��������� �������"
            '   ������������� � ����������� � ���)
Global gPreprocIndex As Integer
            '����� ������������� ���� ��� �������� ��������� J-�� �����
Global gPersonCode(3) As Integer
            '���������� ��������/���������� ����� � "��������� �������"
Global gAddDelRowTableSystem
            '������� �������� ��������� � "��������� �������"
Global gChangesTableSystem As Boolean
            '������� �������� ��������� � "������� ������"
Global gChangesTablePerson As Boolean
            '���� ��������� (�����������) �������� ����� �� "������� ������"
Global gRealDelPerson As Boolean
            '������� �������� ��������� � "������� ���������"
Global gChangesTableCalendar As Boolean
            '���������� ��������� "������� ���������"
Global gVarNumCalendar As Integer
            '���������� ��������/���������� ����� � "������� �������"
Global gAddDelRowTableTime
            '������� �������� ��������� � "������� �������"
Global gChangesTableTime As Boolean
            '���������� ��������� "������� �������"
Global gVarNumTime As Integer
            '���������� ��������/���������� ����� � "������� ����������"
Global gAddDelRowTableTerminal
            '������� �������� ��������� � "������� ����������"
Global gChangesTableTerminal As Boolean
            '���������� ��������� "������� ����������"
Global gVarNumTerminal As Integer
            '�������  �������� ��������� � "������� ����������"
Global gChangesTableTemperature As Boolean
            '�������  �������� ��������� � "������� ���������"
Global gChangesTablePower As Boolean
             '���������� ����� �� ����� �������� ����� "frmPrintPreview"
Global gRowPrintQuan As Integer
            '���������� ���� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � "Host Computer" � ���
            '  ������������ �� ������� ��� ������
Global gDayNum As Integer
            '������� ������������� ������ "������� ������":
            '   ��������������� ������ � "Host Computer'e" � � ��� �������
            '   � "�������������", ����� ��������� ���������� ����
            '   ����������� "������� ������" - "���������� ������a ������"
Global gCompresTablPers As Integer
            '������� ���
Global gYear As Integer
            '����� ���������� ������� (������������ ���������� ��������)
Global gTimeShare As Integer
            '������������ ��������� �������
Global gBeepSound As Integer
            '����������� ��������� ����� ��� ��������������
            '  ������������ ���� "PersonCode"
Global gMinCount As Integer
            '������������ ��������� ����� ��� ��������������
            '  ������������ ���� "PersonCode"
Global gMaxCount As Integer
            '���������� ��������� ���������� ��������� ����
            '  �� ������� �� ��������� �� ��������� ��������� ����
            '  (��� ���������� ���������� ��� ����������
            '  ������/�������)
Global gDisplayDiscount As Integer

            '���������� API �������, ����������� ��� ������
            '  � �������� "Mutex"
Declare Function WaitForSingleObject Lib "kernel32" ( _
   ByVal hHandle As Long, _
   ByVal dwMilliseconds As Long) As Long

Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" ( _
   ByVal lpMutexAttributes As Long, _
   ByVal bInitialOwner As Long, _
   ByVal lpName As String) As Long

Declare Function ReleaseMutex Lib "kernel32" ( _
   ByVal hMutex As Long) As Long
   
            '���������� ������� "Mutex", ������������ ���
            '  ������������� �������, ��������� ������������
            '  ����� ������� - ������� ����������� "������� ������"
            '  � ���� "������� ������" (��� ������������ �� �����)
Global gMutex As Long


