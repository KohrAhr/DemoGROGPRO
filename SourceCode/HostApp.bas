Attribute VB_Name = "HostApp"
Option Explicit
            '����� "����������� ����������" ���������� - ��� �������������� � ��.
Dim intTerminalLogOFF As Integer
            '������� "����� ����" (������������ ��� �������� �
            '   ��������� ������ 'Controllera"
Dim intWhite As Integer
            '����� "�����������" ��� ���������� �����
Dim intParking(3) As Integer
            '����� "����������" ��� ���������� �����
Dim intAccess(3) As Integer
            '����� "��������" ��� ���������� �����
Dim intEmploye(3) As Integer
            '������ ����� �� ����� �������� ���� ����������
Dim intIndex As Integer
            '����� "Controller'a" � "Port'a" ��� ���������������/��������
Dim intAutoRegDel(3) As Integer
            '�����  "������ ����������"
Dim strParole As String
            '����� "������ ���������" ��� �����������/�������� �������
Dim intDocument As Integer
            '����� - ����
Dim intHour As Integer
Dim strHour As String
            '����� - ������
Dim intMinute As Integer
Dim strMinute As String
            '����� ������� ������ � "������� ������"
Dim intRowNum As Integer
            '������ �������� � ������� ��������� ���������� ����
Dim intControlIndex As Integer
            '����� ������ ������ �� "Controller'��" ����������
Dim vntBufferInput(3) As Variant
            '������ ����������� ���������
Dim strMessage As String

            ' ������� ��������� - ��������� ������� �����/������ ��� ����������
Public Sub Main()
            '����� ������� ������ � "��������� �������"
Dim intRowNumSys As Integer
            '����� ������� �������� "/" � ������������� ����
Dim intPosNum As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '������� �������
Dim intCount As Integer
            '�����e� ���e
Dim intWork As Integer

            ' ��������� �� ��������� ����� "frmTableSystem"
    Load frmTableSystem
            ' ��������� �� ��������� ����� "frmDemo"
    Load frmDemo
            '������� ��������� ���� "Parking", "Access" � "Employe" ����� "frmDemo"
    frmDemo.mnuParking.Visible = False
    frmDemo.mnuAccess.Visible = False
    frmDemo.mnuEmploye.Visible = False
            '������� ����������� ���� "Parking", "Access" � "Employe" ����� "frmDemo"
    frmDemo.mnuParking.Enabled = False
    frmDemo.mnuAccess.Enabled = False
    frmDemo.mnuEmploye.Enabled = False
    
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNumSys = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������  "��������� �������"
        frmTableSystem.grdTableSystem.Row = intRowNumSys
            '������������� ������� "��������� �������" (������)
        frmTableSystem.grdTableSystem.Col = 0
            '������������� �������� �������
        If Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� �����
            frmDemo.prtPortC(0).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� �����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(0).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" �����
            frmDemo.prtPortC(0).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� �����
            frmDemo.prtPortC(1).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� �����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(1).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" �����
            frmDemo.prtPortC(1).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� �����
            frmDemo.prtPortC(2).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� �����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(2).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" �����
            frmDemo.prtPortC(2).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtPortC(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� �����
            frmDemo.prtPortC(3).CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� �����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortC(3).Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" �����
            frmDemo.prtPortC(3).InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ModuleStartUp" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ��� "StartUp" ������ �������
            gModuleStartUp = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "txtPassword" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������ �������
            frmDemo.txtPassword.Tag = Trim(frmTableSystem.grdTableSystem.Text)
            '������� ������� "��������� �������"=4(������)
            frmTableSystem.grdTableSystem.Col = 4
            '���������� ��������� ������������� � ��������� ���� ���
            '   "�����" ��� "Host Computer'a"
            If Trim(frmTableSystem.grdTableSystem.Text) <> "" Then
                gPreprocName = Trim(frmTableSystem.grdTableSystem.Text)
            Else
                gPreprocName = ""
            '������������� ������� ������� ������� ���� �����������
            '  ��������� ����
                ReDim gSocketNet(0) As String
                gSocketNet(0) = ""
            End If
            '��� �� "Host Computer"
            If gPreprocName <> "" Then
            '������� ������� "��������� �������"=3(������)
                frmTableSystem.grdTableSystem.Col = 3
            '���������� ������ ������������� (����� ������ "��������� �������"
            '   ������������� � ����������� � ���)
                gPreprocIndex = Trim(frmTableSystem.grdTableSystem.Text)
            '������������� ������� ������� ������� ���� �����������
            '  ��������� ����
                ReDim gSocketNet(0) As String
                gSocketNet(0) = ""
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Host" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ��� "Host Computer'a" � ��������� ����
            gHost = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "txtParole" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������ ����������
            strParole = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Visitor" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� �����
            gVisitor = Left(Trim(frmTableSystem.grdTableSystem.Text), 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Z_Report" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���� � ����� ������������ ���������� "Z_������"
            gZ_Report = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TariffConst" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '����� ����������� ������������ (������������) ��� �����������
            '  �������� (����� ����������� �� ����������� - ����������� ������)
            gTariffConst = Trim(frmTableSystem.grdTableSystem.Text)
            If gTariffConst > 32000 Then gTariffConst = 32000
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingPresButt" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������� ������������� �������� ���������
            '   ��� �����������/���������� ��������� �������� �����������
            gParkingPresButton = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingDeletion" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������/���������� ��������"
            '   ��� ���������� ������� ����������� �� "������� ������"
            gParkingDeletion = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingCode_Info" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ����������� "PersonCode"� ���� "Info"
            '   ��� ����������� ���������� ������� �����������
            gParkingCodeInfo = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingDN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ����������� (��� ���������� ��������)
            gParkingDN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ����� ����������� (��� ���������� ��������)
            gParkingD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������ ����� ����������� (��� ���������� ��������)
            gParkingN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingHourD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������� ����� ����������� (��� ��������� ��������)
            gParkingHourD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingHourN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������ ����� ����������� (��� ��������� ��������)
            gParkingHourN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingMoneyCell" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ����������� (��� ��������� ��������)
            gParkingMoneyCell = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeCell" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ (��������) ����� ������� ��������
            '  ��� �����������/���������� ���������� ������� �����������
            gParkingTimeCell = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkingTimeCell = 0 Then gParkingTimeCell = 15
            If gParkingTimeCell > 1440 Then gParkingTimeCell = 1440
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeLimit" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� (� ���.) ������� ������������ ����������
            '  ����������� ������� �� �����������
            gParkTimeLimit = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkTimeLimit > 1440 Then gParkTimeLimit = 1440
             '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
            gParkingCellLimit = Int(gParkTimeLimit / gParkingTimeCell)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkInpCellNumb" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������������� ����� �������
            '  ��� ����������� ��������� ��������
            gParkInpCellNumb = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkInpCellNumb * gParkingTimeCell > 1440 Then _
            gParkInpCellNumb = 1440 / gParkingTimeCell
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingPlaceNum" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���� �� �����������
            gParkingPlaceNum = Trim(frmTableSystem.grdTableSystem.Text)
            If gParkingPlaceNum > 999 Then gParkingPlaceNum = 999
            '���������� ��������� ���� �� �����������
            gParkFreePlaces = gParkingPlaceNum
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessPresButt" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������� ������������� �������� ���������
            '   ��� �����������/���������� ��������� ����������� �����������
            gAccessPresButton = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessDeletion" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������/���������� ��������"
            '   ��� ���������� ���������� ����������� �� "������� ������"
            gAccessDeletion = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessCode_Info" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ����������� "PersonCode"� ���� "Info"
            '   ��� ����������� ���������� ���������� �����������
            gAccessCodeInfo = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessDN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ����������� (��� ���������� �����������)
            gAccessDN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ����� ����������� (��� ���������� �����������)
            gAccessD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������ ����� ����������� (��� ���������� �����������)
            gAccessN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessHourD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������� ����� ����������� (��� ��������� �����������)
            gAccessHourD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessHourN" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������ ����� ����������� (��� ��������� �����������)
            gAccessHourN = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessMoneyCell" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ����� ��� �������� (��� ��������� �����������)
            intPosNum = InStr(2, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellHuman = _
            Left(frmTableSystem.grdTableSystem.Text, intPosNum - 1)
            gAccessMoneyCell = gAccessMoneyCellHuman
            '���������� ������� ����� ��� ����� (��� ��������� �����������)
            gAccessMoneyCellBaby = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            '���������� ������� ����� ��� ������ (��� ��������� �����������)
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellConvoy = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            '���������� ������� ����� ��� ����� (��� ��������� �����������)
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gAccessMoneyCellFamily = _
            Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeCell" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ (��������) ����� ������� ���������
            '  ��� �����������/���������� ���������� ���������� �����������
            gAccessTimeCell = Trim(frmTableSystem.grdTableSystem.Text)
            If gAccessTimeCell = 0 Then gAccessTimeCell = 15
            If gAccessTimeCell > 1440 Then gAccessTimeCell = 1440
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeLimit" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� (� ���.) ������� ������������ ����������
            '  ����������� ������� �� �����������
            gAcceTimeLimit = Trim(frmTableSystem.grdTableSystem.Text)
            If gAcceTimeLimit > 1440 Then gAcceTimeLimit = 1440
             '���������� ����� �������, � ������� �������� �����������
            '  ����������� ������� ���������� ���������� �� �����������
            '  � ���������� �������������� ����� ������� ��� ���������� �������
            gAccessCellLimit = Int(gAcceTimeLimit / gAccessTimeCell)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceInpCellNumb" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������������� ����� �������
            '  ��� ����������� ��������� ��������
            gAcceInpCellNumb = Trim(frmTableSystem.grdTableSystem.Text)
            If gAcceInpCellNumb * gAccessTimeCell > 1440 Then _
            gAcceInpCellNumb = 1440 / gAccessTimeCell
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessPlaceNum" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���� �� �����������
            gAccessPlaceNum = Trim(frmTableSystem.grdTableSystem.Text)
            If gAccessPlaceNum > 999 Then gAccessPlaceNum = 999
            '���������� ��������� ���� �� �����������
            gAcceFreePlaces = gAccessPlaceNum
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Document" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '����� "������ ���������" ��� �����������/�������� �������
            gDocument = Trim(frmTableSystem.grdTableSystem.Text)
            If gDocument <> 0 Then intDocument = 1
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtDocument" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ����� ��� �������� �������� ��������
            frmDemo.prtPortDocument.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� ����� ��� �������� �������� ��������
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortDocument.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" ����� ��� �������� �������� ��������
            frmDemo.prtPortDocument.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtBarCode" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ����� ��� �������� �����-����
            frmDemo.prtPortBarCode.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� ����� ��� �������� �����-����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortBarCode.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" ����� ��� �������� �����-����
            frmDemo.prtPortBarCode.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "prtDisplay" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ����� ��� ������� ���������� ��������� ����
            frmDemo.prtPortDisplay.CommPort = Left(frmTableSystem.grdTableSystem.Text, 1)
            '���������� ��������� ����� ��� ������� ���������� ��������� ����
            intPosNum = InStr(3, frmTableSystem.grdTableSystem, "/")
            frmDemo.prtPortDisplay.Settings = Mid(frmTableSystem.grdTableSystem.Text, 3, intPosNum - 3)
            '���������� �������� "InputMode" ����� ��� ������� ����������
            '  ��������� ����
            frmDemo.prtPortDisplay.InputMode = Mid(frmTableSystem.grdTableSystem.Text, _
            intPosNum + 1, 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Lease" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ��������� ������� ��������� # 1
            intPosNum = InStr(2, frmTableSystem.grdTableSystem, "/")
            gLease1 = Left(frmTableSystem.grdTableSystem.Text, intPosNum - 1)
            '���������� ��������� ������� ��������� # 2
            gLease2 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            '���������� ��������� ������� ��������� # 3
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gLease3 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1, _
            InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/") - intPosNum - 1)
            '���������� ��������� ������� ��������� # 4
            intPosNum = InStr(intPosNum + 1, frmTableSystem.grdTableSystem, "/")
            gLease4 = Mid(frmTableSystem.grdTableSystem.Text, intPosNum + 1)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "Controller'��" N_0
            frmDemo.tmrTimeOut(0).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "Controller'��" N_1
            frmDemo.tmrTimeOut(1).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "Controller'��" N_2
            frmDemo.tmrTimeOut(2).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTimeOut(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "Controller'��" N_3
            frmDemo.tmrTimeOut(3).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "������ �������������" N_0
            frmDemo.tmrButton(0).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "������ �������������" N_1
            frmDemo.tmrButton(1).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "������ �������������" N_2
            frmDemo.tmrButton(2).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrButton(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� "������ �������������" N_3
            frmDemo.tmrButton(3).Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrTermContr" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ������ "Controller'��" ����������
            frmDemo.tmrTermContr.Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "tmrPasswTimeOut" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "TimeOut' ��� ����� ������
            frmDemo.tmrPasswTimeOut.Interval = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���������� N_0
            frmDemo.chkTerm(0).Value = frmTableSystem.grdTableSystem.Text
            '������� "����������������" ������� ��������� (��������) ���������
            gTermInp = -1
            '������� "����������������" ������� ��������� (���������) ���������
            gTermOut = -1
            '���� ��������� N_0 ������������� ��������
            If frmDemo.chkTerm(0).Value = 1 Then
            '������� ������� "��������� �������"=4(������)
                frmTableSystem.grdTableSystem.Col = 4
            '���� ��� �������� (�������) �������� �������
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            '���������� ������ ��������� (��������) ���������
                    gTermInp = 0
            '���� ��� �������� (��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            '���������� ������ ��������� (���������) ���������
                    gTermOut = 0
            '���� ��� ��������/�������� (�������/��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            '���������� ������ ��������� (��������) ���������
                    gTermInp = 0
            '���������� ������ ��������� (���������) ���������
                    gTermOut = 0
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���������� N_1
            frmDemo.chkTerm(1).Value = frmTableSystem.grdTableSystem.Text
            '���� ��������� N_1 ������������� ��������
            If frmDemo.chkTerm(1).Value = 1 Then
            '������� ������� "��������� �������"=4(������)
                frmTableSystem.grdTableSystem.Col = 4
            '���� ��� �������� (�������) �������� �������
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 1
            '���� ��� �������� (��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 1
            '���� ��� ��������/�������� (�������/��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 1
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 1
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���������� N_2
            frmDemo.chkTerm(2).Value = frmTableSystem.grdTableSystem.Text
            '���� ��������� N_2 ������������� ��������
            If frmDemo.chkTerm(2).Value = 1 Then
            '������� ������� "��������� �������"=4(������)
                frmTableSystem.grdTableSystem.Col = 4
            '���� ��� �������� (�������) �������� �������
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 2
            '���� ��� �������� (��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 2
            '���� ��� ��������/�������� (�������/��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 2
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 2
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkTerm(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���������� N_3
            frmDemo.chkTerm(3).Value = frmTableSystem.grdTableSystem.Text
            '���� ��������� N_3 ������������� ��������
            If frmDemo.chkTerm(3).Value = 1 Then
            '������� ������� "��������� �������"=4(������)
                frmTableSystem.grdTableSystem.Col = 4
            '���� ��� �������� (�������) �������� �������
                If Trim(frmTableSystem.grdTableSystem.Text) = "Inp" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 3
            '���� ��� �������� (��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Out" Then
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 3
            '���� ��� ��������/�������� (�������/��������) �������� �������
                ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "InpOut" Then
            '���������� "���������������" ������ ��������� (��������) ���������
                    If gTermInp = -1 Then gTermInp = 3
            '���������� "���������������" ������ ��������� (���������) ���������
                    If gTermOut = -1 Then gTermOut = 3
                End If
            End If
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���� ��� ���������� N_0
            frmDemo.chkPhoto(0).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���� ��� ���������� N_1
            frmDemo.chkPhoto(1).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���� ��� ���������� N_2
            frmDemo.chkPhoto(2).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "chkPhoto(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ������������ ���� ��� ���������� N_3
            frmDemo.chkPhoto(3).Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optEnglish" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ��� ����������� �����
            frmDemo.optEnglish.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optLatvian" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ��� ���������� �����
            frmDemo.optLatvian.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optRussian" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ��� �������� �����
            frmDemo.optRussian.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optAutomatic" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ��� ��������������� ���������� �����������
            frmDemo.optAutomatic.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "optManual" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �������� ����� ��� ������� ���������� �����������
            frmDemo.optManual.Value = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������������� ���� ��� "Controller'��" N_0
            gPersonCode(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������������� ���� ��� "Controller'��" N_1
            gPersonCode(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������������� ���� ��� "Controller'��" N_2
            gPersonCode(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gPersonCode(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ������������� ���� ��� "Controller'��" N_3
            gPersonCode(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gRowPrintQuan" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���������� ����� �� ����� �������� ����� "frmPrintPreview"
            gRowPrintQuan = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gDayNum" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���������� ���� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � "Host Computer" � ���
            '  ������������ �� ������� ��� ������
            gDayNum = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gCompresTablPers" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������������� ������ "������� ������"
            gCompresTablPers = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gYear" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ���
            gYear = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumTime" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���������� ��������� "������� �������"
            gVarNumTime = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumCalendar" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���������� ��������� "������� ���������"
            gVarNumCalendar = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "gVarNumTerminal" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ���������� ��������� "������� ����������"
            gVarNumTerminal = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "Controller'a" ��� ������� ����������
            '  ����������� "�������" "N_0"
            gAddrManual(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "Controller'a" ��� ������� ����������
            '  ����������� "�������" "N_1"
            gAddrManual(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "Controller'a" ��� ������� ����������
            '  ����������� "�������" "N_2"
            gAddrManual(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AddrManual(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "Controller'a" ��� ������� ����������
            '  ����������� "�������" "N_3"
            gAddrManual(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� a���� "Controller'a" � "Port'a" ���
            '  ���������������/��������
            intAutoRegDel(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� a���� "Controller'a" � "Port'a" ���
            '  ���������������/��������
            intAutoRegDel(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� a���� "Controller'a" � "Port'a" ���
            '  ���������������/��������
            intAutoRegDel(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AutoRegDel(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� a���� "Controller'a" � "Port'a" ���
            '  ���������������/��������
            intAutoRegDel(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����" ��� ���������� "N_0" -
            '  ������ ����
            gSluice(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����" ��� ���������� "N_1" -
            '  ������ ����
            gSluice(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����" ��� ���������� "N_2" -
            '  ������ ����
            gSluice(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Sluice(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����" ��� ���������� "N_3" -
            '  ������ ����
            gSluice(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TerminalLogOFF" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "���������� ����������" ���������� -
            '  ��� �������������� � ��.
            intTerminalLogOFF = frmTableSystem.grdTableSystem.Text
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�����������" ��� ���������� "N_0"
            intParking(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�����������" ��� ���������� "N_1"
            intParking(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�����������" ��� ���������� "N_2"
            intParking(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Parking(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�����������" ��� ���������� "N_3"
            intParking(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�������������������������" ��� ���������� "N_0"
            gParkAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�������������������������" ��� ���������� "N_1"
            gParkAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�������������������������" ��� ���������� "N_2"
            gParkAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkAddrTerm(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "�������������������������" ��� ���������� "N_3"
            gParkAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "ParkingTimeD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ��������� �������� �����������
            '   ��� ��������������� (��� ���������� ��������)
            gParkingTimeD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultParkTime" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ ��������� �������� ������� �
            '   ����������� ��� ��������������� ��������
            gDefaultParkTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultParkCale" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� �������� �� �����������
            gDefaultParkCale = Trim(frmTableSystem.grdTableSystem.Text)
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������" ��� ���������� "N_0"
            intAccess(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������" ��� ���������� "N_1"
            intAccess(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������" ��� ���������� "N_2"
            intAccess(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Access(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "����������" ��� ���������� "N_3"
            intAccess(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_0"
            gAcceAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_1"
            gAcceAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_2"
            gAcceAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AcceAddrTerm(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_3"
            gAcceAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "AccessTimeD" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ��������� �������� ���������
            '   ��� ��������������� (��� ���������� ��������)
            gAccessTimeD = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultAcceTime" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ ��������� �������� �������
            '   ��� ��������������� �����������
            gDefaultAcceTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultAcceCale" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� �����������
            gDefaultAcceCale = Trim(frmTableSystem.grdTableSystem.Text)
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "��������" ��� ���������� "N_0"
            intEmploye(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "��������" ��� ���������� "N_1"
            intEmploye(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "��������" ��� ���������� "N_2"
            intEmploye(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "Employe(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "��������" ��� ���������� "N_3"
            intEmploye(3) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(0)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_0"
            gEmplAddrTerm(0) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(1)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_1"
            gEmplAddrTerm(1) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(2)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_2"
            gEmplAddrTerm(2) = frmTableSystem.grdTableSystem.Text
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "EmplAddrTerm(3)" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� "������������������������" ��� ���������� "N_3"
            gEmplAddrTerm(3) = frmTableSystem.grdTableSystem.Text
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultTime" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ ��������� �������� ��� ��������������� ��������
            gDefaultTime = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultStatus" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������� ������ "Status"
            '  � "������� ������" ��� ��������������� ��������
            gDefaultStatus = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DefaultCalendar" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ �������� ������ "Calendar"
            '  � "������� ������" ��� ��������������� ��������
            gDefaultCalendar = Trim(frmTableSystem.grdTableSystem.Text)
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "NewCalend<='/*'" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ������������� ��������������� ������������
            '  ����� �������� ���� � "������� ���������" ��� ������ ����
            gHolidays = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "PrintSIAName" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ��� ��������-��������� ���
            '   ��� ������ �� ������� �����
            gPrintSIAName = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TalonLength" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ���������/�������� ������
            '   �� ����� ������/������
            gTalonLength = Trim(frmTableSystem.grdTableSystem.Text)
            If gTalonLength < 0 Then gTalonLength = 0
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "CreatePersonCode" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� �����  ��������������� ������������ �����
            '  "PersonCode" � "Info"
            gCreatePersonCode = Trim(frmTableSystem.grdTableSystem.Text)
      
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "MinCount" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����������� ��������� ����� ��� ��������������
            '  ������������ ���� "PersonCode"
            gMinCount = Trim(frmTableSystem.grdTableSystem.Text)
            If gMinCount < 10 Then gMinCount = 10
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "MaxCount" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ ��������� ����� ��� ��������������
            '  ������������ ���� "PersonCode"
            gMaxCount = Trim(frmTableSystem.grdTableSystem.Text)
            If gMaxCount > 99 Then gMaxCount = 99
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DisplayDiscount" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ��������� ���������� ��������� ����
            '  �� ������� �� ��������� �� ��������� ��������� ����
            '  (��� ���������� ���������� ��� ����������
            '  ������/�������)
            gDisplayDiscount = Trim(frmTableSystem.grdTableSystem.Text)
            If gDisplayDiscount < 0 Then gDisplayDiscount = 0
            
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "TimeShare" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� ���������� ������� (������������ ���. ��������)
            gTimeShare = Trim(frmTableSystem.grdTableSystem.Text)

        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "BeepSound" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ������������ ��������� �������
            gBeepSound = Trim(frmTableSystem.grdTableSystem.Text)
        
        ElseIf Trim(frmTableSystem.grdTableSystem.Text) = "DownLoadMSBase" Then
            '������� ������� "��������� �������"=1(���������)
            frmTableSystem.grdTableSystem.Col = 1
            '���������� ����� �������������� ���������� "������� ���������" � ����
            gMSBase = Trim(frmTableSystem.grdTableSystem.Text)
        End If
            '������� ������� "��������� �������" = 2 (���)
        frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor (������������)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '���������� ����������� � ��������� ���� (�������� �����������)
            gNetPreprocNum = gNetPreprocNum + 1
            '�������������� ����������� ������� ���� �����������
            '  ��������� ���� (�������� ����������� ������������)
            ReDim Preserve gSocketNet(gNetPreprocNum) As String
            If gPreprocIndex = intRowNumSys Then
            '��� ���������� ��������� ���� (�������� ����������� ������������)
                gSocketNet(gNetPreprocNum) = gHost
            '������� ������� "��������� �������"=1(���������)
                frmTableSystem.grdTableSystem.Col = 1
            '�������������� ��� ���������� ��������� ���� - ������ ����������
                gPreprocName = Trim(frmTableSystem.grdTableSystem.Text)
            Else
            '������� ������� "��������� �������"=1(���������)
                frmTableSystem.grdTableSystem.Col = 1
            '��� ���������� ��������� ���� (�������� ����������� ������������)
                gSocketNet(gNetPreprocNum) = Trim(frmTableSystem.grdTableSystem.Text)
            End If
        End If
    Next

            ' ���� ��� ������������
    If gPreprocName <> "" Then
            '����������� "������� ������" �� "Host Computer'a"
        Call TablePersonCopy
    End If
            ' ��������� �� ��������� ����� "frmTablePerson"
    Load frmTablePerson
            ' ���� ��� "Host Computer"
    If gPreprocName = "" Then
            ' �������� ������ �� ������������ � "Host Computer'e"
            '   ������ ActiveX.EXE
        Set objTablePerson = New XTablePerson
            ' �������� ������ �� ����������, ����������� ���
            '   ������� "FlexGrid" ("������� ������")
        Set gTablePerson = objTablePerson
            
            ' ���� ������� ������������� � ��������� ����
        If gNetPreprocNum > 0 Then
            ' �������� ������� MSMQQueueInfo ��� ����������
            '  �������� ������������ ���������
            Set qInfoOutput = New MSMQQueueInfo
            ' ������� ��������� ������� ���������� ���������
            Set qMsgOutput = New MSMQMessage
            ' ������������ ����������� ���������
            qMsgOutput.Label = gHost
            qMsgOutput.Body = "Host Started"
        
            '�� ���� ��������� ������� ���� ����������� ��������� ����
            For intCount = 1 To gNetPreprocNum
            ' ���������� ���� � ������� ������������ ���������
                qInfoOutput.FormatName = "DIRECT=OS:" + _
                gSocketNet(intCount) + "\Private$\GeneralQueue"
            ' ������� ������� ��������� � ����������� (��� ��������
            '   ���������, ������ � ������� �������� ����)
                Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' �������� ���������
                qMsgOutput.Send qQueueOutput
            ' ������� ������� ���������
                qQueueOutput.Close
            Next
            
        End If
        
            ' ���� ��� ������������
    Else
            '������� ������������� ������ "������� ������" �� ����������:
            '   "������������" ���������� "������� ������" "Host Computer'�"
        If gCompresTablPers = 0 Then
            ' ��������� ��������� ���������� � "Host Computer'e",
            '   ���� ��� ��� �� �����������, � �������� ������ ��
            '   ������ ActiveX.EXE � ���
            Set objTablePerson = CreateObject("Sel_2Server.XTablePerson")
            ' �������� ������ �� ����������, ����������� ���
            '   ������� "FlexGrid" ("������� ������")
            Set gTablePerson = objTablePerson
        
            '������� ������������� ������ "������� ������" ����������:
            '   "������������" ���������� ����������� "������� ������"
        Else
            ' �������� ������ �� ������������ � "Host Computer'e"
            '   ������ ActiveX.EXE
            Set objTablePerson = New XTablePerson
            ' �������� ������ �� ����������, ����������� ���
            '   ������� "FlexGrid" ("������� ������")
            Set gTablePerson = objTablePerson
        End If
'�������
'Set objTablePerson = New XTablePerson
'Set gTablePerson = objTablePerson
        
            ' �������� ������� MSMQQueueInfo ��� ����������
            '  �������� ������������ ���������
        Set qInfoOutput = New MSMQQueueInfo
            ' ������� ��������� ������� ���������� ���������
        Set qMsgOutput = New MSMQMessage
            ' ������������ ����������� ���������
        qMsgOutput.Label = gPreprocName
            ' ���������� ���� � ������� ������������ ���������
        qInfoOutput.FormatName = "DIRECT=OS:" + gHost + "\Private$\GeneralQueue"
            ' ������� ������� ��������� � ����������� (��� ��������
            '   ���������, ������ � ������� �������� ����)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' ������������ ����������� ���������
        qMsgOutput.Body = "Time"
            ' �������� ���������
        qMsgOutput.Send qQueueOutput
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
        If gParkingPlaceNum <> 0 Then
            ' ������������ ����������� ���������
            qMsgOutput.Body = "ParkFreePlaces "
            ' �������� ���������
            qMsgOutput.Send qQueueOutput
        End If
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
        If gAccessPlaceNum <> 0 Then
            ' ������������ ����������� ���������
            qMsgOutput.Body = "AcceFreePlaces "
            ' �������� ���������
            qMsgOutput.Send qQueueOutput
        End If
            ' ������� ������� ���������
        qQueueOutput.Close
        
            ' ������������ ����������� ���������
        qMsgOutput.Body = "Preprocessor Started"
            '�� ���� ��������� ������� ���� ����������� ��������� ����
        For intCount = 1 To gNetPreprocNum
            ' ���������� ���� � ������� ������������ ���������
            qInfoOutput.FormatName = "DIRECT=OS:" + _
            gSocketNet(intCount) + "\Private$\GeneralQueue"
            ' ������� ������� ��������� � ����������� (��� ��������
            '   ���������, ������ � ������� �������� ����)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' �������� ���������
            qMsgOutput.Send qQueueOutput
            ' ������� ������� ���������
            qQueueOutput.Close
        Next
        
    End If
            ' ��������� �� ��������� ����� "frmTableCalendar"
    Load frmTableCalendar
            ' ��������� �� ��������� ����� "frmTableTime"
    Load frmTableTime
            ' ��������� �� ��������� ����� "frmTableTerminal"
    Load frmTableTerminal
             ' ���� ������� ����� ��������������� ���������� �����������
             '  � ���������� ������� "�����������", �� ��������� �� ��������� �����
             '  "frmDataParkingIn", "frmDataParkingOut", "frmDataParkingInfo,
             '  "frmDataParkingServ � "frmMinus"
    If (frmDemo.chkTerm(0).Value = 1 And intParking(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intParking(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intParking(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intParking(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataParkingIn
        Load frmDataParkingOut
            '��������� ������� �� ������ ���������� ������ ��� �����������
            '  �������� � ����������� �������
        frmDataParkingOut.cmdOutConst.Caption = "Sant=""" + Str(gTariffConst) + """"
        Load frmDataParkingInfo
        Load frmDataParkingServ
        Load frmMinus
            '������� ��������� ���� "Parking" ����� "frmDemo"
        frmDemo.mnuParking.Enabled = True
           '������� �������� �������� ���������� "������������"
        frmDemo.imgParkingIn.Visible = True
        frmDemo.imgParkingOut.Visible = True
        frmDemo.imgParkingInfo.Visible = True
        frmDemo.imgParkingServ.Visible = True
            '����� "������ ���������" ��� �����������/���������� ������� �����������
        frmDataParkingIn.chkDocument.Value = intDocument
        frmDataParkingOut.chkDocument.Value = intDocument
        frmDataParkingInfo.chkDocument.Value = intDocument
        frmDataParkingServ.chkDocument.Value = intDocument
            '������� ���������������� ���� ��� �������� �������� ��������
        If gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7 Then
            frmDemo.prtPortDocument.PortOpen = True
        End If
            '������� ���������������� ���� ��� �������� �����-����
        If gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7 Then
            frmDemo.prtPortBarCode.PortOpen = True
        End If
            '������� ���������������� ���� ��� �������-��������� ��������� ����
        If gParkingPlaceNum <> 0 Then
            frmDemo.prtPortDisplay.PortOpen = True
        End If
            '���������� ������ ����������
        frmDataParkingIn.txtParole.Tag = strParole
        frmDataParkingOut.txtParole.Tag = strParole
        frmDataParkingInfo.txtParole.Tag = strParole
            '���������� ����� "TimeOut' ��� ����� ������
        frmDataParkingIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataParkingOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataParkingInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
           '������� �������� �������� ���������� "������������"
        If frmDemo.chkTerm(0).Value = 1 And intParking(0) = 1 Then
            If gTermInp = 0 Then frmDemo.imgParkingInData(0).Visible = True
            If gTermOut = 0 Then frmDemo.imgParkingOutData(0).Visible = True
            frmDemo.imgParkingInfoData(0).Visible = True
        End If
        If frmDemo.chkTerm(1).Value = 1 And intParking(1) = 1 Then
            If gTermInp = 1 Then frmDemo.imgParkingInData(1).Visible = True
            If gTermOut = 1 Then frmDemo.imgParkingOutData(1).Visible = True
            frmDemo.imgParkingInfoData(1).Visible = True
        End If
        If frmDemo.chkTerm(2).Value = 1 And intParking(2) = 1 Then
            If gTermInp = 2 Then frmDemo.imgParkingInData(2).Visible = True
            If gTermOut = 2 Then frmDemo.imgParkingOutData(2).Visible = True
            frmDemo.imgParkingInfoData(2).Visible = True
        End If
        If frmDemo.chkTerm(3).Value = 1 And intParking(3) = 1 Then
            If gTermInp = 3 Then frmDemo.imgParkingInData(3).Visible = True
            If gTermOut = 3 Then frmDemo.imgParkingOutData(3).Visible = True
            frmDemo.imgParkingInfoData(3).Visible = True
        End If
    Else
                '������� ����������� ���� "Parking" ����� "frmDemo"
        frmDemo.mnuParking.Enabled = False
           '������� ���������� �������� ���������� "������������"
        frmDemo.imgParkingIn.Visible = False
        frmDemo.imgParkingOut.Visible = False
        frmDemo.imgParkingInfo.Visible = False
        frmDemo.imgParkingServ.Visible = False
    End If
           
             ' ���� ������� ����� ��������������� ���������� �����������
             '  � ���������� ������� "����������", �� ��������� �� ��������� �����
             '  "frmDataAccessIn", "frmDataAccessOut", "frmDataAccessInfo",
             '  "frmDataAccessServ", "frmLease" � "frmMinus"
    If (frmDemo.chkTerm(0).Value = 1 And intAccess(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intAccess(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intAccess(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intAccess(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataAccessIn
        Load frmDataAccessOut
            '��������� ������� �� ������ ���������� ������ ��� �����������
            '  �������� � ����������� �������
        frmDataAccessOut.cmdOutConst.Caption = "Sant=""" + Str(gTariffConst) + """"
        Load frmDataAccessInfo
        Load frmDataAccessServ
        Load frmLease
        Load frmMinus
            '������� ��������� ���� "Access" ����� "frmDemo"
        frmDemo.mnuAccess.Enabled = True
           '������� �������� �������� ���������� "������������"
        frmDemo.imgAccessIn.Visible = True
        frmDemo.imgAccessOut.Visible = True
        frmDemo.imgAccessInfo.Visible = True
        frmDemo.imgAccessServ.Visible = True
            '����� "������ ���������" ��� �����������/���������� �������
        frmDataAccessIn.chkDocument.Value = intDocument
        frmDataAccessOut.chkDocument.Value = intDocument
        frmDataAccessInfo.chkDocument.Value = intDocument
        frmDataAccessServ.chkDocument.Value = intDocument
            '������� ���������������� ���� ��� �������� �������� ��������
        If (gDocument = 1 Or gDocument = 3 Or gDocument = 5 Or gDocument = 7) And _
        frmDemo.prtPortDocument.PortOpen = False Then
            frmDemo.prtPortDocument.PortOpen = True
        End If
            '������� ���������������� ���� ��� �������� �����-����
        If (gDocument = 2 Or gDocument = 3 Or gDocument = 6 Or gDocument = 7) And _
        frmDemo.prtPortBarCode.PortOpen = False Then
            frmDemo.prtPortBarCode.PortOpen = True
        End If
            '������� ���������������� ���� ��� �������-��������� ��������� ����
        If gAccessPlaceNum <> 0 And frmDemo.prtPortDisplay.PortOpen = False Then
            frmDemo.prtPortDisplay.PortOpen = True
        End If
            '���������� ������ ����������
        frmDataAccessIn.txtParole.Tag = strParole
        frmDataAccessOut.txtParole.Tag = strParole
        frmDataAccessInfo.txtParole.Tag = strParole
            '���������� ����� "TimeOut' ��� ����� ������
        frmDataAccessIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataAccessOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataAccessInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
           '������� �������� �������� ���������� "������������"
        If frmDemo.chkTerm(0).Value = 1 And intAccess(0) = 1 Then
            If gTermInp = 0 Then frmDemo.imgAccessInData(0).Visible = True
            If gTermOut = 0 Then frmDemo.imgAccessOutData(0).Visible = True
            frmDemo.imgAccessInfoData(0).Visible = True
        ElseIf frmDemo.chkTerm(1).Value = 1 And intAccess(1) = 1 Then
            If gTermInp = 1 Then frmDemo.imgAccessInData(1).Visible = True
            If gTermOut = 1 Then frmDemo.imgAccessOutData(1).Visible = True
            frmDemo.imgAccessInfoData(1).Visible = True
        ElseIf frmDemo.chkTerm(2).Value = 1 And intAccess(2) = 1 Then
            If gTermInp = 2 Then frmDemo.imgAccessInData(2).Visible = True
            If gTermOut = 2 Then frmDemo.imgAccessOutData(2).Visible = True
            frmDemo.imgAccessInfoData(2).Visible = True
        ElseIf frmDemo.chkTerm(3).Value = 1 And intAccess(3) = 1 Then
            If gTermInp = 3 Then frmDemo.imgAccessInData(3).Visible = True
            If gTermOut = 3 Then frmDemo.imgAccessOutData(3).Visible = True
            frmDemo.imgAccessInfoData(3).Visible = True
        End If
    Else
                '������� ����������� ���� "Access" ����� "frmDemo"
        frmDemo.mnuAccess.Enabled = False
           '������� ���������� �������� ���������� "������������"
        frmDemo.imgAccessIn.Visible = False
        frmDemo.imgAccessOut.Visible = False
        frmDemo.imgAccessInfo.Visible = False
        frmDemo.imgAccessServ.Visible = False
    End If
           
             ' ���� ������� ����� ��������������� ���������� �����������
             '  � ���������� ������� "��������", �� ��������� �� ��������� �����
             '  "frmDataEmployeIn", "frmDataEmployeOut" � "frmDataAccessInfo
    If (frmDemo.chkTerm(0).Value = 1 And intEmploye(0) = 1 Or _
    frmDemo.chkTerm(1).Value = 1 And intEmploye(1) = 1 Or _
    frmDemo.chkTerm(2).Value = 1 And intEmploye(2) = 1 Or _
    frmDemo.chkTerm(3).Value = 1 And intEmploye(3) = 1) And _
    frmDemo.optAutomatic = True Then
        Load frmDataEmployeIn
        Load frmDataEmployeOut
        Load frmDataEmployeInfo
            '������� ��������� ���� "Employe" ����� "frmDemo"
        frmDemo.mnuEmploye.Enabled = True
           '������� �������� �������� ���������� "���������" (����������� � ��.)
        frmDemo.imgEmployeInData.Visible = True
        frmDemo.imgEmployeOutData.Visible = True
        frmDemo.imgEmployeInfoData.Visible = True
            '���������� ������ ����������
        frmDataEmployeIn.txtParole.Tag = strParole
        frmDataEmployeOut.txtParole.Tag = strParole
        frmDataEmployeInfo.txtParole.Tag = strParole
            '���������� ����� "TimeOut' ��� ����� ������
        frmDataEmployeIn.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataEmployeOut.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
        frmDataEmployeInfo.tmrParoleTimeOut.Interval = frmDemo.tmrPasswTimeOut.Interval
    Else
                '������� ����������� ���� "Employe" ����� "frmDemo"
        frmDemo.mnuEmploye.Enabled = False
    End If
           
           '������� ������� ����� 'frmDemo"
    frmDemo.Visible = True
            '���������� ����� �� ����� "Dummy"
    If frmDemo.Visible = True Then frmDemo.chkDummy.SetFocus
    
            '��������� � ��������� ������� ����
    frmTableCalendar.Tag = Trim(Format(Now, "dd/mm/yyyy"))
            '������� �����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    intHour = Hour(gProtocol.strProtocTime)
    If intHour < 10 Then
        strHour = "0" + Trim(Str(intHour))
    Else
        strHour = Trim(Str(intHour))
    End If
            '������
    intMinute = Minute(gProtocol.strProtocTime)
    If intMinute < 10 Then
        strMinute = "0" + Trim(Str(intMinute))
    Else
        strMinute = Trim(Str(intMinute))
    End If
            
            '����� ���� � �������
    frmDemo.lblTime.Caption = "   " + frmTableCalendar.Tag + "   " _
    + strHour + ":" + strMinute
            
    
            '��������� ������ ��������� �������
    frmTableCalendar.tmrMinute.Enabled = True
            '������ ������ ����������
    gTermContr = 0
            '�������� ������ ������������ ������ "Controller'��"
            '  �������� ������ ������������� ����
    If frmDemo.chkTerm(0).Value = 1 Or frmDemo.chkTerm(1).Value = 1 Or _
    frmDemo.chkTerm(2).Value = 1 Or frmDemo.chkTerm(3).Value = 1 Then
        frmDemo.tmrTermContr.Enabled = True
    End If
            '����� ����������� ���������� � �������� (��������) ���������
            '  � ������� ���������������� ����� ��� ���������� ����������
    If frmDemo.chkTerm(0).Value = 1 Then
        frmDemo.prtPortC(0).PortOpen = True
        frmDemo.imgViewClose(0).Visible = True
    End If
    If frmDemo.chkTerm(1).Value = 1 Then
        frmDemo.prtPortC(1).PortOpen = True
        frmDemo.imgViewClose(1).Visible = True
    End If
    If frmDemo.chkTerm(2).Value = 1 Then
        frmDemo.prtPortC(2).PortOpen = True
        frmDemo.imgViewClose(2).Visible = True
    End If
    If frmDemo.chkTerm(3).Value = 1 Then
        frmDemo.prtPortC(3).PortOpen = True
        frmDemo.imgViewClose(3).Visible = True
    End If
            '���������� ������ "Controller'��" ��� "������"
            '  ���������� ����������� �� ���������� "������"
    If frmDemo.optManual = True Then
        frmDemo.cmdOpen(0).Tag = CByte(CInt(Trim(gAddrManual(0))))
        frmDemo.cmdOpen(0).Caption = Trim(gAddrManual(0))
        frmDemo.cmdOpen(1).Tag = CByte(CInt(Trim(gAddrManual(1))))
        frmDemo.cmdOpen(1).Caption = Trim(gAddrManual(1))
        frmDemo.cmdOpen(2).Tag = CByte(CInt(Trim(gAddrManual(2))))
        frmDemo.cmdOpen(2).Caption = Trim(gAddrManual(2))
        frmDemo.cmdOpen(3).Tag = CByte(CInt(Trim(gAddrManual(3))))
        frmDemo.cmdOpen(3).Caption = Trim(gAddrManual(3))
    End If
            
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
    If gParkingPlaceNum <> 0 Then
            '������������� �������-���������
        strMessage = "ParkFreePlaces=" + CStr(gParkFreePlaces)
        Call frmDemo.Display(strMessage)
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
    ElseIf gAccessPlaceNum <> 0 Then
            '������������� �������-���������
        strMessage = "AcceFreePlaces=" + CStr(gParkFreePlaces)
        Call frmDemo.Display(strMessage)
    End If
            
            '������ ����������� ���������� � �������� ���������
    frmDemo.imgViewOpen(0).Visible = False
    frmDemo.imgViewOpen(1).Visible = False
    frmDemo.imgViewOpen(2).Visible = False
    frmDemo.imgViewOpen(3).Visible = False
            ' ����� "������" ���������������
    frmDemo.imgPhoto(0).Picture = LoadPicture("")
    frmDemo.imgPhoto(1).Picture = LoadPicture("")
    frmDemo.imgPhoto(2).Picture = LoadPicture("")
    frmDemo.imgPhoto(3).Picture = LoadPicture("")
            '�������� ������� "����� ����"
    intWhite = 0
           '��������� ���������� ������� � �������� ����������� ��������� ����
    intControlIndex = 0
    
               '�������� ����� ������� ������ "������� ������"
    frmDemo.tmrTermContr.Tag = 0
    frmDemo.lblInform(0).Tag = 0
    frmDemo.lblInform(1).Tag = 0
    frmDemo.lblInform(2).Tag = 0
    frmDemo.lblInform(3).Tag = 0
               '���������� ������ ����������
    gTermContr = 1
            ' ���� ������ "Controller'��" ����������
            '    (����������� ��� ������� ������ "cmdExit")
    Do While DoEvents()
            
            '���� ������ ActiveX.EXE � "Host Computer'e" ��������
        If objTablePerson Is Nothing Then
            ' ���� ��� "Host Computer"
            If gPreprocName = "" Then
            '������������ ������ � "Host Computer'e"
                Set objTablePerson = New XTablePerson
                Set gTablePerson = objTablePerson
            End If
        End If
               
               '�������� ����� ������� ������ "������� ������"
        frmDemo.tmrTermContr.Tag = intControlIndex
        frmDemo.lblInform(intControlIndex).Tag = 0
            
            '����� �� "Controller'��" �������� ������
            '  ��������� ������������� ����
            
        If frmDemo.prtPortC(intControlIndex).InBufferCount >= _
        gPersonCode(intControlIndex) Then
           '���������� ������ � �������� ����� ��� ���������� ���������
            vntBufferInput(intControlIndex) = frmDemo.prtPortC(intControlIndex).Input
            '����� "��������� ������������������ �������� ���������"
            Call InitialOpenTerminal(intControlIndex)
        End If
        
'''        DoEvents
        
            '������� ������� �� ������������ ����������
        If frmDemo.prtPortC(intControlIndex).Tag > 0 Then
            '����� ��������� �������� ��� �����, ���������������
            ' �������� ����� � ������� "������� ����������"
            intWork = (frmDemo.prtPortC(intControlIndex).CommPort - 2) * 15
            '�� ���� �������� ������ ������� ��� �������� �����
            For intCount = 1 To 15 Step 1
            '����� �������� �������� ������� "������� ����������",
            '  ��������� ������
                intRequest = intWork + intCount
            '������ �� "��������� ������������������ �������� ���������"
                If Mid(gAddrPort(0, intRequest), 4) = "A" Then
            '����� "��������� ������������������ �������� ���������"
                    Call InitialCloseTerminal(intControlIndex, intRequest)
            '������ �� "������������������ �������� �������� ���������"
                ElseIf Mid(gAddrPort(0, intRequest), 4) = "V" Then
            '����� "������������������ �������� �������� ���������"
                    Call WaitCloseTerminal(intControlIndex, intRequest)
            '������ �� "������������������ �������� ���������"
            '  �� ����������� "������"
                ElseIf Mid(gAddrPort(0, intRequest), 4) = "1" Then
            '����� "������������������ �������� ���������"
            '  �� ����������� "������"
                    Call ButtonOpenTerminal(intControlIndex, intRequest)
                End If
            Next
        End If
            
            'B������ ����������� ��������� ���������
        If frmDemo.tmrButton(intControlIndex).Tag = 1 Then
            PictureTerminalClose intControlIndex
            '�������� ������� "TimeOut" ��� ����������� "������"
            frmDemo.tmrButton(intControlIndex).Tag = 0
        End If
        
        If intControlIndex < 3 Then
            intControlIndex = intControlIndex + 1
        Else
            intControlIndex = 0
        End If
            '���� � ������� ������� ������ - �������� � ���������� �����
        Do While frmDemo.prtPortC(intControlIndex).PortOpen = False And _
        frmDemo.chkSetup.Value = 1
          '���������� ����� �������� � �������� ����������� ��������� ����
            If intControlIndex < 3 Then
                intControlIndex = intControlIndex + 1
            Else
                intControlIndex = 0
                Exit Do
            End If
        Loop
        
    Loop
    
End Sub
            
            '��������� ������������������ �������� ���������
Private Sub InitialOpenTerminal(intIndex As Integer)
            '����� �����, ����� ������� �����������
            '  ������������ ��� �� "Controller'a"
Dim vntReadPortNum As Variant
            '��� �������� ��� ������� ���������� ���������������
Dim intAutoRegistrCode As Integer
            '��� �������� ��� ������� ���������� ������������
Dim intAutoDeleteCode As Integer
            '��� �������� ������� ��������� ������ "Reserve" � "������� ������"
            '  ����� ������/������ ������� �����������
Dim intParkingCode As Integer
            '��� �������� ������� ��������� ������ "Reserve" � "������� ������"
            '  ����� �����/������ ���������� �����������
Dim intAccessCode As Integer
            '��� �������� ������� ��������� ������ "Name" � "������� ������"
            '  ����� �����/������ ��������� �����������
Dim intEmployeCode As Integer
            '��� �������� ��� ������� ������� �������
Dim intStatusCode As Integer
            '��� �������� ��� ������� ��� �������
Dim intCalendarCode As Integer
            '��� �������� ��� ������� ������� �������
Dim intTimeCode As Integer
            '��� �������� ��� ������� ��������� �������
Dim intTerminalCode As Integer
            '��� �������� ��� �������� ��������� ��� ���� ���������
Dim intScriptCode As Integer
            '����� �����������
Dim vntAddr As Variant
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������� ����
Dim vntWork As Variant
            
            '����������� ��� �������������� ������� ������
            ' �� ������ � ����������������� �������������
Dim intCicle As Integer
Dim strBuffer As String
Dim intBuffer1 As Integer
Dim intBuffer2 As Integer
            
            '����� �����, ����������� ������������ ��� �� "Controller'a"
    vntReadPortNum = frmDemo.prtPortC(intIndex).CommPort
            '���������� ����� "Controller'a", ����������� ������������ ���
    vntAddr = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(15)
            '����� �������� �������� � ������� "������� ����������"
    intRequest = (vntReadPortNum - 2) * 15 + vntAddr
            
            '������ ��������� ���������������� "Controller'a"
            
            ' "Controller" ��������� �������� �� �������
    If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then GoTo WaitCycle
            ' "Controller" ����� ������������� ��������� � ��� ��������
            '  ���������� ���������� �� ����������� - ����� �� ���������
    If Mid(gAddrPort(0, intRequest), 4) <> "0" And _
    Mid(gAddrPort(0, intRequest), 4) <> "#" Then GoTo WaitCycle
            '�������� ������� "����� ����"
    intWhite = 0
    
            '�������� ���� ��������� ����������� ��
            '  ����������� ������������� ����
    vntWork = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(96)
            ' "Controller" �� ������ PIN ��� (�� ������� ��������� �� �����������)
    If vntWork <> 32 Then
            ' "Controller" ������������ ������������ ���, ������� ��������
            '   � ��� ��������� ������ (������� ��������� �� �����������) � ���
            '   �������� ���������� ���������� �� �����������
        If vntWork = 0 And Mid(gAddrPort(0, intRequest), 4) <> "#" Then
            '���������� ������� "����� ����"
            intWhite = 1
            GoTo Continue
            ' "Controller" ��������� ������� ������� �������� (�� ����������
            '  ��� ���� ����������� - ��� ����������� ������������� ����,
            '  ������� �������� � ��������� ������ "Controller'a") � �������
            '  ���������� ���������� �� �����������
        ElseIf vntWork = 0 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            GoTo WaitCycle
            ' "Controller" ����� ������� PIN ��� (��������� ��������� �� �����������)
            '   � ������� ���������� ���������� �� �����������
        ElseIf vntWork = 64 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            '�������� ������������� ������� �������� "Controller'a", �
            '   �������� ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            GoTo WaitCycle
            ' "Controller" ����� ������� PIN ��� (��������� ��������� �� �����������)
            '   � �� ������� ���������� ���������� �� �����������
        ElseIf vntWork = 64 And Mid(gAddrPort(0, intRequest), 4) = "0" Then
            GoTo WaitCycle
            ' "Controller" ��������� � ��������� ����������������
            '  (����� ������� � ��������� ���������� �� �����������)
        ElseIf vntWork = 96 Then
            '�������� ��������� ������� �������� "Controller'a", �
            '   �������� ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            GoTo WaitCycle
        End If
            ' "Controller" ������ PIN ��� (������� ��������� �� �����������)
            '   � ������� ���������� ���������� �� �����������
    ElseIf vntWork = 32 And Mid(gAddrPort(0, intRequest), 4) = "#" Then
            '�������� ������������� ������� �������� "Controller'a", �
            '   �������� ��������� ��������� ��������� �� �����������
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
        GoTo Continue
            ' "Controller" ������ PIN ��� (������� ��������� �� �����������)
            '   � �� ������� ���������� ���������� �� �����������
    ElseIf vntWork = 32 And Mid(gAddrPort(0, intRequest), 4) = "0" Then
        GoTo Continue
    End If
    
Continue:
            
            '���������� ��������� ������������� ����
            
            '� "Controller'f" ���� �������� �������
            '  ��� ���������� ����������� ������� �����
    If vntAddr = 0 Then
            '������� ������������ ��� � ����� "Controller'a"
        vntBufferInput(intIndex) = "0000000000000000"
            '������������ ��� � ������� ����� "N_?"
        frmDemo.lblInform(intIndex).Caption = vntBufferInput(intIndex)
            '����� "N_?" - (����� ���)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '������������ ��� � ����� "Controller'a"
        gProtocol.strProtocName = vntBufferInput(intIndex)
            '������������ ��� � ����� "Controller'a"
        gProtocol.strProtocPersonCode = vntBufferInput(intIndex)
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "POWER ON or ADDR=0"
            '�������� ������ � ���� "������� ���������"
        GoTo Protocol
            '� "Controller'y" ����������� ����������� "PROXIMITY GP30"
    ElseIf CByte(Asc(Right(vntBufferInput(intIndex), 1))) = CByte(3) Then
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "00000" + _
        Mid(vntBufferInput(intIndex), 4, 10)
            '� "Controller'y" ����������� ����������� �����-����"VS800"
    ElseIf Right(vntBufferInput(intIndex), 1) = "*" Then
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "00000" + _
        Mid(vntBufferInput(intIndex), 5, 10)
            '� "Controller'y" ����������� ����������� "TM DALLAS" ��� "1990A"
    ElseIf CByte(Asc(Mid(vntBufferInput(intIndex), 9, 1))) = CByte(1) Then
            '������������� ������ �� ������ ��������� � ����������������� ���
        strBuffer = ""
        intCicle = 1
        Do While intCicle <= 6
            intBuffer1 = (CByte(Asc(Mid(vntBufferInput(intIndex), intCicle + 9, 1))) And CByte(240)) / 16
            intBuffer2 = CByte(Asc(Mid(vntBufferInput(intIndex), intCicle + 9, 1))) And CByte(15)
            strBuffer = Hex(intBuffer1) + Hex(intBuffer2) + strBuffer
            intCicle = intCicle + 1
        Loop
        vntBufferInput(intIndex) = Left(vntBufferInput(intIndex), 1) + "000" + Trim(strBuffer)
            '� "Controller'�" ����������� ������������ ��� �����������
    ElseIf vntAddr <> 0 Then
            '�������������� � ����� "Controller'a" � ������� ����� "N_?"
        frmDemo.lblInform(intIndex).Caption = CStr(vntAddr) + "||" + "UndefinedErr"
            '����� "N_?" - (����� ���)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
            '�������������� � ����� "Controller'a"
        gProtocol.strProtocName = CStr(vntAddr) + "||" + "ErReaderType"
            '�������������� � ����� "Controller'a"
        gProtocol.strProtocPersonCode = CStr(vntAddr) + "||" + "CoflictComm"
            '������
        gProtocol.strProtocStatus = ""
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "UNDEFINED ERROR"
            '�������� ������ � ���� "������� ���������"
        GoTo Protocol
    End If
            
            '����� ��������� � ��������������� - ������ � ������
    frmDemo.lblErrorInpOut(intIndex).Visible = False
    frmDemo.lblErrorBarCodePrinter.Visible = False
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessPlus
            
            '�������� ����� ����������� �� ����������� ������������� ����
    vntWork = CByte(Asc(Left(vntBufferInput(intIndex), 1))) And CByte(16)
    If vntWork = 16 Then vntWork = 1
            '������������ ��� �������� �� "Controller'a" � "Port'a"
            '  ���������������/��������
    If vntAddr * 10 + vntReadPortNum = intAutoRegDel(intIndex) Then
            '������������ ��� ��� ������ "Controller'a"
        gProtocol.strProtocPersonCode = "0" + Mid(vntBufferInput(intIndex), 2, _
        gPersonCode(intIndex) - 1)
            '����� �����������="0" - ���������������
        If vntWork = 0 Then
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ��������������� ������� �� �����������
            If intParking(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoParkReg(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ��������������� ���������� �����������
            ElseIf intAccess(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoAcceReg(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ��������������� ��������� �����������
            ElseIf intEmploye(intIndex) = 1 Then
                intAutoRegistrCode = _
                frmDemo.AutoEmplReg(gProtocol.strProtocPersonCode)
                GoTo WaitCycle
            End If
            '����� �����������="1" - ������������
        Else
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ������������ ������� � �����������
            If intParking(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoParkDel(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ������������ ���������� �����������
            ElseIf intAccess(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoAcceDel(gProtocol.strProtocPersonCode, intIndex)
                GoTo WaitCycle
            '����� ���������-������� ������������� ����� ������������� ����
            '  � ������ ������ ��� ������������ ��������� �����������
            ElseIf intEmploye(intIndex) = 1 Then
                intAutoDeleteCode = _
                frmDemo.AutoEmplDel(gProtocol.strProtocPersonCode)
                GoTo WaitCycle
            End If
        End If
    End If
    
'''            '����������� ����� "����" ��� ���������� �����
'''    If gSluice(intIndex) <> 0 Then
'''            '����������� ��� "Controller'�" ����� - ���������� � ��������� �������
'''        frmDemo.prtPortC(intIndex).Output = Chr(192)
'''             '����� ���������� �������� ������� ����������
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''            '�������� ��������������� "Controller" - ����� � ����������� �������
'''        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
'''             '����� ���������� �������� ������� �����
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''    End If

            '������������ ��� � ������� ����� "N_?"
    frmDemo.lblInform(intIndex).Caption = CStr(vntAddr) + "||" + _
    Mid(vntBufferInput(intIndex), 7, gPersonCode(intIndex) - 6)
    
            '������������ ��� ��� ������ "Controller'a"
    gProtocol.strProtocPersonCode = "0" + Mid(vntBufferInput(intIndex), 2, _
    gPersonCode(intIndex) - 1)
        '������� ������� "������� ������" = 1 (������������ ���)
    gTablePerson.Col = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ������"
        If Trim(gTablePerson.Text) = gProtocol.strProtocPersonCode Then
            '��������� ����� �� �����
            Exit For
        End If
    Next
            
            ' ������� ���� ���������������
    If frmDemo.chkPhoto(intIndex).Value = 1 Then
        frmDemo.imgPhoto(intIndex).Picture = LoadPicture("")
    End If
            
            '���������������� �������
            
            '���������������� ������������ ���
    If intRowNum = gTablePerson.Rows Then
                '��������������� ��� (����� ������� ����������� ������������ ���)
        gProtocol.strProtocName = "@"
            '���������� ������� "����� ����"
        If intWhite = 1 Then
            '���������� - "����� ����"
            
            '����� �����������="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            '����� �����������="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            End If
            
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            '�������� ������� "����� ����"
            intWhite = 0
            '������� "����� ����" �� ���������� - (�������� ������������ ���)
        Else
            '����� �����������="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  INVALID KEY"
            '����� �����������="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  INVALID KEY"
            End If
        
        End If
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
        gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (����� ���)
        frmDemo.lblInform(intIndex).BackColor = vbWhite
        GoTo ResetController
            '�������������� ������������ ���
    Else
            
            '��������� ����� ������� ������ "������� ������"
        frmDemo.lblInform(intIndex).Tag = intRowNum
                
            '������� ������� "������� ������" = 0 (������� ��� ��������)
        gTablePerson.Col = 0
        gProtocol.strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
        gTablePerson.Col = 2
        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
        intHour = Hour(gProtocol.strProtocTime)
        intMinute = Minute(gProtocol.strProtocTime)
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '���������� �� ������� ����� "N_?"
        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
    
            '����� ����������� ��������������� �� ����� "Photo" �����������
        If frmDemo.chkPhoto(intIndex).Value = 1 Then
            '�������� ��������� ������ �����/������
            On Error GoTo PhotoError
            '�� ���������� ������� "�����������" ��� ������� "�����������"
            If intParking(intIndex) = 0 And intAccess(intIndex) = 0 Then
                frmDemo.imgPhoto(intIndex).Picture = LoadPicture("C:\Photo\" + _
                Trim(Left(gProtocol.strProtocName, 15)) + ".bmp")
            '���������� ������� "�����������" ��� ������� "�����������"
            Else
                frmDemo.imgPhoto(intIndex).Picture = LoadPicture("C:\Photo\" + _
                Trim(gProtocol.strProtocName) + ".bmp")
            End If
            GoTo PhotoOK
PhotoError:
            Resume PhotoOK
PhotoOK:
            On Error GoTo 0
        End If
            
            '������ ������� �������
        intStatusCode = StatusCode()
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���������� ������� "����� ����", �� ������
            '  �� ����� ��������� �������� - ��������������
        If intWhite = 1 And intStatusCode <> 0 Then
            
            '����� �����������="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            '����� �����������="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  WHITE KEY"
            End If
            
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            '�������� ������� "����� ����"
            intWhite = 0
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (����� ���)
            frmDemo.lblInform(intIndex).BackColor = vbWhite
            '�� ������ ����
            If intStatusCode <> 2 Then GoTo ResetController
        
        End If
            '������ ����� ��������� �������� - �� ����������� ��� �����������
            '  � �������� �����������
        If intStatusCode = 0 And intParking(intIndex) = 0 And intAccess(intIndex) = 0 Then
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (����� ���)
            frmDemo.lblInform(intIndex).BackColor = vbBlue
            GoTo ResetController
            '������ ����
        ElseIf intStatusCode = 2 Then
            '���������������� ������� - (������ ����)
            
            '����� �����������="0"
            If vntWork = 0 Then
                gProtocol.strProtocReserve = "0/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  BLACK KEY"
            '����� �����������="1"
            Else
                gProtocol.strProtocReserve = "1/" + CStr(vntAddr) + "/" + _
                CStr(vntReadPortNum) + "  BLACK KEY"
            End If
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
            gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������� ���)
            frmDemo.lblInform(intIndex).BackColor = vbRed
            GoTo ResetController
            '������ "Relay" - ������ ��� ������ ������������� ������ � "��������������"
            '  ���������� ����������� - �� ����������� ��� �����������, ��������
            '  ����������� ��� ��������� ��� ��������
        ElseIf intStatusCode = 3 And frmDemo.optAutomatic = True And _
        intParking(intIndex) = 0 And intAccess(intIndex) = 0 And intEmploye(intIndex) = 0 Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
            If Trim(gTablePerson.Text) <> "Dallas" And _
            frmDemo.cmdOpen(intIndex).Tag = 0 Then
'''            '������ ����������� ��������� ������� ��� ����������� ��������
'''                frmDemo.BeepSound
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '������� ����������� "������" �������� ���������
                frmDemo.cmdOpen(intIndex).Enabled = True
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
                gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '����� "N_?" - (������ ���)
                frmDemo.lblInform(intIndex).BackColor = vbYellow
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(intIndex).Enabled = True
                GoTo ResetController
            End If
            
            '������ "Relay" - ������ ��� "��������������" ���������� �����������
            '  � ��� ����������a ��� ������e ����������e
        ElseIf intStatusCode = 3 And frmDemo.optAutomatic = True And _
        (intParking(intIndex) = 1 Or intAccess(intIndex) = 1) Then
            '������� ������� "������� ������" = 0 (������� ��� ��������)
            gTablePerson.Col = 0
            '����������a ��� ������e ����������e
            If Trim(gTablePerson.Text) = "DallasParkAcce" Then
            '���� ��� ��������� ���� - ������������ �������
                If (gParkingPlaceNum <> 0 And gParkFreePlaces = 0) Then _
                GoTo WaitCycle
            '������� ������� "������� ������" = 5 (Addr Port Type)
                gTablePerson.Col = 5
            '���� ��� ��������� ������ ����������� ��� �����������
                If Mid(Trim(gTablePerson.Text), 4) <> "CONTR" Then
            '����� ���������-������� ������������ ������������� ����
            '  � ������ ������ ��� ��������������� ��������
            '  (����� ����������� "Controller" � ������� "Dallas")
                    intAutoRegistrCode = _
                    frmDemo.AutoRegDallasButton(gProtocol.strProtocPersonCode, _
                    intIndex, Trim(gTablePerson.Text))
            '��������� ������ ��������� ������� �� ���������� ������ "Dallas"
            '  ��� ������������ �������� �� ������� ��� �������� ������
                    frmDemo.tmrRelay.Interval = _
                    frmDemo.tmrButton(intIndex).Interval
                    frmDemo.tmrRelay.Tag = 0
                    frmDemo.tmrRelay.Enabled = True
            '������ ������ ����������
                    gTermContr = 0
            '���� ��������  ��������� ������� ��������
                    Do While frmDemo.tmrRelay.Tag = 0
            '���������� ��������� �������
                        DoEvents
                    Loop
                    frmDemo.tmrRelay.Enabled = False
            '���������� ������ ����������
                    gTermContr = 1
                    GoTo WaitCycle
            '������ "������" ��������� ���� "CONTR" - ���������� �������
                ElseIf Right(Trim(gTablePerson.Text), 5) = "CONTR" Then
            
            ''' ��������
            
                    GoTo WaitCycle
                End If
            
            End If
            
            '������ ����� ������������ ���������� ����������, � �����
            '  ���� ��� �����������, ������� ����������� ���
            '  ��������� ��� ��������
        ElseIf intStatusCode = 1 And intParking(intIndex) = 0 And intAccess(intIndex) = 0 And _
        intEmploye(intIndex) = 0 Or _
        (intStatusCode = 5 Or intStatusCode = 6 Or intStatusCode = 7) And _
        intParking(intIndex) = 1 Or _
        (intStatusCode = 8 Or intStatusCode = 9 Or intStatusCode = 10) And _
        intAccess(intIndex) = 1 Or _
        (intStatusCode = 0 Or intStatusCode = 1) And _
        intEmploye(intIndex) = 1 Then
            '������ ��� �������
            intCalendarCode = CalendarCode()
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '������ ������� �������
            intTimeCode = TimeCode()
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag


 
'##############################################################################
            '������ �������� ������ (� ����� ����) � O���� (�� ��������������) �����
            '  ������� ���������o�
            If intCalendarCode = 0 And intTimeCode = 0 Then
            'E��� ��� �� �����������, �� ������� ����������� � �� ��������� ��� ��������
                If intParking(intIndex) = 0 And intAccess(intIndex) = 0 And intEmploye(intIndex) = 0 Then
            '����� "N_?" - (������� ���)
                    frmDemo.lblInform(intIndex).BackColor = vbGreen
            '�������� �������� ����� �����
                    frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
                    Do
                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
                    intScriptCode = ScriptTermClose(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '��� ��������� ��������� ��������� ��������
                    If intScriptCode = 0 Then
            '�������� �������� ����� �����
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
                    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
                    
            '��� ��������� ��������� ��������� � ���������
            '  �� ������� ������� �������� ���������,
                    If intScriptCode = 0 Then
            '����� �����������="0" - �������� ������ � ��������� (�� ����������)
                        If vntWork = 0 Then
            '��������� ���������� �� ������� ����� "N_?"
                            gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
            '����� �����������="1" - �������� ������ �� ��������� (� ����������)
                        Else
            '��������� ���������� �� ������� ����� "N_?"
                            gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
                        End If
                    End If
                    
'''            '��� ��������� ��������� ��������� ��� ���������
'''            '  �� ������� ������� �������� �� ���������, �����������
'''            '   ����� "����" ��� ���������� �����
'''                    If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            '�������� ��� "Controller'�" ����� - ����� � ��������� �������
'''                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             '����� ���������� �������� ������� �����
'''                        Do
'''                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                    End If
            'E��� ��� �����������,������� ����������� ��� ��������� ��� ��������
                Else
            '������������ ������ ����������� ���, ��������� � ������� ������� - ��� ���������
                    intTimeCode = 0
                    intTerminalCode = 0
                    intCalendarCode = 1
                End If
'##############################################################################

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
            '����������� ����� ������� � O���� (�� ��������������) �����
            '  ������� ���������o�
            ElseIf intCalendarCode = 3 And intTimeCode = 0 Then
            '������ ��������������� ������� �������
                intTimeCode = IndividualTime(vntAddr, vntReadPortNum)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '������ ��������������� ��������� �������
                intTerminalCode = IndividualTerminal(vntAddr, vntReadPortNum)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '������ ��������������� ��� �������
                intCalendarCode = IndividualCalendar()
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '�������������� ����, ����� � �������� ������� - (�����������)
            '  �, ���� ��� �� �����������, �� ������� ����������� � ��
            '  ��������� ��� ��������
                If intCalendarCode = 1 And intTimeCode = 0 _
                And intTerminalCode = 0 And _
                intParking(intIndex) = 0 And intAccess(intIndex) = 0 And _
                intEmploye(intIndex) = 0 Then
            '����� "N_?" - (������� ���)
                    frmDemo.lblInform(intIndex).BackColor = vbGreen
            '�������� �������� ����� �����
                    frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
                    Do
                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
                    intScriptCode = ScriptTermClose(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '��� ��������� ��������� ��������� ��������
                    If intScriptCode = 0 Then
            '�������� �������� ����� �����
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
                    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
                    
            '��� ��������� ��������� ��������� � ���������
            '  �� ������� ������� �������� ���������,
                    If intScriptCode = 0 Then
            '����� �����������="0" - �������� ������ � ���������
                        If vntWork = 0 Then
            '��������� ���������� �� ������� ����� "N_?"
                            gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
            '����� �����������="1" - �������� ������� �� ���������
                        Else
            '��������� ���������� �� ������� ����� "N_?"
                            gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                            CStr(vntReadPortNum) + "Y"
                        End If
                    End If
                    
'''            '��� ��������� ��������� ��������� ��� ���������
'''            '  �� ������� ������� �������� �� ���������, �����������
'''            '   ����� "����" ��� ���������� �����
'''                    If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            '�������� ��� "Controller'�" ����� - ����� � ��������� �������
'''                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             '����� ���������� �������� ������� �����
'''                        Do
'''                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                    End If
                
                End If
            End If
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



            '����, ����� � �������� ������� - (�����������) �
            '  ��� �����������, ������� ����������� ��� ��������� ��� ��������
            If intCalendarCode = 1 And intTimeCode = 0 And _
            intTerminalCode = 0 And _
            (intParking(intIndex) = 1 Or intAccess(intIndex) = 1 _
            Or intEmploye(intIndex) = 1) Then
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(intIndex).BackColor = vbGreen

            '�������� �������� ����� �����
                frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
                intScriptCode = ScriptTermClose(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '��� ��������� ��������� ��������� ��������
                If intScriptCode = 0 Then
                
            '����������� ����� - "�����������"
                    If intParking(intIndex) = 1 Then
            '������ ������ "Reserve" � "������� ������"
                        intParkingCode = frmTablePerson.AnalysisParking(vntWork)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���� ������������ �������� ��� ������� ������ "Reserve"
            '  � "������� ������" (���������� ������ ��� ������� �����/�����)
                        If intParkingCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - ���������� �������� �� �����������
                            If vntWork = 0 Then
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - ���������� �������� c �����������
                            Else
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                If intParkingCode = 2 Then
                                    gProtocol.strProtocReserve = "Extra $?"
                                Else
                                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                    CStr(vntReadPortNum) + "N"
                                End If
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '����������� "������" ��������� �� "Controller'a"
                            intScriptCode = 1
                        Else
            '�������� �������� ����� �����
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            '����������� ����� - "�����������"
                    ElseIf intAccess(intIndex) = 1 Then
            '������ ������ "Reserve" � "������� ������"
                        intAccessCode = frmTablePerson.AnalysisAccess(vntWork)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���� ������������ �������� ��� ������� ������ "Reserve"
            '  � "������� ������" (���������� ������ ��� ������� ����/�����)
                        If intAccessCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - ���������� ������ �� �����������
                            If vntWork = 0 Then
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - ���������� ������� � �����������
                            Else
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                If intParkingCode = 2 Then
                                    gProtocol.strProtocReserve = "Extra $?"
                                Else
                                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                    CStr(vntReadPortNum) + "N"
                                End If
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '����������� "������" ��������� �� "Controller'a"
                            intScriptCode = 1
                        Else
            '�������� �������� ����� �����
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            '����������� ����� - "�������� - �� ���������"
                    ElseIf intEmploye(intIndex) = 1 Then
            '������ ������ "Name" � "������� ������"
                        intEmployeCode = frmTablePerson.AnalysisEmploye(vntWork)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���� ������������ �������� ��� ������� ������ "Name"
            '  � "������� ������" (������� ����/�����)
                        If intEmployeCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - �������� ������ �� �����������
                            If vntWork = 0 Then
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - �������� ������� � �����������
                            Else
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '����������� "������" ��������� �� "Controller'a"
                            intScriptCode = 1
                        Else
            '�������� �������� ����� �����
                            frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                            Do
                            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                            intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
            
            '����� "�����������", "����������" � �������� - �� ���������" �� �����������
                    ElseIf intParking(intIndex) <> 1 And intAccess(intIndex) <> 1 And _
                    intEmploye(intIndex) <> 1 Then
            '�������� �������� ����� �����
                        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
                        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
                        Do
                        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
                        intScriptCode = ScriptOpen(intIndex, vntAddr)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                    End If
            
            '��� ��������� ��������� ��������� � ���������
            '  �� ������� ������� �������� ���������, �����������
            '   ����� - "�����������"
                    If intScriptCode = 0 And intParking(intIndex) = 1 Then
            
            '������� ������� "������� ������" = 0 (���)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 1 (������������ ���)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '���������� �� ������� ����� "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            '����� �����������="0" - ���������� ������ �� �����������
                        If vntWork = 0 Then
            '��������� ������ "Reserve" � "������� ������"
                            intParkingCode = frmTablePerson.InputParking(intIndex)
            '����� �����������="1" - ���������� ������ c �����������
                        Else
            '��������� ������ "Reserve" ��� ���������� ������ � "������� ������"
                            intParkingCode = frmTablePerson.OutputParking(intIndex, intStatusCode)
                        End If
            '���� ������������ �������� ��� ��������� ������"Reserve"
            '  � "������� ������" (���������� ������ ��� ������� �����/�����)
                        If intParkingCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - ���������� �������� �� �����������
                            If vntWork = 0 Then
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - ���������� �������� c �����������
                            Else
            '����� �������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���������� ��������
                        Else
            '����� �����������="0" - ���������� ������ �� �����������
                            If vntWork = 0 Then
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "AutoParking || " + "+"
            '������������ ����������� ���������
                                strMessage = "ParkFreePlaces-1"
            '����� �����������="1" - ���������� ������ c �����������
                            Else
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "AutoParking || " + "-"
            '������������ ����������� ���������
                                strMessage = "ParkFreePlaces+1"
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                       End If
                    
            '��� ��������� ��������� ��������� � ���������
            '  �� ������� ������� �������� ���������, �����������
            '   ����� - "����������"
                    ElseIf intScriptCode = 0 And intAccess(intIndex) = 1 Then
            
            '������� ������� "������� ������" = 0 (���)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 1 (������������ ���)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '���������� �� ������� ����� "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            '����� �����������="0" - ���������� ����� �� �����������
                        If vntWork = 0 Then
            '��������� ������ "Reserve" � "������� ������"
                            intAccessCode = frmTablePerson.InputAccess(intIndex)
            '����� �����������="1" - ���������� ����� � �����������
                        Else
            '��������� ������ "Reserve" ��� ���������� ������ � "������� ������"
                            intAccessCode = frmTablePerson.OutputAccess(intIndex, intStatusCode)
                        End If
            '���� ������������ �������� ��� ��������� ������"Reserve"
            '  � "������� ������" (���������� ������ ��� ������� �����/�����)
                        If intAccessCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - ���������� ������ �� �����������
                            If vntWork = 0 Then
            '����� ��������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - ���������� ������� � �����������
                            Else
            '����� ��������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���������� ��������
                        Else
            '����� �����������="0" - ���������� ����� �� �����������
                            If vntWork = 0 Then
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "AutoAccess || " + "+"
            '������������ ����������� ���������
                                strMessage = "AcceFreePlaces-1"
            '����� �����������="1" - ���������� ����� � �����������
                            Else
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "AutoAccess || " + "-"
            '������������ ����������� ���������
                                strMessage = "AcceFreePlaces+1"
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
                        End If
                    
            '��� ��������� ��������� ��������� � ���������
            '  �� ������� ������� �������� ���������, �����������
            '   ����� - "�������� - �� ���������"
                    ElseIf intScriptCode = 0 And intEmploye(intIndex) = 1 Then
            
            '������� ������� "������� ������" = 0 (���)
                        gTablePerson.Col = 0
                        gProtocol.strProtocName = gTablePerson.Text
            '������� ������� "������� ������" = 1 (������������ ���)
                        gTablePerson.Col = 1
                        gProtocol.strProtocPersonCode = gTablePerson.Text
            '������� ������� "������� ������" = 2 (������)
                        gTablePerson.Col = 2
                        gProtocol.strProtocStatus = Trim(gTablePerson.Text)
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '���������� �� ������� ����� "N_?"
                        gProtocol.strProtocReserve = frmDemo.lblInform(intIndex).Caption
            
            '����� �����������="0" - �������� ����� �� �����������
                        If vntWork = 0 Then
            '��������� ������ "Name" � "������� ������"
                            intEmployeCode = frmTablePerson.InputEmploye(intIndex)
            '����� �����������="1" - �������� ����� � �����������
                        Else
            '��������� ������ "Name" � "������� ������"
                            intEmployeCode = frmTablePerson.OutputEmploye(intIndex)
                        End If
            '���� ������������ �������� ��� ��������� ������"Name"
            '  � "������� ������" (������� ����/�����)
                        If intEmployeCode <> 0 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
                            frmDemo.BeepSound
            '����� �����������="0" - �������� ������ �� �����������
                            If vntWork = 0 Then
            '����� ��������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Input !!! " + gProtocol.strProtocReserve
            '����� �����������="1" - �������� ������� � �����������
                            Else
            '����� ��������� � ��������������� - �� �����
                                frmDemo.lblErrorInpOut(intIndex).Visible = True
            '��������� ���������� �� ������� ����� "N_?"
                                gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                                CStr(vntReadPortNum) + "N"
            '������������ ����������� ���������
                                strMessage = "Error Output !!! " + gProtocol.strProtocReserve
                            End If
            '�������� ���������
                            Call frmDemo.SendMessage(strMessage)
'������������ ����������� ����� ����� ������� ������ "������� ������"
gTablePerson.Row = frmDemo.lblInform(intIndex).Tag
            '���������� ��������
                        Else
            '����� �����������="0" - �������� ����� �� �����������
                            If vntWork = 0 Then
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "0/" + _
                                CStr(vntAddr) + "/" + _
                                CStr(vntReadPortNum) + "  Input"
            '����� �����������="1" - �������� ����� � �����������
                            Else
            '��������� ���� "Reserve" � "���������"
                                gProtocol.strProtocReserve = "1/" + _
                                CStr(vntAddr) + "/" + _
                                CStr(vntReadPortNum) + "  Output"
                            End If
                        End If
                    
                    End If
                
                End If
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
                gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"

'''            '��� ��������� ��������� ��������� ��� ���������
'''            '  �� ������� ������� �������� �� ���������, �����������
'''            '   ����� "����" ��� ���������� �����
'''                If intScriptCode <> 0 And gSluice(intIndex) <> 0 Then
'''            '�������� ��� "Controller'�" ����� - ����� � ��������� �������
'''                    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             '����� ���������� �������� ������� �����
'''                    Do
'''                    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''                End If
            
            '����, ����� ��� �������� ������� - (�������������)
            ElseIf intCalendarCode > 1 Or intTimeCode <> 0 Or _
            intTerminalCode <> 0 Then
            '����� "N_?" - (����� ���)
                frmDemo.lblInform(intIndex).BackColor = vbWhite
                    
            '����� �����������="0"
                If vntWork = 0 Then
            '��������� ���������� �� ������� ����� "N_?"
                    gProtocol.strProtocReserve = "0/" + Trim(gProtocol.strProtocReserve) + "||" + _
                    CStr(vntReadPortNum) + "N"
            '����� �����������="1"
                Else
            '��������� ���������� �� ������� ����� "N_?"
                    gProtocol.strProtocReserve = "1/" + Trim(gProtocol.strProtocReserve) + "||" + _
                    CStr(vntReadPortNum) + "N"
                End If
                    
                GoTo ResetController
            End If
        End If
        
    End If
Protocol:
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus
            '�������� ������ � ���� "������� ���������"
    frmDemo.WriteProtocol
            '����� �� ���������
    Exit Sub
ResetController:
            '�������� ������ � ���� "������� ���������"
    frmDemo.WriteProtocol
            
'''            '����������� ����� "����" ��� ���������� �����
'''    If gSluice(intIndex) <> 0 Then
'''            '�������� ��� "Controller'�" ����� - ����� � ��������� �������
'''        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208)))
'''             '����� ���������� �������� ������� �����
'''        Do
'''        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
'''    End If

            '���� �������� �������
WaitCycle:
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub

            '������ ������� �������
            '   ��� ��������: 0 - ������ ����� ��������� ��������;
            '                 1 - ������ ����� ������������ ���������� ����������;
            '                 2 - "������ ����";
            '                 3 - "Relay" ��� � �������������� �� ������;
            '                 5 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ���������� �������� � ������ ��� �����������;
            '                 6 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ��������� �������� � ������ ��� �����������;
            '                 7 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ���������� �������� � ������ ��� �����������.
            '                 8 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ���������� �������� � ������ ��� �����������;
            '                 9 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ��������� �������� � ������ ��� �����������;
            '                10 - ������ ����� ������������ ���������� ���������� -
            '                     ��� ���������� �������� � ������ ��� �����������.
Private Function StatusCode()
            '������ ������� ����� ��������� ��������
    If Left(gProtocol.strProtocStatus, 2) = "00" Then
        StatusCode = 0
            '������ ������� ����� ������������ ���������� ����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "01" Then
        StatusCode = 1
            '������ ����
    ElseIf Left(gProtocol.strProtocStatus, 2) = "02" Then
        StatusCode = 2
            '������ ������� - "Relay" ��� � �������������� �� ������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "03" Then
        StatusCode = 3
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ���������� �������� � ������ ��� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "05" Then
        StatusCode = 5
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ��������� �������� � ������ ��� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "06" Then
        StatusCode = 6
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ���������� �������� � ������ ��� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "07" Then
        StatusCode = 7
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ���������� ����������� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "08" Then
        StatusCode = 8
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ��������� ����������� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "09" Then
        StatusCode = 9
            '������ ������� ����� ������������ ���������� ���������� -
            '  ��� ���������� ����������� �����������
    ElseIf Left(gProtocol.strProtocStatus, 2) = "10" Then
        StatusCode = 10
    End If
    
End Function

            '������ ��� �������
            '   ��� ��������: 0 - ������ � ����� ����;
            '                 1 - ����������� ������ ������� - ��������� ����;
            '                 2 - ����������� ������ ������� - ����������� ����;
            '                 3 - ����������� ����� �������.
Private Function CalendarCode()
            '������� ������� "������� ������" = 4 (���������)
    gTablePerson.Col = 4
            '������ � ����� ����
    If Left(Trim(gTablePerson.Text), 2) = "00" Then
        CalendarCode = 0
            '����������� ������ �������
    ElseIf Left(Trim(gTablePerson.Text), 2) = "01" Then
            '��������� ����
        If InStr(1, gToday(0), "/") = 0 Then CalendarCode = 1
            '����������� ����
        If InStr(1, gToday(0), "/*") <> 0 Then CalendarCode = 2
    End If
            '����������� ����� �������
    If Left(Trim(gTablePerson.Text), 2) = "02" Or _
    InStr(1, gToday(0), "/^") <> 0 Then CalendarCode = 3
                    
End Function

            '������ ��������������� ��� �������
            '   ��� ��������: 1 - ��������� ����;
            '                 2 - ����������� ����.
Private Function IndividualCalendar()
            '����� ������� �������� "/" � ���� "Reservation"
Dim intPosNum As Integer
            '����� �������� "������� ���������"
Dim intCalendarNum As Integer
            
            '���� �������� ������ ��� ������� ����������
    On Error GoTo CheckError
            
            '������� ������� "������� ������" = 5 (�������������)
    gTablePerson.Col = 5
            '����� ������� ������� �������� ..."/"
    intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            '����� ������� ������� �������� ..."/"
    intPosNum = InStr(intPosNum + 1, Trim(gTablePerson.Text), "/")
            '����� �������� "������� ���������" ����������� - ��������� ����
    If Len(Trim(gTablePerson.Text)) = intPosNum Then
        IndividualCalendar = 1
    Else
            '����� �������� "������� ���������"
        intCalendarNum = Mid(Trim(gTablePerson.Text), intPosNum + 1)
            '���� � �������� ���� ����������� ������� "/" - ��������� ����
        If InStr(1, gToday(intCalendarNum), "/") = 0 Then
            IndividualCalendar = 1
            '����������� ����
        Else
            IndividualCalendar = 2
        End If
    
    Exit Function
            '���� �������� ������ ��� ������� ����������
CheckError:
            '���� ������� - (�������������)
        IndividualCalendar = 2
    End If
    
End Function

            '������ ������� �������
            '   ��� ��������: 0 - ����������� ����� �������;
            '                 1 - ������������� ����� �������.
Private Function TimeCode()
            '������� ������� "������� ������" = 3 (�����)
    gTablePerson.Col = 3
            '����� ������� - (�����������)
    If ((Left(Trim(gTablePerson.Text), 2) < intHour _
    Or Left(Trim(gTablePerson.Text), 2) = intHour _
    And Mid(Trim(gTablePerson.Text), 4, 2) <= intMinute) _
    And (Mid(Trim(gTablePerson.Text), 7, 2) > intHour _
    Or Mid(Trim(gTablePerson.Text), 7, 2) = intHour _
    And Mid(Trim(gTablePerson.Text), 10, 2) >= intMinute)) Or _
((((CInt(Left(Trim(gTablePerson.Text), 2)) * 60 + _
CInt(Mid(Trim(gTablePerson.Text), 4, 2)))) > _
(CInt(Mid(Trim(gTablePerson.Text), 7, 2)) * 60 + _
CInt(Mid(Trim(gTablePerson.Text), 10, 2)))) And _
        (((intHour * 60 + intMinute) >= _
    (CInt(Left(Trim(gTablePerson.Text), 2)) * 60 + _
    CInt(Mid(Trim(gTablePerson.Text), 4, 2))) And _
    (intHour * 60 + intMinute) <= 24 * 60) Or _
        ((intHour * 60 + intMinute) <= _
    (CInt(Mid(Trim(gTablePerson.Text), 7, 2)) * 60 + _
    CInt(Mid(Trim(gTablePerson.Text), 10, 2))) And _
    (intHour * 60 + intMinute) >= 0))) Then

        TimeCode = 0
            
            '����� ������� - (�������������)
    Else
        TimeCode = 1
    End If
                    
End Function

            '������ ��������������� ������� �������
            '   ��� ��������: 0 - ����������� ����� �������;
            '                 1 - ������������� ����� �������.
Private Function IndividualTime(ByVal vntAddr As Variant, ByVal vntReadPortNum _
                                As Variant)
            '����� ������� �������� "/"
Dim intPosNum As Integer
            '����� �������� "������� �������"
Dim intTimeNum As Integer
            '����� �������� �������������� "������� ����������"
Dim intTerminalNum As Integer
            '����� �������� �������������� "������� ���������"
Dim intCalendarNum As Integer
            '������� �������
Dim intCount As Integer
            '������� �������
Dim intCount_1 As Integer
            
            '���� �������� ������ ��� ������� ����������
    On Error GoTo CheckError
    
            '������� ������� "������� ������" = 5 (�������������)
    gTablePerson.Col = 5
            '����� �������� "������� �������" ����������� - ����������� ����� �������
    If Left(Trim(gTablePerson.Text), 1) = "/" Then
        IndividualTime = 0
            '��������� ����� �������� "������� �������"
    Else
            '����� ������� ������� �������� ..."/"
        intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            '����� �������� "������� �������"
        intTimeNum = Left(Trim(gTablePerson.Text), intPosNum - 1)
            '����� ������� - (�������������)
        IndividualTime = 1
            '�� ���� "��������" �������� ������� ������ ������� ����������
            '  � ������� �������������� (��������� ��������) ����������
            '  � ����������
        For intCount = 1 To CInt(gInterval(intTimeNum, 0)) - 1 Step 1
            '����� ������� - (�����������)
            If (Left(gInterval(intTimeNum, intCount), 2) < intHour _
            Or Left(gInterval(intTimeNum, intCount), 2) = intHour _
            And Mid(gInterval(intTimeNum, intCount), 4, 2) <= intMinute) _
            And (Mid(gInterval(intTimeNum, intCount), 7, 2) > intHour _
            Or Mid(gInterval(intTimeNum, intCount), 7, 2) = intHour _
            And Mid(gInterval(intTimeNum, intCount), 10, 2) >= intMinute) Then
            '��� �������������� (��������� ��������)���������� ��� ����������
                If Left(gTerCal(intTimeNum, intCount), 8) = "Interval" Then
                    IndividualTime = 0
                    Exit For
            '���� �������������� (��������� ��������)��������� ��� ���������
                Else
            '����� ������� �������� "/" � ������� �������������� ����������
            '  � ����������
                    intPosNum = InStr(1, Trim(gTerCal(intTimeNum, _
                    intCount)), "/")
            '����� �������� �������������� "������� ����������" �����������
            '  - ����������� ����� �������
                    If intPosNum = 1 Then
                        IndividualTime = 0
            '��������� ����� �������� �������������� "������� ����������"
                    Else
            '����� �������� �������������� "������� ����������"
                        If intPosNum = 0 Then
                            intTerminalNum = Trim(gTerCal(intTimeNum, _
                            intCount))
                        Else
                            intTerminalNum = Left(Trim(gTerCal(intTimeNum, _
                            intCount)), intPosNum - 1)
                        End If
            '�������������� �������� ������� �����������
            '  - ��������� ����� ���������
                        IndividualTime = 1
            '�� ���� "��������" �������� ������� ������ ������� ����������
                        For intCount_1 = 1 To _
                        CInt(gAddrPort(intTerminalNum, 0)) - 1 Step 1
            '�������������� �������� ������� �����������
            '  - ��������� ����� ��������
                            If vntAddr = (CByte(Left(gAddrPort(intTerminalNum, _
                            intCount_1), 1) * 16) Or _
                            CByte(Mid(gAddrPort(intTerminalNum, intCount_1), 2, 1))) _
                            And (vntReadPortNum = _
                            CByte(Mid(gAddrPort(intTerminalNum, intCount_1), 3, 1))) Then
                                IndividualTime = 0
                                Exit For
                            End If
                        Next
                    End If
                    
            '� ������� �������������� ���������� � ���������� ������� ������� "/"
            '  - ���������� �������� �������������� "������� ���������"
                    If intPosNum <> 0 And IndividualTime = 0 Then
            '��������� ����� �������� �������������� "������� ���������"
                        intCalendarNum = Mid(Trim(gTerCal(intTimeNum, intCount)), _
                        intPosNum + 1)
            '���� � �������� ���� ������� ������� "/" - ����������� ����
                        If InStr(1, gToday(intCalendarNum), "/") <> 0 Then
                            IndividualTime = 1
                        End If
                    End If
            '��� �������� �������� - ������ ��������
                    If IndividualTime = 0 Then Exit For
                End If
            End If
        Next
    End If
            
    Exit Function
            '���� �������� ������ ��� ������� ����������
CheckError:
            '����� ������� - (�������������)
        IndividualTime = 1
                    
End Function

            '������ ��������������� ��������� �������
            '   ��� ��������: 0 - ������ ����� ������ �������� ��������;
            '                 1 - ������ ����� ������ �������� ��������.
Private Function IndividualTerminal(ByVal vntAddr As Variant, ByVal vntReadPortNum _
                                    As Variant)
            '����� ������� �������� "/" � ���� "Reservation"
Dim intPosNum As Integer
            '����� �������� "������� ����������"
Dim intTerminalNum As Integer
            '������� �������
Dim intCount As Integer
            
            '���� �������� ������ ��� ������� ����������
    On Error GoTo CheckError
            
            '������� ������� "������� ������" = 5 (�������������)
    gTablePerson.Col = 5
            '����� ������� ������� �������� ..."/"
    intPosNum = InStr(1, Trim(gTablePerson.Text), "/")
            '����� �������� "������� ����������" ����������� - ������ ��������
    If Mid(Trim(gTablePerson.Text), intPosNum + 1, 1) = "/" Then
            '��������� ����� �������� "������� ����������"
        IndividualTerminal = 0
    Else
            '����� �������� "������� ����������"
        intTerminalNum = Mid(Trim(gTablePerson.Text), intPosNum + 1, _
        InStr(intPosNum + 1, Trim(gTablePerson.Text), "/") - intPosNum - 1)
            '�������� ������� - (�����������)
        IndividualTerminal = 1
            '�� ���� "��������" �������� ������� ������ ������� ����������
        For intCount = 1 To CInt(gAddrPort(intTerminalNum, 0)) - 1 Step 1
            '�������� ������� - (�����������)
            If vntAddr = (CByte(Left(gAddrPort(intTerminalNum, intCount), 1) * 16) _
            Or CByte(Mid(gAddrPort(intTerminalNum, intCount), 2, 1))) _
            And (vntReadPortNum = CByte(Mid(gAddrPort(intTerminalNum, intCount), 3, 1))) Then
                IndividualTerminal = 0
                Exit For
            End If
        Next
    End If
    
    Exit Function
            '���� �������� ������ ��� ������� ����������
CheckError:
            '�������� ������� - (�������������)
        IndividualTerminal = 1
                    
End Function

            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
            '   � ����� �� ������� ����� ���������
            '   ��� ��������: 0 - �������� ��� ��������� ��������� ���������;
            '                 1 - ��� ��������� �������� ��� �����������;
            '                 2 - �������� ��� ��������� ��������� ���������.
Private Function ScriptTermClose(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������ ������ ����������
    gTermContr = 0
            '��������� ������� ����� �����
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            '����� �������� ��������
            ' � ������� "������� ����������",
    intRequest = (vntWork - 2) * 15 + vntAddr
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            '������������ �������� ��� ���������
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(32) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            '���� �������� ���� ��������� "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ���� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(63)
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            If vntWork = vntWork1 Then
                ScriptTermClose = 0
            '����� �� ���������
                Exit Do
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            ElseIf vntWork = vntWork2 Then
                ScriptTermClose = 2
            '����� �� ���������
                Exit Do
            '�������� �������� ��� ��������� ��� ����� "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ��������� ��������� ��� ���������
            '  ��������� �� "Controller'a" �� �������
                ScriptTermClose = 1
            '����� �� ���������
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ��������� ��������� ��� ���������
            '  ��������� �� "Controller'a" �� �������
            ScriptTermClose = 1
            '����� �� ���������
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            '��������� ������� TimeOut ��� �������� ��� ���������
    If ScriptTermClose = 1 Then
            '���������������� ������� - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            '��������� ������
        gProtocol.strProtocPersonCode = "Command=E/16"
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "COMMAND TimeOut"
            
            '��� ������������� ����� ���������
            '   ��������� "Controller" �� �������
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            '�������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ��������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            '���������� ������ ����������
    gTermContr = 1
                    
End Function

            '�������� ��������� �� "Controller'a" �� ������� ������� ��������
            '   ��� ��������: 0 - ��������� ���������;
            '                 1 - ��������� �����������.
Private Function ScriptOpen(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������ ������ ����������
    gTermContr = 0
            '��������� ������� ����� �����
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            '����� �������� ��������
            ' � ������� "������� ����������",
    intRequest = (vntWork - 2) * 15 + vntAddr
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "F"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            '������������ �������� ��� ���������
    vntWork1 = CByte(240) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� �������� ��������� �� "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            '���� ������ "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            '��������� �� "Controller'a" � ��������������� �������
            If CByte(Asc(vntWork)) = CByte(vntWork1) Then
                ScriptOpen = 0
            'B���� ��������� ������ ����������� ��������� ���������
                PictureTerminalOpen intIndex
            '����� �� ���������
                Exit Do
            '�������� �� ��������� ��� �������� ����� "Controller'a"
            Else
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ���������
                ScriptOpen = 1
            '����� �� ���������
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��������� ������� TimeOut
            ScriptOpen = 1
            '����� �� ���������
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            '��������� ������� TimeOut ��� ��� ���������
    If ScriptOpen = 1 Then
            '��������� ������� ����� �����
        vntWork = frmDemo.prtPortC(intIndex).CommPort
            '���������������� ������� - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            '��������� ������
        gProtocol.strProtocPersonCode = "Command=1/16"
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "COMMAND TimeOut"
            
            '��� ������������� ����� ���������
            '   ��������� "Controller" �� �������
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            '�������� ������� �������� "Controller'a",
            '  �� �������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            '���������� ������ ����������
    gTermContr = 1
                    
End Function
            
            '��������� ������������������ �������� ���������
Private Sub InitialCloseTerminal(intIndex As Integer, intRequest As Integer)
            
            '��� �������� ��� �������� ���� ��������� ��� ���������
Dim intScriptCode As Integer
            '����� �����������
Dim vntAddr As Variant
    
            ' "Controller" ��������� �������� �� �������
    If Mid(gAddrPort(0, intRequest), 1, 2) = "00" Then GoTo WaitCycle

        '������������ ����� "Controller'a", �������� ������������� ���������
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            '�������� �������� ����� �����
    frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
    Do
    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
    intScriptCode = ScriptTermOpen(intIndex, vntAddr)
            '���� �������� �������
WaitCycle:

End Sub

            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
            '   � ����� �� ������� ����� ���������
            '   ��� ��������: 0 - �������� ��� ��������� ��������� ���������;
            '                 1 - ��� ��������� �������� ��� �����������;
            '                 2 - �������� ��� ��������� ��������� ���������.
Private Function ScriptTermOpen(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������ ������ ����������
    gTermContr = 0
            '��������� ������� ����� �����
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            '����� �������� ��������
            ' � ������� "������� ����������",
    intRequest = (vntWork - 2) * 15 + vntAddr
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            '������������ �������� ��� ���������
    vntWork1 = CByte(32) Or CByte(vntAddr)
    vntWork2 = CByte(16) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            '���� ������ "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ���� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(63)
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            If vntWork = vntWork1 Then
                ScriptTermOpen = 0
            '����� �� ���������
                Exit Do
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            ElseIf vntWork = vntWork2 Then
                ScriptTermOpen = 2
            '����� �� ���������
                Exit Do
            '�������� �������� ��� ��������� ��� ����� "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ��������� ��������� ��� ���������
            '  ��������� �� "Controller'a" �� �������
                ScriptTermOpen = 1
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��������� ������� TimeOut
            ScriptTermOpen = 1
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            '��������� ������� TimeOut ��� �������� ��� ���������
    If ScriptTermOpen = 1 Then
            '���������������� ������� - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            '��������� ������
        gProtocol.strProtocPersonCode = "Command=E/16"
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "COMMAND TimeOut"
''            '�������� ������ � ���� "������� ���������"
''        frmDemo.WriteProtocol
            
            '��� ������������� ����� ���������
            '   ��������� "Controller" �� �������
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            '�������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ��������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            '���������� ������ ����������
    gTermContr = 1
                    
End Function

            '�������� ��������� �� "Controller'a" �� ������� ������� ��������
            '   ��� ��������: 0 - ��������� ���������;
            '                 1 - ��������� �����������.
Private Function ScriptClose(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
            '����� �������� � ������� "������� ����������"
Dim intRequest As Integer
            '������ ������ ����������
    gTermContr = 0
            '��������� ������� ����� �����
    vntWork = frmDemo.prtPortC(intIndex).CommPort
            '����� �������� ��������
            ' � ������� "������� ����������",
    intRequest = (vntWork - 2) * 15 + vntAddr
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "F"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
            '������������ ��������� ��� ���������
    vntWork1 = CByte(240) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� �������� ��������� �� "Controller'a"
    frmDemo.tmrTimeOut(intIndex).Tag = 0
    frmDemo.tmrTimeOut(intIndex).Enabled = True
            '���� ������ "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            '��������� �� "Controller'a" � ��������������� �������
            If CByte(Asc(vntWork)) = CByte(vntWork1) Then
                ScriptClose = 0
            '����� �� ���������
                Exit Do
            '�������� �� ��������� ��� �������� ����� "Controller'a"
            Else
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��������� �� "Controller'a" �� ��������
                ScriptClose = 1
            '����� �� ���������
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrTimeOut(intIndex).Tag <> 0 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
            frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
            Do
            Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��������� �� "Controller'a" �� ��������
            ScriptClose = 1
            '����� �� ���������
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrTimeOut(intIndex).Enabled = False
            '��������� ������� TimeOut ��� ��� ���������
    If ScriptClose = 1 Then
            '��������� ������� ����� �����
        vntWork = frmDemo.prtPortC(intIndex).CommPort
            '���������������� ������� - "TimeOut"
        gProtocol.strProtocName = "Addr=" + Left(gAddrPort(0, intRequest), 2) + _
        " / " + "Port=" + Mid(gAddrPort(0, intRequest), 3, 1)
            '��������� ������
        gProtocol.strProtocPersonCode = "Command=A/16"
            '������
        gProtocol.strProtocStatus = ""
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "COMMAND TimeOut"
''            '�������� ������ � ���� "������� ���������"
''        frmDemo.WriteProtocol
            
            '��� ������������� ����� ���������
            '   ��������� "Controller" �� �������
        If intTerminalLogOFF <> 0 Then gAddrPort(0, intRequest) = "00" _
        + Mid(gAddrPort(0, intRequest), 3, 1) + "0"

    End If
            '������������ ������� �������� "Controller'a",
            '  �� �������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "A"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            '���������� ������ ����������
    gTermContr = 1
                    
End Function
            
            '������������������ �������� �������� ���������
Private Sub WaitOpenTerminal(intIndex As Integer, intRequest As Integer)
            '��� �������� ��� �������� ���� ���������
Dim intScriptCode As Integer
            '����� �����������
Dim vntAddr As Variant
    
            '������ ������ ����������
    gTermContr = 0
            
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
        
        '������������ ����� "Controller'a", �������� ������������� ���������
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrButton(intIndex).Tag = 0
    frmDemo.tmrButton(intIndex).Enabled = True
            '���� �������� ���� �������� (��������� ���������)
    Do While DoEvents()
            
            '�������� �������� ����� �����
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� �������� (��������� ���������)
        intScriptCode = TerminalOpen(intIndex, vntAddr)
            
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        If frmDemo.tmrButton(intIndex).Tag <> 0 Then
            '����� �� ���������
            Exit Do
            'K�� �������� ��������� ��������� ��������
        ElseIf intScriptCode = 0 Then
            '����� �� ���������
            Exit Do
            'K�� �������� ��������� ��������� ��� �� ��������
            '  - ���������� ����� ���� �������� ��������� ���������
        ElseIf intScriptCode <> 0 Then
            
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrButton(intIndex).Enabled = False
            
            '�������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ��������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
            '����� ��������� ������ ����������� ��������� ���������
    PictureTerminalOpen intIndex
            
            '���������� ������ ����������
    gTermContr = 1
    
End Sub
            
            '������������������ �������� �������� ���������
Private Sub WaitCloseTerminal(intIndex As Integer, intRequest As Integer)
            '��� �������� ��� �������� ���� ���������
Dim intScriptCode As Integer
            '����� �����������
Dim vntAddr As Variant
    
            '������ ������ ����������
    gTermContr = 0
            
            '���������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "E"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag + 1
        
        '������������ ����� "Controller'a", �������� ������������� ���������
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrButton(intIndex).Tag = 0
    frmDemo.tmrButton(intIndex).Enabled = True
            '���� �������� ���� �������� (��������� ���������)
    Do While DoEvents()
            
            '�������� �������� ����� �����
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� �������� (��������� ���������)
        intScriptCode = TerminalClose(intIndex, vntAddr)
            
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        If frmDemo.tmrButton(intIndex).Tag <> 0 Then
            '����� �� ���������
            Exit Do
            'K�� �������� ��������� ��������� ��������
        ElseIf intScriptCode = 0 Then
            '����� �� ���������
            Exit Do
            'K�� �������� ��������� ��������� ��� �� ��������
            '  - ���������� ����� ���� �������� ��������� ���������
        ElseIf intScriptCode <> 0 Then
                
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrButton(intIndex).Enabled = False
            
            '�������� ������� �������� "Controller'a",
            '  �� �������� ��������� ��� ��������� ��������� ���������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "0"
            '��������� ������� "Controller'��" ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
            '����� ��������� ������ ����������� ��������� ���������
    PictureTerminalClose intIndex
            
            '���������� ������ ����������
    gTermContr = 1
    
End Sub
            
            ' "������������������ �������� ���������"
            '    �� ����������� "������"
Private Sub ButtonOpenTerminal(intIndex As Integer, intRequest As Integer)
            '����� "Controller'a"
Dim vntAddr As Variant
            '��� �������� ��� �������� ��������� ��� ���� ���������
Dim intScriptCode As Integer

            '���������� ����� "Controller'a"
    vntAddr = CByte(Left(gAddrPort(0, intRequest), 1) * 16) _
    Or CByte(Mid(gAddrPort(0, intRequest), 2, 1))
            
            '�������� �������� ����� �����
    frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ����� ��������� � ����������� �������
    frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(224) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ����� ���������
    Do
    Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
    intScriptCode = ScriptTermClose(intIndex, vntAddr)
            '��� ��������� ��������� ��������� ��������
    If intScriptCode = 0 Then
            '�������� �������� ����� �����
        frmDemo.prtPortC(intIndex).InBufferCount = 0
            '������� "Controller'y" ������� - ������� �������� � ����������� �������
        frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(16) Or CByte(vntAddr)))
             '����� ���������� �������� ������� ������� ��������
        Do
        Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '�������� ��������� �� "Controller'a"
        intScriptCode = ScriptOpen(intIndex, vntAddr)
    End If
            '���������� ������� �������� "Controller'a", � ��������
            '   ��������� ��������� ��������� �� �����������
    gAddrPort(0, intRequest) = Left(gAddrPort(0, intRequest), 3) + "#"
            '��������� ������� "Controller'��", ������� ������������� ����������
    frmDemo.prtPortC(intIndex).Tag = frmDemo.prtPortC(intIndex).Tag - 1
            
End Sub
            
            '��������� ������ ����������� ��������� ���������
Private Sub PictureTerminalOpen(intIndex As Integer)
        frmDemo.imgViewClose(intIndex).Visible = False
        frmDemo.imgViewOpen(intIndex).Visible = True
            '��������� ������� ����������� "������" (������������
            '  ��� ������ ����������� ��������� ���������)
        frmDemo.tmrButton(intIndex).Tag = 0
        frmDemo.tmrButton(intIndex).Enabled = True

End Sub
            
            '��������� ������ ����������� ��������� ���������
Private Sub PictureTerminalClose(intIndex As Integer)
        frmDemo.imgViewOpen(intIndex).Visible = False
        frmDemo.imgViewClose(intIndex).Visible = True

End Sub

            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
            '   � ����� �� ������� ����� ���������
            '   ��� ��������: 0 - �������� ��� ��������� ��������� ���������;
            '                 1 - ��� ��������� �������� ��� �����������.
Private Function TerminalOpen(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            
            '������������ �������� ��� ���������
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(0) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrRelay.Interval = frmDemo.tmrTimeOut(intIndex).Interval
    frmDemo.tmrRelay.Tag = 0
    frmDemo.tmrRelay.Enabled = True
            '���� �������� ���� ��������� "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ���� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(31)
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            If vntWork = vntWork2 Then
                TerminalOpen = 0
            '����� �� ���������
                Exit Do
            '��� ��������� ��������� ��������� �� "Controller'a" � ���������������
            '  ������� - ���������� ����� ���� ��������� ��������� ���������
            ElseIf vntWork = vntWork1 Then
            
                TerminalOpen = 2 '����� ����� �� ������������
                
            '�������� �������� ��� ��������� ��� ����� "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ��������� ��������� ��� ���������
            '  ��������� �� "Controller'a" �� �������
                TerminalOpen = 1
            '����� �� ���������
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrRelay.Tag <> 0 Then
            '��� ��������� ��������� ��������� �� "Controller'a" �� �������
            TerminalOpen = 1
            '����� �� ���������
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrRelay.Enabled = False
                    
End Function

            '�������� ���� ��������� ��������� ��������� �� "Controller'a"
            '   � ����� �� ������� ����� ���������
            '   ��� ��������: 0 - �������� ��� ��������� ��������� ���������;
            '                 1 - ��� ��������� �������� ��� �����������.
Private Function TerminalClose(intIndex As Integer, ByVal vntAddr As Variant)
            '������� ����
Dim vntWork As Variant
Dim vntWork1 As Variant
Dim vntWork2 As Variant
            
            '������������ �������� ��� ���������
    vntWork1 = CByte(16) Or CByte(vntAddr)
    vntWork2 = CByte(0) Or CByte(vntAddr)

            '�������� �������� ������� TimeOut ��� ��������
            '  ���� ��������� �� "Controller'a"
    frmDemo.tmrRelay.Interval = frmDemo.tmrTimeOut(intIndex).Interval
    frmDemo.tmrRelay.Tag = 0
    frmDemo.tmrRelay.Enabled = True
            '���� �������� ���� ��������� "Controller'a"
    Do While DoEvents()
            '����� �� "Controller'a" ��������� ���� ��������� �� ������� TimeOut
        If frmDemo.prtPortC(intIndex).InBufferCount >= 1 Then
            '���������� ������ � �������� ����� ��� ���������� ���������
            vntWork = frmDemo.prtPortC(intIndex).Input
            vntWork = CByte(Asc(vntWork)) And CByte(31)
            '��� ��������� ��������� ���������
            '  �� "Controller'a" � ��������������� �������
            If vntWork = vntWork1 Then
                TerminalClose = 0
            '����� �� ���������
                Exit Do
            '��� ��������� ��������� ��������� �� "Controller'a" � ���������������
            '  ������� - ���������� ����� ���� ��������� ��������� ���������
            ElseIf vntWork = vntWork2 Then
            
                TerminalClose = 2 '����� ����� �� ������������
                
            '�������� �������� ��� ��������� ��� ����� "Controller'a"
            ElseIf vntWork <> vntWork1 Then
            '�������� ��������������� "Controller" - ����� � ����������� �������
                frmDemo.prtPortC(intIndex).Output = Chr(CLng(CByte(208) Or CByte(vntAddr)))
             '����� ���������� �������� ������� �����
                Do
                Loop Until frmDemo.prtPortC(intIndex).OutBufferCount = 0
            '��� ��������� ��������� ��� ���������
            '  ��������� �� "Controller'a" �� �������
                TerminalClose = 1
            '����� �� ���������
                Exit Do
            End If
            '��������� ������� TimeOut ��� �������� ���� ��������� �� "Controller'a"
        ElseIf frmDemo.tmrRelay.Tag <> 0 Then
            '��� ��������� ��������� ��������� �� "Controller'a" �� �������
            TerminalClose = 1
            '����� �� ���������
            Exit Do
        End If
    Loop
            '��������� �������� ������� TimeOut
    frmDemo.tmrRelay.Enabled = False
            
End Function

            '����������� "������� ������" �� "Host Computer'a"
Private Sub TablePersonCopy()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '������ ��� �����-����� (� ��������� "����" � ����)
Dim strCopyFileName As String
            '������ ��� �����-����� "Host Computer'a" (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            
            '������� ������� "��������� �������" = 0 (���)
    frmTableSystem.grdTableSystem.Col = 0
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = intRowNum
            '������ �������� ����������� "������� ������" �� "Host Computer'a"
        If Trim(frmTableSystem.grdTableSystem.Text) = _
        "CopyTablePerson" Then
            '������� ������� "��������� �������" = 1
            frmTableSystem.grdTableSystem.Col = 1
            '��������� ����������� "������� ������" �� "Host Computer'a"
            If Mid(Trim(frmTableSystem.grdTableSystem.Text), 2, 2) = ":\" Then
            '������ ��� �����-����� "Host Computer'a" (� ��������� "����" � ���)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
                Exit For
            '�� ��������� ����������� "������� ������" �� "Host Computer'a"
            Else
                Exit Sub
            End If
        End If
    Next
            '����������� ������ "��������� �������", �������� �������
            '  ������������� ����������� "������� ������" �� "Host Computer'a"
    If intRowNum = frmTableSystem.grdTableSystem.Rows Then Exit Sub
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strCopyFileName = App.Path
    If Right(strCopyFileName, 1) <> "\" Then
            '������ ��� ����� ��� �����-����� (� ��������� "����" � ���)
        strCopyFileName = strCopyFileName + "\"
    End If
    
            '�������� ������������� �����-����� "Host Computer'a"
    On Error GoTo UnExist
            '�����-���� ������� - ����������
    If (FSO.FolderExists(strPathFolderName)) Then
            '������ ��� ����������� ����� "Host Computer'a"
            '  (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\" + "TablePerson.dat"

        If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ������ � "Host Computer"
            FSO.CopyFile strPathFileName, strCopyFileName
        End If
    End If

UnExist:
    On Error GoTo 0

End Sub


