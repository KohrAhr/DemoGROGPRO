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
            '������� ����� �����
Dim intFileNum As Integer
            '������ "������� ���������"
Dim gSystem As SystemInfo



            '��������� "������" ���� �� �����e "cboPreprocessors"
Private Sub cboPreprocessors_Click()
            
            '������� ��������� ������� ������
    cboFileName.ListIndex = _
    cboPreprocessors.ListIndex

End Sub

            '��������� ������� "RestartWorkStation"
Private Sub cmdRestartWorkStation_Click()

            '������ �������� ��������� �� "Preprocessor"
Dim strMessage As String

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If

            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������ �������� ��������� �� "Preprocessor"
    strMessage = "StartApp"
            '������� ��� �������������: "Whole"
    If Trim(cboFileName.Text) = "Whole" Then
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������ ���� ������������
    Else
            '������������ ����������� ���������
        qMsgOutput.Body = strMessage
            '���������� ���� � ������� ������������ ���������
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        Trim(cboPreprocessors.Text) + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
        qMsgOutput.Send qQueueOutput
            '������� ������� ���������
        qQueueOutput.Close
    End If
            
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '��������� ������� "StopWorkStation"
Private Sub cmdStopWorkStation_Click()

            '������ �������� ��������� �� "Preprocessor"
Dim strMessage As String

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If

            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������ �������� ��������� �� "Preprocessor"
    strMessage = "StopApp"
            '������� ��� �������������: "Whole"
    If Trim(cboFileName.Text) = "Whole" Then
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '������ ���� ������������
    Else
            '������������ ����������� ���������
        qMsgOutput.Body = strMessage
            '���������� ���� � ������� ������������ ���������
        qInfoOutput.FormatName = "DIRECT=OS:" + _
        Trim(cboPreprocessors.Text) + "\Private$\GeneralQueue"
            '������� ������� ��������� � ����������� (��� ��������
            '  ���������, ������ � ������� �������� ����)
        Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            '�������� ���������
        qMsgOutput.Send qQueueOutput
            '������� ������� ���������
        qQueueOutput.Close
    End If
            
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '��������� ������� "Cancel"
Private Sub cmdCancel_Click()

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '��������� ������� "DownLoad Protocol"
Private Sub cmdProtocol_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '��� �������� ��� ���������� "��������� �������"
Dim intSaveTableSystem As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '������ ��� ���������� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ������ ���������� �� ������������� "������� ���������"
Dim lngRecordLen As Long
            '���������� ����� � ���������� �������
            '   "TableProtocol" �������������
Dim intRowQuanP As Integer
            '������� ����� ������ ���������� �������
            '   "TableProtocol" �������������
Dim intRowNumP As Integer

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If

            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '������ ��� �����-����� �������������
            '   (� ��������� "����" � ���) ��� "Whole"
    strPathFolderName = cboFileName.Text
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������� ��� �������������: "Whole"
    If strPathFolderName = "Whole" Then
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 0 (������)
                frmTableSystem.grdTableSystem.Col = 0
            '������� ������� �� ������ ���������������� ���� "cboFileName"
            '   - �����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            
            '�������� ������������� �����-����� �������������
                On Error GoTo UnAccessable
                If (FSO.FolderExists(strPathFolderName)) Then
            '�����-���� ������� - ����������
                    On Error GoTo CopyingMistake
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
                    If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ����� �� ������������� ����� �����
            '   ������� "TableProtocol" �� �����-����� ������������� � �����
            '   ����� ������� "TableProtocol" ��� "Host Computer'a"
                
            '��������� ����� ������ (������)
            '   "������� ���������" �������������
                        lngRecordLen = Len(gProtocol)
            '���������� ����� � "������� ���������" �������������
                        intRowQuanP = FileLen(strPathFileName) / lngRecordLen
            '������� ������� "��������� �������" = 4 (���������)
                        frmTableSystem.grdTableSystem.Col = 4
            '�� ���� ������ "������� ���������" ������������� ����� ��
            '   ������������ � "Host Computer" ��� "������� ���������"
            '   ������������� ������ �������������� ������ (��������,
            '   ��� �������������) - ���������� � ������ ������
                        If Trim(frmTableSystem.grdTableSystem.Text) = "" Or _
                        intRowQuanP < _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = 1
            '��� ������ "������� ���������" ������������� ����� ����
            '   ����������� � "Host Computer" ��� "������� ���������"
            '   ��� �� ������������ �������������� � ����������� -
            '   � ���������� �������������
                        ElseIf intRowQuanP = 0 Or intRowQuanP = _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            GoTo EndCycle
            '�� ��� ������ "������� ���������" ������������� ����� ����
            '   ����������� � "Host Computer" - ���������� ���������� ������
                        ElseIf intRowQuanP > _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = Trim(frmTableSystem.grdTableSystem.Text) _
                            + 1
                        End If
            
            '�������� ��������� ����� �����
            '   "������� ���������" �������������
                        intFileNum = FreeFile
            '������� ���� "TableProtocol" ������������� ���
            '   ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '������������ �������������� ������
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
            '���������������� ������� - "�������� �������� �������������"
                        gProtocol.strProtocName = strPathFolderName
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                        gProtocol.strProtocReserve = "DownLoad Protocol"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                        
            '���� �� ���� ������� "������� ���������" �������������
                        For intRowNumP = intRowNumP To intRowQuanP Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                            Get intFileNum, intRowNumP, gProtocol
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� "������� ���������" �������������
                        Close intFileNum
            '��������� ����� ��������� ������ "������� ���������"
            '   �������������, ������������� � "Host Computer"
                        frmTableSystem.grdTableSystem.Text = intRowQuanP
            '���������� ������� ��������� � "��������� �������" ���������
                        gChangesTableSystem = True
            '������������ �������������� ������
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                
                    Else
            '���� ����������� - ����� �� ��������� � ����������
                        GoTo CopyingMistake
                    End If
                
                Else
            '�����-���� ����������� - ����� �� ��������� � ����������
                    GoTo UnAccessable
                End If
                
                
                GoTo EndCycle
UnAccessable:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndCycle
CopyingMistake:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
            End If
EndCycle:
        Next
            '������ ���� ������������
    Else
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 0 (������)
            frmTableSystem.grdTableSystem.Col = 0
            '��������� ������������
            If Trim(frmTableSystem.grdTableSystem.Text) = Trim(cboFileName.Text) Then
            '������� ������� �� ������ ���������������� ���� "cboFileName"
            '   - �����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(cboFileName.Text)
            
            '�������� ������������� �����-����� �������������
                On Error GoTo UnExist
                If (FSO.FolderExists(strPathFolderName)) Then
            '�����-���� ������� - ����������
                    On Error GoTo CopyingError
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
                    If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ����� �� ������������� ����� �����
            '   ������� "TableProtocol" �� �����-����� ������������� � �����
            '   ����� ������� "TableProtocol" ��� "Host Computer'a"
                
            '��������� ����� ������ (������)
            '   "������� ���������" �������������
                        lngRecordLen = Len(gProtocol)
            '���������� ����� � "������� ���������" �������������
                        intRowQuanP = FileLen(strPathFileName) / lngRecordLen
            '������� ������� "��������� �������" = 4 (���������)
                        frmTableSystem.grdTableSystem.Col = 4
            '�� ���� ������ "������� ���������" ������������� ����� ��
            '   ������������ � "Host Computer" ��� "������� ���������"
            '   ������������� ������ �������������� ������ (��������,
            '   ��� �������������) - ���������� � ������ ������
                        If Trim(frmTableSystem.grdTableSystem.Text) = "" Or _
                        intRowQuanP < _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = 1
            '��� ������ "������� ���������" ������������� ����� ����
            '   ����������� � "Host Computer" ��� "������� ���������"
            '   ��� �� ������������ �������������� � ����������� -
            '   � ���������� �������������
                        ElseIf intRowQuanP = 0 Or intRowQuanP = _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            GoTo EndProcedure
            '�� ��� ������ "������� ���������" ������������� ����� ����
            '   ����������� � "Host Computer" - ���������� ���������� ������
                        ElseIf intRowQuanP > _
                        Trim(frmTableSystem.grdTableSystem.Text) Then
                            intRowNumP = Trim(frmTableSystem.grdTableSystem.Text) _
                            + 1
                        End If
            '�������� ��������� ����� �����
            '   "������� ���������" �������������
                        intFileNum = FreeFile
            '������� ���� "TableProtocol" ������������� ���
            '   ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '������������ �������������� ������
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
            '���������������� ������� - "�������� �������� �������������"
                        gProtocol.strProtocName = strPathFolderName
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                        gProtocol.strProtocReserve = "DownLoad Protocol"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                        
            '���� �� ���� ������� "������� ���������" �������������
                        For intRowNumP = intRowNumP To intRowQuanP Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                            Get intFileNum, intRowNumP, gProtocol
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� "������� ���������" �������������
                        Close intFileNum
            '��������� ����� ��������� ������ "������� ���������"
            '   �������������, ������������� � "Host Computer"
                        frmTableSystem.grdTableSystem.Text = intRowQuanP
            '���������� ������� ��������� � "��������� �������" ���������
                        gChangesTableSystem = True
            '������������ �������������� ������
                        gProtocol.strProtocName = "=========="
                        gProtocol.strProtocPersonCode = "=========="
                        gProtocol.strProtocStatus = "=========="
                        gProtocol.strProtocTime = "=========="
                        gProtocol.strProtocDate = "=========="
                        gProtocol.strProtocReserve = "=========="
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                
                    Else
            '���� ����������� - ����� �� ��������� � ����������
                        GoTo CopyingError
                    End If
                
                Else
            '�����-���� ����������� - ����� �� ��������� � ����������
                    GoTo UnExist
                End If
                
                GoTo EndProcedure
UnExist:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndProcedure
CopyingError:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
                GoTo EndProcedure
            End If
        Next
EndProcedure:
    End If
            
            '���������� ������� ��������� � "��������� �������"
            '   ��������� - ��������� ������� � ������������ �����
    If gChangesTableSystem = True Then
            '��������� ���������� '��������� �������"
        intSaveTableSystem = frmTableSystem.SaveTableSystem()
    End If
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '��������� ������� "DownLoad Archives"
Private Sub cmdArchives_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '������ ��� �����-����� (� ��������� "����" � ����)
Dim strHostFileName As String
            '������ ��� ���������� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ��� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � "Host Computer"
Dim intDayArchiveCopy As Integer

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If

            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '������ ��� �����-����� �������������
            '   (� ��������� "����" � ���) ��� "Whole"
    strPathFolderName = cboFileName.Text
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strHostFileName = App.Path
    If Right(strHostFileName, 1) <> "\" Then
            '������ ��� ����� "Host Computera" ��� �����-�����
            '  �������������(� ��������� "����" � ���)
        strHostFileName = strHostFileName + "\"
    End If
    
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������� ��� �������������: "Whole"
    If strPathFolderName = "Whole" Then
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 0 (������)
                frmTableSystem.grdTableSystem.Col = 0
            '������� ������� �� ������ ���������������� ���� "cboFileName"
            '   - �����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            
            '�������� ������������� �����-����� �������������
                On Error GoTo UnAccessable
            '�����-���� ������� - ����������
                If (FSO.FolderExists(strPathFolderName)) Then
            '���������������� ������� - "�������� ������ �������������"
                    gProtocol.strProtocName = strPathFolderName
            '��������� ������
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                    gProtocol.strProtocStatus = "04 - Manager"
            '�����
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "DownLoad Archives"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                    On Error GoTo CopyingMistake
                    
            '��������� "���������" �� ������� ����
                    frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                    For intDayArchiveCopy = 1 To gDayNum Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                        frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ �������������
            '  (� ��������� "����" � ����)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            '������������ ����������
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

                        If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ������ � "Host Computer"
                            FSO.CopyFile strPathFileName, _
                            strHostFileName
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        
            '���������������� ������� - "����������� ������ �������������"
                            gProtocol.strProtocName = "Copy Archive"
            '��������� ������
                            gProtocol.strProtocPersonCode = _
                            frmDemo.txtPassword.Tag
            '������
                            gProtocol.strProtocStatus = "04 - Manager"
            '�����
                            gProtocol.strProtocTime = _
                            Format(Now, "h:mm:ss")
            '����
                            gProtocol.strProtocDate = _
                            Format(Now, "dd/mm/yyyy")
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            '��������� "���������" �� ���������� ����
                        frmTableCalendar.comCalendar.PreviousDay
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
                
            '�����-���� ����������� - ����� �� ��������� � ����������
                Else
                    GoTo UnAccessable
                End If
                
                
                GoTo EndCycle
UnAccessable:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndCycle
CopyingMistake:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
            End If
EndCycle:
        Next
            '������ ���� ������������
    Else
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 0 (������)
            frmTableSystem.grdTableSystem.Col = 0
            '��������� ������������
            If Trim(frmTableSystem.grdTableSystem.Text) = Trim(cboFileName.Text) Then
            '������� ������� �� ������ ���������������� ���� "cboFileName"
            '   - �����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(cboFileName.Text)
            
            '�������� ������������� �����-����� �������������
                On Error GoTo UnExist
            '�����-���� ������� - ����������
                If (FSO.FolderExists(strPathFolderName)) Then
            '���������������� ������� - "�������� ������ �������������"
                    gProtocol.strProtocName = strPathFolderName
            '��������� ������
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                    gProtocol.strProtocStatus = "04 - Manager"
            '�����
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "DownLoad Archives"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                    On Error GoTo CopyingError
                    
            '��������� "���������" �� ������� ����
                    frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                    For intDayArchiveCopy = 1 To gDayNum Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                        frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ �������������
            '  (� ��������� "����" � ����)
                        strPathFileName = strPathFolderName + "\" + _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        gProtocol.strProtocReserve = _
                        Trim(frmTableSystem.grdTableSystem.Text)
                        If frmTableCalendar.comCalendar.Day < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Day)
                        End If
                        If frmTableCalendar.comCalendar.Month < 10 Then
                            strPathFileName = strPathFileName + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_0" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        Else
                            strPathFileName = strPathFileName + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
            '������������ ����������
                            gProtocol.strProtocReserve = _
                            Trim(gProtocol.strProtocReserve) + "_" + _
                            CStr(frmTableCalendar.comCalendar.Month)
                        End If
                        strPathFileName = strPathFileName + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)
            '������������ ����������
                        gProtocol.strProtocReserve = _
                        Trim(gProtocol.strProtocReserve) + "_" + _
                        Right(CStr(frmTableCalendar.comCalendar.Year), 2)

                        If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ������ � "Host Computer"
                            FSO.CopyFile strPathFileName, _
                            strHostFileName
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        
            '���������������� ������� - "����������� ������ �������������"
                            gProtocol.strProtocName = "Copy Archive"
            '��������� ������
                            gProtocol.strProtocPersonCode = _
                            frmDemo.txtPassword.Tag
            '������
                            gProtocol.strProtocStatus = "04 - Manager"
            '�����
                            gProtocol.strProtocTime = _
                            Format(Now, "h:mm:ss")
            '����
                            gProtocol.strProtocDate = _
                            Format(Now, "dd/mm/yyyy")
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            '��������� "���������" �� ���������� ����
                        frmTableCalendar.comCalendar.PreviousDay
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
                
            '�����-���� ����������� - ����� �� ��������� � ����������
                Else
                    GoTo UnExist
                End If
                
                GoTo EndProcedure
UnExist:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " impossible !", vbExclamation, "Error"
                GoTo EndProcedure
CopyingError:
            '������ �������� ������
                frmDemo.BeepSound
                MsgBox "The downloading from " + strPathFolderName + _
                " error !", vbExclamation, "Error"
                GoTo EndProcedure
            End If
        Next
EndProcedure:
    End If
            
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub
            
            '������������ ��� ��������� � ����������� � ������� ACCESS"
Public Sub BasesConvert()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            
            '���������� ����������� ������ ���� ��� ������ "Preprocessors"
            '  - ��� �������� � ��������� �������� �����
    frmPreprocessors.MousePointer = 0
            
            '����� ��������� ��������� ������� "cmdProtocolBase_Click"
    Call cmdProtocolBase_Click
            '����� ��������� ��������� ������� "cmdBookKeeperBase_Click"
    Call cmdBookKeeperBase_Click

End Sub
            
            '��������� ������� "Protocol Base"
Private Sub cmdProtocolBase_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '���������� ����� � "���� ���������"
Dim lngProtocolBaseCount As Long
            '����� ����� ������
Dim intFileNum As Integer
            '����� ������ "������� ���������" � DUMMY �����
Dim lngRecordLen As Long
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
Dim strDummyFileName As String
            '������� ����� ������ ������� DUMMY �����
Dim lngRowDummy As Long
            '������ ��� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ��� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � DUMMY ����
Dim intDayArchive As Integer
            '���������� ����� � ���������� ����� (������ ��� "TableProtocol")
Dim intRowQuan As Integer
            '������� ����� ������ ����������� ������
            '   ��� ������� "TableProtocol"
Dim intRowNumArchive As Integer

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            '������ ��� ����� "������� ��������� "(� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
    gFileDummy = FreeFile
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            '��������� ������� � ������ ����� DUMMY �����(�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '������� "������" DUMMY ����, ���� �� ����������
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
    
            '��������� ������
    On Error GoTo UnDefError
            '������� DUMMY ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            '������� �����  ��������� ������ DUMMY �����
    gDummyRowNum = 1
            
            '������� ��� �������������: "All"
    If Trim(cboPreprocessors.Text) = "All" Then
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 1 (��� �������������)
                frmTableSystem.grdTableSystem.Col = 1
            
            '��������� "���������" �� ������� ����
                frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
                    frmTableCalendar.comCalendar.PreviousDay
                Next
            '���� �� ���� �����, ������� � ��������� ����
                For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                    frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
                    If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                        intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '���� �� ���� ������� ������
                        For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                            Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                            WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� ������
                        Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                        gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                    End If
            '��������� "���������" �� ��������� ����
                    frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                    
                Next
            '������� ������� "��������� �������" = 0 (������)
                frmTableSystem.grdTableSystem.Col = 0
            '�����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� ������������� � ����� DUMMY �����
                If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" �������������
                    intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                    intFileNum = FreeFile
            '������� ���� "������� ���������" ������������� ���
            '   ������������� �������
                    Open strPathFileName For Random As intFileNum _
                    Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" �������������
                    For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                        Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                        WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    Next
            '������� ���� "������� ���������" �������������
                    Close intFileNum
                        
            '���������������� ������� - "���������� �������� �������������"
                    gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                    gProtocol.strProtocStatus = "04 - Manager"
            '�����
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "Protocol From Preproc"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                End If
            
            '������������ �������������� ������
                gProtocol.strProtocName = "=========="
                gProtocol.strProtocPersonCode = "=========="
                gProtocol.strProtocStatus = "=========="
                gProtocol.strProtocTime = "=========="
                gProtocol.strProtocDate = "=========="
                gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
                WriteDummy
            
            '���������� �������������� "����" � �������� ����������� ���������
                strPathFolderName = App.Path
                If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
                    strPathFolderName = Left(strPathFolderName, _
                    Len(strPathFolderName) - 1)
                End If
            
            End If
        Next
        
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   "Host Computer'a" (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� "Host Computer'a" DUMMY �����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" "Host Computer'a" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" "Host Computer'a" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������" "Host Computer'a"
            Close intFileNum
                        
            '���������������� ������� - "���������� �������� "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "Protocol From Host"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
            
            '������ ���� ������������
    Else
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 1 (��� �������������)
                frmTableSystem.grdTableSystem.Col = 1
            
            '��������� ������������ ������ - ����������
                If Trim(cboPreprocessors.Text) = _
                Trim(frmTableSystem.grdTableSystem.Text) Then
            
            '��������� "���������" �� ������� ����
                    frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                    For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
                        frmTableCalendar.comCalendar.PreviousDay
                    Next
            '���� �� ���� �����, ������� � ��������� ����
                    For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                        frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
                        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                            intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                            Open strPathFileName For Random As intFileNum _
                            Len = lngRecordLen
            '���� �� ���� ������� ������
                            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                                frmPreprocessors.MousePointer = vbHourglass
                            Next
            '������� ���� ������
                            Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                            gProtocol.strProtocStatus = "04 - Manager"
            '�����
                            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            '��������� "���������" �� ��������� ����
                        frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
            '������� ������� "��������� �������" = 0 (������)
                    frmTableSystem.grdTableSystem.Col = 0
            '�����-���� ������������� (� ������ ����� � ���)
                    strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� ������������� � ����� DUMMY �����
                    If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" �������������
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                        intFileNum = FreeFile
            '������� ���� "������� ���������" ������������� ���
            '   ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" �������������
                        For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                            Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                            WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� "������� ���������" �������������
                        Close intFileNum
                        
            '���������������� ������� - "���������� �������� �������������"
                        gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                        gProtocol.strProtocReserve = "Protocol From Preproc"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                    
                    End If
            '��������� ����
                    Exit For
                End If
            End If
        Next
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
            
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFolderName = App.Path
        If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
            strPathFolderName = Left(strPathFolderName, _
            Len(strPathFolderName) - 1)
        End If
    
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   "Host Computer'a" (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� "Host Computer'a" DUMMY �����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" "Host Computer'a" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" "Host Computer'a" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������" "Host Computer'a"
            Close intFileNum
                        
            '���������������� ������� - "���������� �������� "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "Protocol From Host"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
    
    End If
            
            '���������� �������������� "����" � ��������
            '  ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            '��������� ������� �������� "Data" ������� � "���� ���������"
    frmDemo.datBase.DatabaseName = strPathFileName + "ProtocolBase.mdb"
    frmDemo.datBase.RecordSource = "Protocol"
            
            '���������� ���������� ������� � "���� ���������"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngProtocolBaseCount = frmDemo.datBase.Recordset.RecordCount
            '�������� "���� ���������"
    frmDemo.datBase.Recordset.MoveFirst
            '���� �� ���� ������� DUMMY �����
    For lngRowDummy = 1 To gDummyRowNum - 1 Step 1
            '��������� ���������� ��� ��������� ��������� �������
        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
        frmPreprocessors.MousePointer = vbHourglass
            '������ ������ DUMMY ����� � �����
        Get gFileDummy, lngRowDummy, gProtocol
            '�������� ������� ������ "���� ���������"
        frmDemo.datBase.Recordset.Edit
        frmDemo.datBase.Recordset.Fields("Name").Value = gProtocol.strProtocName
        frmDemo.datBase.Recordset.Fields("CodeOrPassword").Value = _
        gProtocol.strProtocPersonCode
        frmDemo.datBase.Recordset.Fields("Status").Value = gProtocol.strProtocStatus
        frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
        frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
        frmDemo.datBase.Recordset.Fields("ReservOrNote").Value = gProtocol.strProtocReserve
        frmDemo.datBase.Recordset.Update
            '�� ��������� ������ ������ "���� ���������"
        If lngRowDummy < lngProtocolBaseCount Then
            frmDemo.datBase.Recordset.MoveNext
            '��������� ������ ������ "���� ���������"
        Else
            frmDemo.datBase.Recordset.AddNew
            frmDemo.datBase.Recordset.Update
            frmDemo.datBase.Recordset.MoveNext
        End If
    Next
            '�������� ����� ������ ������ ��  "���� ���������"
    If lngRowDummy > lngProtocolBaseCount Then
        frmDemo.datBase.Recordset.Delete
            '�������� ������ ������� ��  "���� ���������"
    Else
        For lngRowDummy = lngRowDummy To lngProtocolBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
        Next
    End If
            
            '���������������� ������� - "������������ ���� ���������"
    gProtocol.strProtocName = "ProtocolBase"
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Creation"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol
    
    GoTo EndProcedure
            '�������������� ������
UnDefError:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            '������� DUMMY ����
    Close gFileDummy
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '��������� ������� "BookKeeper Base"
Private Sub cmdBookKeeperBase_Click()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ����� ������
Dim intFileNum As Integer
            '����� ������ "������� ���������" � DUMMY �����
Dim lngRecordLen As Long
            '������� ������� "\" � ������ ����� �����
Dim intSymbPos As Integer
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
Dim strDummyFileName As String
            '������� ����� ������ ������� DUMMY �����
Dim lngRowDummy As Long
            '������ ��� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '����� ��� (�������� ������, ������� � �������� ���),
            '  ������� ��������������� �������� ��� �����������
            '  ������� ������������ � DUMMY ����
Dim intDayArchive As Integer
            '���������� ����� � ���������� ����� (������ ��� "TableProtocol")
Dim intRowQuan As Integer
            '������� ����� ������ ����������� ������
            '   ��� ������� "TableProtocol"
Dim intRowNumArchive As Integer
            '���������� ����� � "���� �����������"
Dim lngBookKeepingBaseCount As Long
            '������� ����� ����������������� ������ "���� �����������"
Dim lngBookKeepingRowNum As Long

            '���� ������ ���� = "�������� ����", �� �����
    If frmPreprocessors.MousePointer = vbHourglass Then
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If

            '���� ������������� �����������, �� �����
    If cboFileName.Text = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The preprocessors  are missing !", vbExclamation, "Error"
            '������ � ������ �����
        frmPreprocessors.Hide
        Exit Sub
    End If
            
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFolderName = App.Path
    If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
        strPathFolderName = Left(strPathFolderName, _
        Len(strPathFolderName) - 1)
    End If
            
            '������ ��� ����� "������� ��������� "(� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� �����
    gFileDummy = FreeFile
            '������ ��� DUMMY ����� (� ��������� "����" � ����)
    strPathFileName = strPathFolderName + "\Dummy.dat"
    
            '��������� ������� � ������ ����� DUMMY �����(�� ��������� "C:\")
    intSymbPos = 4
            '����� ��������� ������� ���������� ����� �����
    Do While InStr(intSymbPos, strPathFileName, "\") <> 0
        If InStr(intSymbPos, strPathFileName, "\") <> 0 Then intSymbPos = _
        InStr(intSymbPos, strPathFileName, "\") + 1
    Loop
            '������� "������" DUMMY ����, ���� �� ����������
    If Dir(strPathFileName) = Mid(strPathFileName, intSymbPos) Then
        Kill strPathFileName
    End If
        
            '��������� ������
    On Error GoTo UnDefError
            '������� DUMMY ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As gFileDummy Len = lngRecordLen
            '������� �����  ��������� ������ DUMMY �����
    gDummyRowNum = 1
            
            '������� ��� �������������: "All"
    If Trim(cboPreprocessors.Text) = "All" Then
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 1 (��� �������������)
                frmTableSystem.grdTableSystem.Col = 1
            
            '��������� "���������" �� ������� ����
                frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
                    frmTableCalendar.comCalendar.PreviousDay
                Next
            '���� �� ���� �����, ������� � ��������� ����
                For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                    frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
                    If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                        intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '���� �� ���� ������� ������
                        For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                            Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                            WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� ������
                        Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                        gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                    End If
            '��������� "���������" �� ��������� ����
                    frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                    
                Next
            '������� ������� "��������� �������" = 0 (������)
                frmTableSystem.grdTableSystem.Col = 0
            '�����-���� ������������� (� ������ ����� � ���)
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� ������������� � ����� DUMMY �����
                If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" �������������
                    intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                    intFileNum = FreeFile
            '������� ���� "������� ���������" ������������� ���
            '   ������������� �������
                    Open strPathFileName For Random As intFileNum _
                    Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" �������������
                    For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                        Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                        WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    Next
            '������� ���� "������� ���������" �������������
                    Close intFileNum
                        
            '���������������� ������� - "���������� �������� �������������"
                    gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                    gProtocol.strProtocStatus = "04 - Manager"
            '�����
                    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "Protocol From Preproc"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                    frmDemo.WriteProtocol
                    
                End If
            
            '������������ �������������� ������
                gProtocol.strProtocName = "=========="
                gProtocol.strProtocPersonCode = "=========="
                gProtocol.strProtocStatus = "=========="
                gProtocol.strProtocTime = "=========="
                gProtocol.strProtocDate = "=========="
                gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
                WriteDummy
            
            '���������� �������������� "����" � �������� ����������� ���������
                strPathFolderName = App.Path
                If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
                    strPathFolderName = Left(strPathFolderName, _
                    Len(strPathFolderName) - 1)
                End If
            
            End If
        Next
        
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   "Host Computer'a" (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� "Host Computer'a" DUMMY �����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" "Host Computer'a" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" "Host Computer'a" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������" "Host Computer'a"
            Close intFileNum
                        
            '���������������� ������� - "���������� �������� "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "Protocol From Host"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
            
            '������ ���� ������������
    Else
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 _
        Step 1
            '������� ������ "��������� �������"
            frmTableSystem.grdTableSystem.Row = intRowNum
            '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
            '���="03" - Preprocessor
            If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 1 (��� �������������)
                frmTableSystem.grdTableSystem.Col = 1
            
            '��������� ������������ ������ - ����������
                If Trim(cboPreprocessors.Text) = _
                Trim(frmTableSystem.grdTableSystem.Text) Then
            
            '��������� "���������" �� ������� ����
                    frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
                    For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
                        frmTableCalendar.comCalendar.PreviousDay
                    Next
            '���� �� ���� �����, ������� � ��������� ����
                    For intDayArchive = 1 To gDayNum + 1 Step 1
            '������� ������� "��������� �������" = 1 (��� �������������)
                        frmTableSystem.grdTableSystem.Col = 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
                        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                            intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                            Open strPathFileName For Random As intFileNum _
                            Len = lngRecordLen
            '���� �� ���� ������� ������
                            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                                frmPreprocessors.MousePointer = vbHourglass
                            Next
            '������� ���� ������
                            Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                            gProtocol.strProtocStatus = "04 - Manager"
            '�����
                            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                            frmDemo.WriteProtocol
                        End If
            '��������� "���������" �� ��������� ����
                        frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
                        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                        frmPreprocessors.MousePointer = vbHourglass
                    
                    Next
            '������� ������� "��������� �������" = 0 (������)
                    frmTableSystem.grdTableSystem.Col = 0
            '�����-���� ������������� (� ������ ����� � ���)
                    strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   ������������� (� ��������� "����" � ����)
                    strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� ������������� � ����� DUMMY �����
                    If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" �������������
                        intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                        intFileNum = FreeFile
            '������� ���� "������� ���������" ������������� ���
            '   ������������� �������
                        Open strPathFileName For Random As intFileNum _
                        Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" �������������
                        For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" ������������� �� ����� � �����
                            Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                            WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                            frmPreprocessors.MousePointer = vbHourglass
                        Next
            '������� ���� "������� ���������" �������������
                        Close intFileNum
                        
            '���������������� ������� - "���������� �������� �������������"
                        gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                        gProtocol.strProtocStatus = "04 - Manager"
            '�����
                        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                        gProtocol.strProtocReserve = "Protocol From Preproc"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                        frmDemo.WriteProtocol
                    
                    End If
            '��������� ����
                    Exit For
                End If
            End If
        Next
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
            
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFolderName = App.Path
        If Right(strPathFolderName, 1) = "\" Then
            '������ ��� ����� "Host Computera" ��� DUMMY �����
            '  (� ��������� "����" � ���)
            strPathFolderName = Left(strPathFolderName, _
            Len(strPathFolderName) - 1)
        End If
    
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            
            '���� �� ���� �����, ������� � ������� ����
        For intDayArchive = 1 To gDayNum Step 1
            '��������� "���������" �� ���������� ����
            frmTableCalendar.comCalendar.PreviousDay
        Next
            '���� �� ���� �����, ������� � ��������� ����
        For intDayArchive = 1 To gDayNum + 1 Step 1
            '������ ��� ����������� ������ (� ��������� "����" � ����)
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

            '���� ������ �������
            If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � ������
                intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
                intFileNum = FreeFile
            '������� ���� ������ ��� ������������� �������
                Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� ������
                For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ ������ �� ����� � �����
                    Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                    WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                    DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                    frmPreprocessors.MousePointer = vbHourglass
                Next
            '������� ���� ������
                Close intFileNum
                            
            '���������������� ������� - "����������� ������ � DUMMY ����"
                gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
                gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
                gProtocol.strProtocStatus = "04 - Manager"
            '�����
                gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������������ ����������
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
            
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
                frmDemo.WriteProtocol
            End If
            '��������� "���������" �� ��������� ����
            frmTableCalendar.comCalendar.NextDay
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
                    
        Next
            '������ ��� ����������� ����� ������� "TableProtocol"
            '   "Host Computer'a" (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\TableProtocol.dat"
            '���� ������� - ����������� ����� ������� "TableProtocol"
            '    �� �����-����� "Host Computer'a" DUMMY �����
        If (FSO.FileExists(strPathFileName)) Then
            '���������� ����� � "������� ���������" "Host Computer'a"
            intRowQuan = FileLen(strPathFileName) / lngRecordLen
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '������� ���� "������� ���������" "Host Computer'a" ���
            '   ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������" "Host Computer'a"
            For intRowNumArchive = 1 To intRowQuan Step 1
            '������ ������ "������� ���������" "Host Computer'a" �� ����� � �����
                Get intFileNum, intRowNumArchive, gProtocol
            '�������� ������ � DUMMY ����
                WriteDummy
            '��������� ���������� ��� ��������� ��������� �������
                DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
                frmPreprocessors.MousePointer = vbHourglass
            Next
            '������� ���� "������� ���������" "Host Computer'a"
            Close intFileNum
                        
            '���������������� ������� - "���������� �������� "Host'a" "
            gProtocol.strProtocName = "Copy To Dummy"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = "Protocol From Host"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
            frmDemo.WriteProtocol
                    
        End If
            
            '������������ �������������� ������
        gProtocol.strProtocName = "=========="
        gProtocol.strProtocPersonCode = "=========="
        gProtocol.strProtocStatus = "=========="
        gProtocol.strProtocTime = "=========="
        gProtocol.strProtocDate = "=========="
        gProtocol.strProtocReserve = "=========="
            '�������� ������ � DUMMY ����
        WriteDummy
    
    End If
        
            '���������� �������������� "����" � ��������
            '  ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
            
            '��������� ������� �������� "Data" ������� � "���� �����������"
    frmDemo.datBase.DatabaseName = strPathFileName + "BookKeepingBase.mdb"
    frmDemo.datBase.RecordSource = "BookKeeping"
            
            '���������� ���������� ������� � "���� �����������"
    frmDemo.datBase.Refresh
    frmDemo.datBase.Recordset.MoveLast
    lngBookKeepingBaseCount = frmDemo.datBase.Recordset.RecordCount
            '�������� "���� �����������"
    frmDemo.datBase.Recordset.MoveFirst
            '������� ����� ����������������� ������ "���� �����������"
    lngBookKeepingRowNum = 0
    For lngRowDummy = 0 To gDummyRowNum - 1 Step 1
            '��������� ���������� ��� ��������� ��������� �������
        DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
        frmPreprocessors.MousePointer = vbHourglass
            '������� ������ ������������ "���������" ������
        If lngRowDummy = 0 Then
            '��������������� ������� ������ "���� �����������"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = "Fiktive Record"
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = "0000000000000000"
            frmDemo.datBase.Recordset.Fields("Status").Value = "00"
            frmDemo.datBase.Recordset.Fields("Time").Value = "00:00:00AM"
            frmDemo.datBase.Recordset.Fields("Date").Value = "01.01.2000"
            '���������� ������ � "���� �����������"
            frmDemo.datBase.Recordset.Update
        Else
            '������ ������ DUMMY ����� � �����
            Get gFileDummy, lngRowDummy, gProtocol
            '��������������� ������� ������ "���� �����������"
            frmDemo.datBase.Recordset.Edit
            frmDemo.datBase.Recordset.Fields("Person").Value = gProtocol.strProtocName
            frmDemo.datBase.Recordset.Fields("PersonCode").Value = gProtocol.strProtocPersonCode
            frmDemo.datBase.Recordset.Fields("Status").Value = Left(Trim(gProtocol.strProtocStatus), 2)
            frmDemo.datBase.Recordset.Fields("Time").Value = gProtocol.strProtocTime
            frmDemo.datBase.Recordset.Fields("Date").Value = gProtocol.strProtocDate
            '������� ���������:
            '                                  - ����/����� ("18"/"19") ��������� ���
            '                                  - ��������������� ("16) ��������� ���
            '                                  - ������������ ("17") ��������� ���
            '                                  - ����������� ("12") �������� ������� ����������� ���
            '                                  - ���������� ("13") �������� ������� �����������
            '                                  - ����������� ("14") �������� ���������� ����������� ���
            '                                  - ���������� ("15") �������� ���������� �����������
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
            '�������  - ��������������� ��������� (���������� �������
            '  � "���� �����������")
                If Trim(gProtocol.strProtocReserve) = "AutoRegistration" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "16"
            '�������  - ������������ ��������� (���������� �������
            '  � "���� �����������")
                ElseIf Trim(gProtocol.strProtocReserve) = "AutoDelete" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "17"
            '�������  - ���� ��������� �� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 5) = "Input" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "18"
            '�������  - ����� ��������� � ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Right(Trim(gProtocol.strProtocReserve), 6) = "Output" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "19"
            '�������  - ����������� ������� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "12"
            '�������  - ���������� ������� ����������� (���������� �������
            '  � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "13"
            '�������  - ����������� ���������� ����������� (����������
            '  ������� � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "14"
            '�������  - ���������� ���������� ����������� (����������
            '  ������� � "���� �����������")
                ElseIf Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce" Then
                    frmDemo.datBase.Recordset.Fields("Status").Value = "15"
                End If
            '���������� ������ � "���� �����������"
                frmDemo.datBase.Recordset.Update
            '������� ����� ����������������� ������ "���� �����������"
                lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            '�� ��������� ������ ������ "���� �����������"
                If lngBookKeepingRowNum < lngBookKeepingBaseCount Then
                    frmDemo.datBase.Recordset.MoveNext
            '��������� ������ ������ "���� �����������"
                Else
                    frmDemo.datBase.Recordset.AddNew
                    frmDemo.datBase.Recordset.Update
                    frmDemo.datBase.Recordset.MoveNext
                End If
            End If
        End If
    Next
            '������� ����� ������ "���� �����������"
    lngBookKeepingRowNum = lngBookKeepingRowNum + 1
            '�������� ����� ������ ������ ��  "���� �����������"
    If lngBookKeepingRowNum > lngBookKeepingBaseCount Then
        frmDemo.datBase.Recordset.Delete
            '�������� ������ ������� ��  "���� �����������",
            '  ����� ������������
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum = 1 Then
        frmDemo.datBase.Recordset.MoveFirst
        frmDemo.datBase.Recordset.MoveNext
        For lngBookKeepingRowNum = 2 To lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
        Next
            '�������� ������ ������� ��  "���� �����������"
    ElseIf lngBookKeepingBaseCount <> 1 And _
    lngBookKeepingRowNum <> 1 Then
        For lngBookKeepingRowNum = lngBookKeepingRowNum To _
        lngBookKeepingBaseCount Step 1
            frmDemo.datBase.Recordset.Delete
            frmDemo.datBase.Recordset.MoveNext
            '��������� ���������� ��� ��������� ��������� �������
            DoEvents
            '�������� ����������� ������ ����  �� "�������� ����"
            frmPreprocessors.MousePointer = vbHourglass
        Next
    End If
            
            '���������������� ������� - "������������ ���� ���������"
    gProtocol.strProtocName = "BookKeeperBase"
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "Creation"

            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol
            
    GoTo EndProcedure
            '�������������� ������
UnDefError:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox "UnDefined Error !", vbExclamation, "Error"

EndProcedure:
            '������� DUMMY ����
    Close gFileDummy
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0
            '������ � ������ �����
    frmPreprocessors.Hide

End Sub

            '����������� ������ �� ������������� � "Host Computer"
Public Function ArchiveCopy(ByVal strMessage As String)
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer
            '������ ��� ����������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '������ ��� �����-����� (� ��������� "����" � ����)
Dim strHostFileName As String
            '������ ��� ���������� �����-����� (� ��������� "����" � ���)
Dim strPathFolderName As String
            '������ "FileSystemObject" - "�������� �������"
Dim FSO As Variant
            '������� ������� "_" � ����� �����
Dim intSymbPos As Integer

            '��������� ������� � ������ ����� ����� (�� ��������� "Archive ")
    intSymbPos = 9
            '����� �������� ������� ���������� ����� �����
    If InStr(intSymbPos, strMessage, "_") <> 0 Then
        intSymbPos = InStr(intSymbPos, strMessage, "_")
    Else
        intSymbPos = Len(strMessage)
    End If
            
            '������� ������� "��������� �������" = 2 (���)
    frmTableSystem.grdTableSystem.Col = 2
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = intRowNum
            '���="03" - Preprocessor (������������)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 1 (��� �������������)
            frmTableSystem.grdTableSystem.Col = 1
            '��������� ��� �������������
            If Trim(frmTableSystem.grdTableSystem.Text) = _
            Mid(strMessage, 9, intSymbPos - 9) Then
            '������� ������� "��������� �������" = 0 (������)
                frmTableSystem.grdTableSystem.Col = 0
            '������ ��� �����-����� �������������
            '   (� ��������� "����" � ���) ��� "Whole"
                strPathFolderName = Trim(frmTableSystem.grdTableSystem.Text)
                Exit For
            End If
            frmTableSystem.grdTableSystem.Col = 2
        End If
    Next
            
            '������� ������ "FSO" - "�������� �������"
    Set FSO = CreateObject("Scripting.FileSystemObject")
            '���������� �������������� "����" � �������� ����������� ���������
    strHostFileName = App.Path
    If Right(strHostFileName, 1) <> "\" Then
            '������ ��� ����� "Host Computera" ��� �����-�����
            '  �������������(� ��������� "����" � ���)
        strHostFileName = strHostFileName + "\"
    End If
    
            '�������� ����������� ������ ����  �� "�������� ����"
    frmPreprocessors.MousePointer = vbHourglass
            
            '�������� ������������� �����-����� �������������
    On Error GoTo UnExist
            '�����-���� ������� - ����������
    If (FSO.FolderExists(strPathFolderName)) Then
            '���������������� ������� - "�������� ������ �������������"
        gProtocol.strProtocName = strPathFolderName
            '��������� ������
        gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
        gProtocol.strProtocStatus = "04 - Manager"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "DownLoad Archives"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
        frmDemo.WriteProtocol
                    
        On Error GoTo CopyingError
                    
            '������ ��� ����������� ������ �������������
            '  (� ��������� "����" � ����)
        strPathFileName = strPathFolderName + "\" + _
        Mid(strMessage, 9)

        If (FSO.FileExists(strPathFileName)) Then
            '���� ������� - ����������� ������ � "Host Computer"
            FSO.CopyFile strPathFileName, strHostFileName
                        
            '���������������� ������� - "����������� ������ �������������"
            gProtocol.strProtocName = "Copy Archive"
            '��������� ������
            gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
            gProtocol.strProtocStatus = "04 - Manager"
            '�����
            gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
            gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
            gProtocol.strProtocReserve = Mid(strMessage, 9, intSymbPos - 9)
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
            frmDemo.WriteProtocol
        End If
                
            '�����-���� ����������� - ����� �� ��������� � ����������
    Else
        GoTo UnExist
    End If
                
    GoTo EndProcedure

UnExist:
            '���������������� ������� - "������ ��� ����������� ������"
    gProtocol.strProtocName = strPathFolderName
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "DownLoad Archives Err"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol
    GoTo EndProcedure

CopyingError:
            '���������������� ������� - "������ ��� ����������� ������"
    gProtocol.strProtocName = "Copy Archive"
            '��������� ������
    gProtocol.strProtocPersonCode = frmDemo.txtPassword.Tag
            '������
    gProtocol.strProtocStatus = "04 - Manager"
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = _
    Trim(gProtocol.strProtocReserve) + " Err"
            '�������� ������ � ���� "������� ���������" "Host Computer'a"
    frmDemo.WriteProtocol

EndProcedure:
            '������������ ����������� ������ ����
    frmPreprocessors.MousePointer = 0
    On Error GoTo 0

End Function

            '�������� �����
Private Sub Form_Load()
            '������� ����� ������ ������� "TableSystem"
            '   � "Host Computer'e"
Dim intRowNum As Integer

            '���������� ����������� ������ ����
    frmPreprocessors.MousePointer = 0
            
            '������� ������� "��������� �������" = 2 (���)
    frmTableSystem.grdTableSystem.Col = 2
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        frmTableSystem.grdTableSystem.Row = intRowNum
            '���="03" - Preprocessor (������������)
        If Left(Trim(frmTableSystem.grdTableSystem.Text), 2) = "03" Then
            '������� ������� "��������� �������" = 0 (������)
            frmTableSystem.grdTableSystem.Col = 0
            '��������� ������ ���������������� ���� "cboFileName"
            If cboFileName.ListCount = 0 Then
                cboFileName.AddItem "Whole"
                cboPreprocessors.AddItem "All"
            End If
            cboFileName.AddItem _
            frmTableSystem.grdTableSystem.Text
            '��� ������������� ��������� ����
            cboPreprocessors.AddItem _
            gSocketNet(cboPreprocessors.ListCount)
        '������� ������� "��������� �������" = 2 (���)
            frmTableSystem.grdTableSystem.Col = 2
        End If
    Next
            '����������� ��������� ������
    On Error GoTo UnLoad
            '������� ������ ������� ������: "��� �������������"
    cboFileName.ListIndex = 0
    cboPreprocessors.ListIndex = 0
UnLoad:

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            '��������� ������ ������ � DUMMY ����
Public Sub WriteDummy()

            '�������� ������ � ���� "������� ���������"
    Put gFileDummy, gDummyRowNum, gProtocol
            '����� ��������� ��������� ������ DUMMY �����
    gDummyRowNum = gDummyRowNum + 1
    
End Sub

