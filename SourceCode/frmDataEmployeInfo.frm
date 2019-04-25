VERSION 5.00
Begin VB.Form frmDataEmployeInfo 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmployeInfoData"
   ClientHeight    =   3975
   ClientLeft      =   7440
   ClientTop       =   2925
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4260
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   12
      Tag             =   "0"
      ToolTipText     =   "Info"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Pressing"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   2280
      Top             =   240
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   10
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtParole 
      Enabled         =   0   'False
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   8
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox txtEmployeReg 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtEmployeOut 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtEmployeIn 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3360
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.ListBox lstInfo 
      Height          =   690
      ItemData        =   "frmDataEmployeInfo.frx":0000
      Left            =   720
      List            =   "frmDataEmployeInfo.frx":0002
      TabIndex        =   2
      ToolTipText     =   "Information"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox lstPersonCode 
      Height          =   690
      ItemData        =   "frmDataEmployeInfo.frx":0004
      Left            =   720
      List            =   "frmDataEmployeInfo.frx":0006
      TabIndex        =   1
      ToolTipText     =   "PersonCode"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
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
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label lblParole 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Parole"
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
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgEmployeReg 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":0008
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgEmployeOut 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":045A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgEmployeIn 
      Height          =   375
      Left            =   2880
      Picture         =   "frmDataEmployeInfo.frx":08A0
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblPersonCode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "#### "
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
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Info "
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
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgEmployeInfo 
      Height          =   615
      Left            =   3360
      Picture         =   "frmDataEmployeInfo.frx":0CE6
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "frmDataEmployeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             '��������� ������
Dim strPassword As String

            '������ ������ � "������� ������" (������ "Pressing")
Private Sub cmdOK_Click()
            '��� �������� ��� ������ ������ � "������� ������"
Dim intAutoPressingCode  As Integer

            '����� ���������-������� ������ ������
            '  � "������� ������"
    intAutoPressingCode = frmTablePerson.AutoPresEmploye()
            
            '������������ �����
    If intAutoPressingCode = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ���������
        If frmDemo.optEnglish = True Then
            MsgBox ("The Visitors are Present")
        Else
            MsgBox ("Viesi ir")
        End If
            
            '���������� - �����������
        gProtocol.strProtocName = "PRESSING TabPers"
            '������������ ��� - �����������
        gProtocol.strProtocPersonCode = "PRESSING TabPers"
            '������
        gProtocol.strProtocStatus = "04 - Operator"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Pressing EmployeInfo"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ ������ � "������� ������"
        Me.Tag = 1
            '������� � ��������� ���������
        cmdCancel_Click
            '������ ���������� - ���������������� �������
    ElseIf intAutoPressingCode = 2 Then
            '���������� - �����������
        gProtocol.strProtocName = "PRESSING TabPers"
            '������������ ��� - �����������
        gProtocol.strProtocPersonCode = "PRESSING TabPers"
            '������
        gProtocol.strProtocStatus = "04 - Operator"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid Press. Employe"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ �� ������ ������ � "������� ������"
        Me.Tag = 2
            '������� � ��������� ���������
        cmdCancel_Click
    End If
            
End Sub
            
            '������� � ��������� ��������� (������ "Cancel _ Exit")
Private Sub cmdCancel_Click()
        '������� ������ �� ������ ������ � "������e ������"
    If Me.Tag = 0 Then Me.Tag = 2
            '������� ��������� ������� �����
    Me.Visible = False
            
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '����������� ������� �����
Private Sub Form_Activate()
            '������� ����� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            
            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������, �������� ��������� ����)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessPlus
            
            '�������� ��������� ����
    txtEmployeIn.Text = "0"
    txtEmployeOut.Text = "0"
    txtEmployeReg.Text = "0"
    txtParole.Text = ""
    txtPersonCode.Text = ""
    txtInfo.Text = ""
             '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
            '������� ������������ ��������� ���� "Parole" � "PersonCode"
    txtParole.Enabled = False
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
            '������� ������������ ������ "lstInfo" � "lstPersonCode"
    lstInfo.Enabled = False
    lstPersonCode.Enabled = False
            '�������� ������ "lstInfo" � "lstPersonCode"
    lstInfo.Clear
    lstPersonCode.Clear
            ' "������� ������" �� �������� ��������������� �����
    If gTablePerson.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ��������� ����������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TablePerson is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
            '���������� ����� �� ������ "Exit_Cancel"
        cmdCancel.SetFocus
        Exit Sub
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� ����������
            If Left(Trim(gTablePerson.Text), 2) = "00" Or _
            Left(Trim(gTablePerson.Text), 2) = "01" Then
            '������� - "PersonCode"
                gTablePerson.Col = 1
            '���� ������ "������� ������" �� ������� ���������
                If gTablePerson.Text <> "Deleted" Then
            '���������� ������ "lstPersonCode" �������� �� "������� ������"
                    lstPersonCode.AddItem gTablePerson.Text
            '������� - "Person or Terminal"
                    gTablePerson.Col = 0
            '���������� ������ "lstInfo" �������� �� "������� ������"
                    lstInfo.AddItem gTablePerson.Text
            '���������� ����� ���������� � �����������
                    If Len(Trim(gTablePerson.Text)) < 16 Then
                        txtEmployeReg.Text = Str(CInt(txtEmployeReg.Text) + 1)
                    ElseIf Right(Trim(gTablePerson.Text), 1) = "+" Then
                        txtEmployeIn.Text = Str(CInt(txtEmployeIn.Text) + 1)
                    ElseIf Right(Trim(gTablePerson.Text), 1) = "-" Then
                        txtEmployeOut.Text = Str(CInt(txtEmployeOut.Text) + 1)
                    End If
                End If
            End If
        Next
            '������ ������
        If lstInfo.ListCount = 0 Then
            '������ �������� ������
            frmDemo.BeepSound
            '����� ��������� � ������ ������
            If frmDemo.optEnglish = True Then
                MsgBox ("The ClientList is Empty")
            Else
                MsgBox ("Klientu saraksts ir neaizpild.")
            End If
            '���������� ����� �� ������ "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
            '�������  �������� �������
        lstInfo.ListIndex = 0
        lstPersonCode.ListIndex = 0
            '������� ��������� ��������� ���� "txtParole"
        txtParole.Enabled = True
            
            '���������� ����� �� ��������� ���� "txtParole"
        If txtParole.Enabled = True Then txtParole.SetFocus
             '���������� ���� ���������� ����������� ������� �����
        Me.Tag = 1
            '���������� �������� ������� ����� ������
        tmrParoleTimeOut.Enabled = True
            '����������� ������ ����� �� ��������� ���� "txtParole"
        txtParole_Click
        Exit Sub
    End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtPersonCode.Text = "Unknown Error"
            '���������� ����� �� ������ "Exit_Cancel"
    cmdCancel.SetFocus
             '���������� ���� ���������� ����������� ������� �����
    Me.Tag = 1
    
End Sub

            '������������� ������� �����
Private Sub Form_Deactivate()
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub
            
            '�������� ������� �����
Private Sub Form_Load()
            '�������� ��������� ����
    txtEmployeIn.Text = ""
    txtEmployeOut.Text = ""
    txtEmployeReg.Text = ""
    txtParole.Text = ""
    txtPersonCode.Text = ""
    txtInfo.Text = ""
             '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            '�������� ������� ���������� ������ "Alt"+ {"^" � "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������ ������
    If lstInfo.ListCount = 0 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������ ������
        If frmDemo.optEnglish = True Then
            MsgBox ("The List is Empty")
        Else
            MsgBox ("Saraksts ir neaizpild.")
        End If
    Else
            '������������ "������" ���� �� ���������� �������� ������
        If KeyCode = 38 And Shift = 4 And lstInfo.ListIndex <> 0 Then
            '�������  �������� �������
            lstInfo.ListIndex = lstInfo.ListIndex - 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            '������������ "������" ���� �� ��������� �������� ������
        ElseIf KeyCode = 40 And Shift = 4 And _
        lstInfo.ListIndex <> lstInfo.ListCount - 1 Then
            '�������  �������� �������
            lstInfo.ListIndex = lstInfo.ListIndex + 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            '������������ "������" ���� �� ������ �������� ������
        ElseIf KeyCode = 33 And Shift = 4 And lstInfo.ListIndex <> 0 Then
            '�������  �������� �������
            lstInfo.ListIndex = 0
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            '������������ "������" ���� �� ��������� �������� ������
        ElseIf KeyCode = 34 And Shift = 4 And _
        lstInfo.ListIndex <> lstInfo.ListCount - 1 Then
            '�������  �������� �������
            lstInfo.ListIndex = lstInfo.ListCount - 1
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
            '������������ "������" ���� �� ������� �������� ������
        ElseIf (KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or _
        KeyCode = 34) And Shift = 4 Then
            lstPersonCode.ListIndex = lstInfo.ListIndex
            GoTo DataCorrect
        End If
        
    End If
    Exit Sub
DataCorrect:
            '�������� �������oe ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
            '��������� ��������� ���� "PersonCode"
    txtPersonCode.Text = lstPersonCode.Text
    txtInfo.Text = lstInfo.Text
            '������� ��� ���������� ����
    txtPersonCode.BackColor = vbCyan
    txtInfo.BackColor = vbCyan
            '���������� ����� �� ������ "Exit_Cancel"
    cmdCancel.SetFocus

End Sub

            '����� ������ "������� ������" ��� "������" �� ������ "PersonCode"
Private Sub lstPersonCode_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '�������� ��������� ����
        txtPersonCode.Text = ""
        txtInfo.Text = ""
             '����� ��� ���������� ����
        txtPersonCode.BackColor = vbWhite
        txtInfo.BackColor = vbWhite
            '����� ������ "������� ������"
        lstInfo.ListIndex = lstPersonCode.ListIndex
            '����� ���������-������� ����������
            '������������� ���� ��� ����� ���������
        intAutoFindCode = frmTablePerson.AutoFindEmploye(lstPersonCode.Text, _
        txtInfo.Text, strStatus)
            '(����)����� ������������� ���� �������� �������
        If intAutoFindCode = 0 Then
    
            '������ ������� ���������
            
            '������������ ������
            If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            '���� �������� � ��������  ������� - �� �����
                intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
                frmDemo.BeepSound
                If frmDemo.optEnglish = True Then
                    MsgBox "Status Error", intButtonsAndIcons, "Error"
                Else
                    MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                End If
                GoTo UnknownError
            End If
            '��������� ��������� ����
            txtPersonCode.Text = lstPersonCode.Text
            txtInfo.Text = lstInfo.Text
            '������� ��� ���������� ����
            txtPersonCode.BackColor = vbCyan
            txtInfo.BackColor = vbCyan
            '���������� ����� �� ������ "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
        frmDemo.BeepSound
        txtPersonCode.Text = "Unknown Error"
            '���������� ����� �� ������ "Exit_Cancel"
        cmdCancel.SetFocus
    End If

End Sub

            '����� ������ "������� ������" ��� "������" �� ������ "Info"
Private Sub lstInfo_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '�������� ��������� ����
        txtPersonCode.Text = ""
        txtInfo.Text = ""
             '����� ��� ���������� ����
        txtPersonCode.BackColor = vbWhite
        txtInfo.BackColor = vbWhite
            '����� ������ "������� ������"
        lstPersonCode.ListIndex = lstInfo.ListIndex
            '����� ���������-������� ����������
            '������������� ���� ��� ����� ���������
        intAutoFindCode = frmTablePerson.AutoFindEmploye(lstPersonCode.Text, _
        txtInfo.Text, strStatus)
            '(����)����� ������������� ���� �������� �������
        If intAutoFindCode = 0 Then
    
            '������ ������� ���������
            
            '������������ ������
            If Left(Trim(strStatus), 2) <> "00" And Left(Trim(strStatus), 2) <> "01" Then
            '���� �������� � ��������  ������� - �� �����
                intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
                frmDemo.BeepSound
                If frmDemo.optEnglish = True Then
                    MsgBox "Status Error", intButtonsAndIcons, "Error"
                Else
                    MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                End If
                GoTo UnknownError
            End If
            '��������� ��������� ����
            txtPersonCode.Text = lstPersonCode.Text
            txtInfo.Text = lstInfo.Text
            '������� ��� ���������� ����
            txtPersonCode.BackColor = vbCyan
            txtInfo.BackColor = vbCyan
            '���������� ����� �� ������ "Exit_Cancel"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
        frmDemo.BeepSound
        txtInfo.Text = "Unknown Error"
            '���������� ����� �� ������ "Exit_Cancel"
        cmdCancel.SetFocus
    End If

End Sub

            '��������� ��������� "������ ����" �� ���� ������������� ����
Private Sub txtPersonCode_Click()
            
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite

End Sub

            '��������� ����� � ������� "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            '������� ����� ������ ������ "lstPersonCode"
Dim intRowNum As Integer
            
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '������� ��� ���������� ����
        txtPersonCode.BackColor = vbCyan
            '������� �� ������ �������������� ������
        On Error GoTo PersonCodeError
            '������������ ��� � ���������� ���������
        If Len(Trim(txtPersonCode.Text)) > 0 And _
        Len(Trim(txtPersonCode.Text)) < 17 Then
            '����� ������������� ���� ������ 16-� ��������
            If Len(Trim(txtPersonCode.Text)) < 16 Then
            '�������� ����������� ���������� ���������� �����
                txtPersonCode.Text = Left("0000000000000000", _
                16 - Len(Trim(txtPersonCode.Text))) + Trim(txtPersonCode.Text)
            End If
            
            '���� �� ���� ������� ������ "lstPersonCode"
            For intRowNum = 0 To lstPersonCode.ListCount - 1 Step 1
            '������� ������ ������
                lstPersonCode.ListIndex = intRowNum
            '��������� ������ ������ "lstPersonCode" �������
                If Trim(lstPersonCode.Text) = Trim(txtPersonCode.Text) Then
            '����� ������ "������� ������"
                    lstInfo.ListIndex = lstPersonCode.ListIndex
            '���������� ����� �� ������ "Exit_Cancel"
                    cmdCancel.SetFocus
                    Exit Sub
                End If
            Next
        
        End If
        
            '������������ ��� � ������������ ���������
PersonCodeError:
            '������ �������� ������
        frmDemo.BeepSound
        txtPersonCode.Text = "Error"
            '����� ��� ���������� ����
        txtPersonCode.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "PersonCode"
        If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
    End If

End Sub

            '��������� ��������� "������ ����" �� ���� ����������
Private Sub txtInfo_Click()
            
             '����� ��� ���������� ����
    txtInfo.BackColor = vbWhite

End Sub

            '��������� ����� � ������� "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            '������� ����� ������ ������ "lstPersonCode"
Dim intRowNum As Integer
            
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '������� ��� ���������� ����
        txtInfo.BackColor = vbCyan
            '������� �� ������ �������������� ������
        On Error GoTo InfoError
            '���������� � ���������� ���������
        If Len(Trim(txtInfo.Text)) > 0 And _
        Len(Trim(txtInfo.Text)) < 17 Then
            
            '���� �� ���� ������� ������ "lstInfo"
            For intRowNum = 0 To lstInfo.ListCount - 1 Step 1
            '������� ������ ������
                lstInfo.ListIndex = intRowNum
            '��������� ������ ������ "lstInfo" �������
                If InStr(1, Trim(lstInfo.Text), Trim(txtInfo.Text)) <> 0 Then
            '����� ������ "������� ������"
                    lstPersonCode.ListIndex = lstInfo.ListIndex
            '���������� ����� �� ������ "Exit_Cancel"
                    cmdCancel.SetFocus
                    Exit Sub
                End If
            Next
        
        End If
        
            '������������ ��� � ������������ ���������
InfoError:
            '������ �������� ������
        frmDemo.BeepSound
        txtInfo.Text = "Error"
            '����� ��� ���������� ����
        txtInfo.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "Info"
        If txtInfo.Enabled = True Then txtInfo.SetFocus
    End If

End Sub
            
            '��������� �������� ������� ��� ����� ������ - ������� "TimeOut"
Private Sub tmrParoleTimeOut_Timer()
            '������ �������� ������
    frmDemo.BeepSound
    
                '���������������� ������� - "TimeOut" ��� ����� ������
    gProtocol.strProtocName = "????????????????"
            '��������� ������
    gProtocol.strProtocPersonCode = ""
            '������
    gProtocol.strProtocStatus = ""
            '�����
    gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
    gProtocol.strProtocReserve = "PASSWORD TimeOut"
            '�������� ������ � ���� "������� ���������"
    frmDemo.WriteProtocol

            ' "�������" ���� ������ ���������
    txtParole.Text = ""
            '�������� �������� ������� ����� ������
    tmrParoleTimeOut.Enabled = False
            '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
            '� (����)������ ��������
    Me.Tag = 2
            '������� � ��������� ���������
    cmdCancel_Click
    
End Sub

            '��������� ��������� "������ ����" �� ���� ������
Private Sub txtParole_Click()
            
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            '������� ����������� ��������� ����
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
            '������� ������������ ������ "lstInfo" � "lstPersonCode"
    lstInfo.Enabled = False
    lstPersonCode.Enabled = False
            ' "�������" ���� ������ ���������
    strPassword = ""
            '���������� �������� ������� ����� ������
    tmrParoleTimeOut.Enabled = True
           '��������� ������ ���������� �� ���� ������ �� ��� �����
           '  ��� ��������� ������������ �������
    Do While strPassword = "" And tmrParoleTimeOut.Enabled = True
        DoEvents
    Loop

End Sub

            '��������� ����� � ������� ������
Private Sub txtParole_KeyPress(KeyAscii As Integer)
            '������ ������
    If KeyAscii = vbKeyReturn Then
             '������� ��� ���������� ����
        txtParole.BackColor = vbCyan
           '������
        strPassword = txtParole.Text
            '���������������� ������� - "���� ������"
        gProtocol.strProtocName = "????????????????"
            '��������� ������
        gProtocol.strProtocPersonCode = txtParole.Text
            '������
        gProtocol.strProtocStatus = "04 - Operator"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "PASSWORD Input"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
        
            '������ ������������ �������� ������ - ����������
        If txtParole.Text = txtParole.Tag Then
            '������� ���������� ������ "lstInfo" � "lstPersonCode"
            lstInfo.Enabled = True
            lstPersonCode.Enabled = True
            '������� ��������� ��������� ����
            txtPersonCode.Enabled = True
            txtInfo.Enabled = True
            '���������� ����� �� ��������� ���� "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            '������ ��������
        Else
            '������ �������� ������
            frmDemo.BeepSound
             '����� ��� ���������� ����
            txtParole.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "Parole"
            If txtParole.Enabled = True Then txtParole.SetFocus
        End If
            '�������� �������� ������� ����� ������
        tmrParoleTimeOut.Enabled = False
            ' "�������" ���� ������ ���������
        txtParole.Text = ""
    End If

End Sub
