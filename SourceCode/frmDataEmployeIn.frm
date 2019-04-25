VERSION 5.00
Begin VB.Form frmDataEmployeIn 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmployeInData"
   ClientHeight    =   3735
   ClientLeft      =   9120
   ClientTop       =   2925
   ClientWidth     =   2580
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
   ScaleHeight     =   3735
   ScaleWidth      =   2580
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   480
      Top             =   600
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   6
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   240
      Width           =   972
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
      TabIndex        =   3
      Top             =   3000
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   1200
      Width           =   1695
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
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgEmployeIn 
      Height          =   615
      Left            =   1800
      Picture         =   "frmDataEmployeIn.frx":0000
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   615
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
      TabIndex        =   5
      Top             =   1680
      Width           =   495
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
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmDataEmployeIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             '��������� ������
Dim strPassword As String

            '������� � ��������� ��������� (������ "OK _ +")
Private Sub cmdOK_Click()
            '��� �������� ��� ��������������� � "������� ������"
Dim intAutoRegistrCode  As Integer

            '����������� ������� �� ������ "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            '����� ���������-������� ���������������
            '������������� ����
    intAutoRegistrCode = frmTablePerson.AutoRegEmploye(txtPersonCode.Text, _
    txtInfo.Text)
            
            '(����)����������� ������������� ���� ��������� -
            '   ���������������� �������
    If intAutoRegistrCode = 0 Then
            '��������� (�����) ������ "������� ������"
        gTablePerson.Row = gTablePerson.Rows - 1
            '������� ������� "������� ������" = 0 (������� ��� ��������)
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
            '����������
        gProtocol.strProtocReserve = "AutoRegistration"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
            '������� (����)����������� ������������� ����
        frmDataEmployeIn.Tag = 1
            '������� � ��������� ���������
        cmdCancel_Click
            '����� � ��������������� ������������� ���� -
            '   ���������������� �������
    Else
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '������������ ����� (������)
        gProtocol.strProtocStatus = gDefaultStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid AutoRegEmploye"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ �� (����)����������� ������������� ����
        frmDataEmployeIn.Tag = 2
    
            '������� � ��������� ���������
        cmdCancel_Click
    End If
            
End Sub
            
            '������� � ��������� ��������� (������ "Cancel _ Exit")
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            
            '���� �� ����������� ��������� � ��������� ����� ������� �����
    If frmDataEmployeIn.Tag = 1 And txtPersonCode.Tag = 1 _
    And txtInfo.Tag = 1 Then
            '���� �������� � �������� ��������� "������� ������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
            '������ �������� ������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" + """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" + """, intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '����� �� ���������
            Exit Sub
        End If
    End If
    
            '������� ������ �� (����)����������� ������������� ����
    If frmDataEmployeIn.Tag = 0 Then frmDataEmployeIn.Tag = 2
            '������� ��������� ������� �����
    frmDataEmployeIn.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '����������� ������� �����
Private Sub Form_Activate()

            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������, �������� ��������� ����)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessPlus
            
            '������� ����������� ��������� ���� "Info"
    txtInfo.Enabled = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            '������� ��������� ��������� ���� "PersonCode"
    txtPersonCode.Enabled = True
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
           
            '���������� ����� �� ��������� ���� "PersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
           '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop
             '���������� ���� ���������� ����������� ������� �����
    frmDataEmployeIn.Tag = 1

End Sub

            '������������� ������� �����
Private Sub Form_Deactivate()
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub
            
            '�������� ������� �����
Private Sub Form_Load()
            
            '������� ������������ ��������� ���� "PersonCode" � "Info"
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '��������� ��������� "������ ����" �� ���� ������������� ����
Private Sub txtPersonCode_Click()
            
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
            '�������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub

            '��������� ����� � ������� "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
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
            '���������� �������  ��������� � ��������� ���� "PersonCode"
            txtPersonCode.Tag = 1
            '���������� ����� �� ��������� ���� "txtInfo"
            If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
            txtInfo.SetFocus
            '����������� "PersonCode"� ���� "Info" � ����������
            txtInfo = "*" + Right(Trim(txtPersonCode), 14) + " "
            '������� ��� ���������� ����
            txtInfo.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
            txtInfo.Tag = 1
            '��� ����������� ���������� �������
            If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
                 cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                If cmdOK.Visible = True Then cmdOK.SetFocus
            End If
            
            Exit Sub
            '������������ ��� � ������������ ���������
PersonCodeError:
            '������ �������� ������
            frmDemo.BeepSound
            txtPersonCode.Text = "Error"
            '�������� �������  ��������� � ��������� ���� "PersonCode"
            txtPersonCode.Tag = 0
            '����� ��� ���������� ����
            txtPersonCode.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        Else
            '������ �������� ������
            frmDemo.BeepSound
            txtPersonCode.Text = "Error"
            '�������� �������  ��������� � ��������� ���� "PersonCode"
            txtPersonCode.Tag = 0
            '����� ��� ���������� ����
            txtPersonCode.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub

            '��������� ������� "PersonCode" ��� ��������������� ���������
            '  ����� ����������� "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
             '����� ���������� ����������� ������� �����
    Do While frmDataEmployeIn.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '������� ������������ ��� � ���������������
            '  ��������� ����
    txtPersonCode.Text = Trim(vntPersonCode)
            '������� ����������� ��������� ���� "PersonCode"
    txtPersonCode.Enabled = False
            '������� ��������� ��������� ���� "Info"
    txtInfo.Enabled = True
            '������� ��� ���������� ���� "PersonCode"
    txtPersonCode.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 1
            '���������� ����� �� ��������� ���� "Info"
    If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
    txtInfo.SetFocus
    
End Function
            '��������� ��������� "������ ����" �� ���� ����������
Private Sub txtInfo_Click()
            '����� ��� ���������� ����
    txtInfo.BackColor = vbWhite
            '�������� �������  ��������� � ��������� ���� "Info"
    txtInfo.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop

End Sub
            '��������� ����� � ������� ���������� ���� "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '������� ��� ���������� ����
        txtInfo.BackColor = vbCyan
            '���������� � ���������� ���������
        If Len(Trim(txtInfo.Text)) < 16 And Len(Trim(txtInfo.Text)) > 0 Then
            '���������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 1
            '��� ����������� ���������� �������
            If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
                 cmdOK.MousePointer = 0
            '���������� ����� �� ������ "�� _+"
                If cmdOK.Visible = True Then cmdOK.SetFocus
            End If
            Exit Sub
            '��� � ������������ ���������
        Else
            '������ �������� ������
            frmDemo.BeepSound
            txtInfo.Text = "Error"
            '�������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 0
            '����� ��� ���������� ����
            txtInfo.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "Info"
            If txtInfo.Enabled = True And frmDataEmployeIn.Visible = True Then _
            txtInfo.SetFocus
            '������� ����������� ������� �� ������ "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
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
            '� (����)����������� ��������
    frmDataEmployeIn.Tag = 2
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
            '������� ������������ ��������� ���� "PersonCode" � "Info"
    txtPersonCode.Enabled = False
    txtInfo.Enabled = False
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
            '������� ���������� ��������� ���� "PersonCode" � "Info"
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
