VERSION 5.00
Begin VB.Form frmDataAccessOut 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessOutData"
   ClientHeight    =   3960
   ClientLeft      =   3000
   ClientTop       =   2745
   ClientWidth     =   8715
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
   ScaleHeight     =   3960
   ScaleWidth      =   8715
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.CommandButton cmdOutConst 
      BackColor       =   &H00FF0000&
      Caption         =   "Sant=""50"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOutFree 
      BackColor       =   &H00FF0000&
      Caption         =   "Ls=""000,00"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   16
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   15
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4200
      TabIndex        =   14
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   6480
      Top             =   120
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "*"
      TabIndex        =   10
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   840
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "--"
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
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
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
      TabIndex        =   12
      Top             =   3240
      Width           =   1212
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      LargeChange     =   320
      Left            =   4920
      Max             =   320
      SmallChange     =   320
      TabIndex        =   9
      Top             =   2280
      Width           =   1452
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   6960
      Max             =   99
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Frame fraStatus 
      Caption         =   "????"
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
      Height          =   3375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessOut.frx":24A4
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line3 
         X1              =   1440
         X2              =   0
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   2040
      Picture         =   "frmDataAccessOut.frx":28FE
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   1200
      Picture         =   "frmDataAccessOut.frx":2F30
      Stretch         =   -1  'True
      Top             =   840
      Width           =   735
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   120
      Picture         =   "frmDataAccessOut.frx":3CB6
      Stretch         =   -1  'True
      Top             =   840
      Width           =   495
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   720
      Picture         =   "frmDataAccessOut.frx":44F0
      Stretch         =   -1  'True
      Top             =   840
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
      TabIndex        =   22
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
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
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   7080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   8520
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2280
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   1800
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      Picture         =   "frmDataAccessOut.frx":4B2E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblParole 
      Alignment       =   2  'Center
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
      Left            =   4920
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   1440
      Y2              =   600
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   6240
      X2              =   7080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Image imgAccessOut 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessOut.frx":4F44
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   8520
      X2              =   8520
      Y1              =   1560
      Y2              =   3720
   End
   Begin VB.Label lblLat0 
      Alignment       =   2  'Center
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label lblLat320 
      Alignment       =   2  'Center
      Caption         =   "320"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   8520
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Ls"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5280
      X2              =   6240
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmDataAccessOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������ "��������" ��� �����������
Dim strChecking As String * 8
            '��������� "��������" ��� �����������
Dim strCheckingInfo As String * 8
            '�������� ����� ������ � ��������
Dim lngAccessMoney As Long
            '���������� ���� ���������
Dim intAccessDay As Integer
            '����� ������ ��� ��������� (�����)
Dim intAccessTariffFull As Integer
            '����� ������ ��� ��������� (����)
Dim intAccessTariffDay As Integer
            '����� ������ ��� ��������� (����)
Dim intAccessTariffNight As Integer
            '����� ������ ���� ��������� (���������� ��� ���������)
Dim intAccessTariffHour As Integer
            '����� ����������� (���������� ��� ���������)
Dim intAccessTariff As Integer
            '������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ������� "������� ���������"
Dim intColNum As Integer
            '����, ��������������� ���� �����������
            '  ������� ����������� (��� ���������� ������������ ���)
Dim intDayReg As Integer
            '�����, ��������������� ���� �����������
            '  ������� ����������� (��� ���������� ������������ ���)
Dim intMonthReg As Integer
            '���, ��������������� ���� �����������
            '  ������� ����������� (��� ���������� ������������ ���)
Dim intYearReg As Integer
            '����� ������� ��������� ������� � ������
Dim intPosNum As Integer
             '��������� ������
Dim strPassword As String

            '�������� ������� ���������� ������ "Alt"+ {"--", "E" , "L" � "S"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������� ����� ��������
    If Me.Enabled = True Then
            '������������ "������" ���� �� ������ "--"
        If KeyCode = 189 And Shift = 4 Then
            If cmdOK.Enabled = True Then
                Call cmdOK_Click
                Exit Sub
            End If
            '������������ "������" ���� �� ������ "Cancel"
        ElseIf KeyCode = 69 And Shift = 4 Then
            If cmdCancel.Enabled = True Then
                Call cmdCancel_Click
                Exit Sub
            End If
            '������������ "������" ���� �� ������ "0 Ls"
        ElseIf KeyCode = 76 And Shift = 4 Then
            If cmdOutFree.Visible = True Then
                Call cmdOutFree_Click
                Exit Sub
            End If
            '������������ "������" ���� �� ������ "XX San"
        ElseIf KeyCode = 83 And Shift = 4 Then
            If cmdOutConst.Visible = True Then
                Call cmdOutConst_Click
                Exit Sub
            End If
        End If
    End If
    
End Sub

            '������� � ��������� ��������� (������ "0" Ls)
Private Sub cmdOutFree_Click()
            
            '����������� ������� �� ������ "0 Ls"
    If cmdOutFree.MousePointer = vbNoDrop Then Exit Sub

            '��������� ���������� ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Mid(txtMoneyDate.Text, 11)

End Sub

            '������� � ��������� ��������� (������ "XX" San)
Private Sub cmdOutConst_Click()
            
            '����������� ������� �� ������ "XX San"
    If cmdOutConst.MousePointer = vbNoDrop Then Exit Sub

            '��������� ���������� ���� "����������"
    If Int(gTariffConst / 100) < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gTariffConst / 100) < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gTariffConst / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gTariffConst / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gTariffConst - Int(gTariffConst / 100) * 100 < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(Str(gTariffConst - Int(gTariffConst / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(Str(gTariffConst - Int(gTariffConst / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If

End Sub

            '����������� ������� ������ ��������� - "Document"
Private Sub chkDocument_Click()
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��� ���������� ����
        txtMoneyDate.BackColor = vbCyan
            '������� ��������� ������� �� ������ "OK_-", "0 Ls" � "XX San"
        cmdOK.MousePointer = 0
        cmdOutFree.MousePointer = 0
        cmdOutConst.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub

            '������� � ��������� ��������� (������ "OK_-")
Private Sub cmdOK_Click()
            '������
Dim strStatus As String
            '��� �������� ��� ������������ � "������� ������"
Dim intAutoDeletionCode  As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer

            '����������� ������� �� ������ "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            '���� ���������� ���������
    If imgTime.Visible = True Then
            '�������� ���� "Tag" ����� "frmLease"
        frmLease.Tag = 0
            '������� �� ����� ����� "frmLease" � ������� ����������� 1
        frmLease.Show 1
            '��������� � ���� "txtInfo" ���������� � ������� ���������
        If frmLease.Tag <> "Exit" Then _
        txtInfo.Text = Left(CStr(frmLease.Tag), 4) + Mid(txtInfo.Text, 5)
    End If

            '�� ������� ����� ��������� ������
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            '�������� ���� "Tag" ����� "frmMinus"
        frmMinus.Tag = 0
            '������� �� ����� ����� "frmMinus" � ������� ����������� 1
        frmMinus.Show 1
            '����� �� ������ � �� (����)�������� ������������� ����
        If frmMinus.Tag = "Exit" Then
            '������� � ��������� ���������
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            '������� ������������ ������ "OK _ +" � "Cancel _ Exit"
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
            
            '��������
    If imgHuman.Visible = True Then
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
        gAccessMoneyCell = gAccessMoneyCellHuman
            '����
    ElseIf imgBaby.Visible = True Then
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
        gAccessMoneyCell = gAccessMoneyCellBaby
            '������
    ElseIf imgConvoy.Visible = True Then
            '������� ����� ����������� ��� ������ (��� ��������� ��������)
        gAccessMoneyCell = gAccessMoneyCellConvoy
            '�����
    ElseIf imgFamily.Visible = True Then
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
        gAccessMoneyCell = gAccessMoneyCellFamily
            '�� ��������� ����������
    ElseIf imgTime.Visible = False Then
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
        gAccessMoneyCell = gAccessMoneyCellHuman
    End If
            '��������� ���������� - ���������� �������� ����� � ����
            '  �������� ����������, �� ������� ��� ����� �������� �����
    If imgTime.Visible = True Then Call Prolong(strStatus)
            
            '����� ���������-������� ������������ ������������� ����
    intAutoDeletionCode = frmTablePerson.AutoDelAccess(txtPersonCode.Text, strStatus)
                                                                                             
            '(����)�������� ������������� ���� ��������o -
            '   ���������������� �������
    If intAutoDeletionCode = 0 Then
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '������
        gProtocol.strProtocStatus = strStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        If gAccessDeletion = 1 Then
            '���������� ��������
            gProtocol.strProtocReserve = "AutoDelAcce " + Left(Trim(txtMoneyDate.Text), 9)
        Else
            '���������� ��������
            gProtocol.strProtocReserve = "LogDelAcce " + Left(Trim(txtMoneyDate.Text), 9)
        End If
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtMoneyDate.Tag = 0
            '������� (����)�������� ������������� ����
        Me.Tag = 1
        
            '����� "������ ���������" �����������
        If chkDocument.Value = 1 Then
            '������ ��������� (�������� �� �����-�����, ��������
            '  ���� �/��� ��������� ����)
            Call frmDemo.PrintDocument(gProtocol.strProtocName, _
            gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
            gProtocol.strProtocTime, gProtocol.strProtocDate, _
            gProtocol.strProtocReserve, intError)
        End If
        
        
            '������� �������� ?
            
            '����������� ����� ���������� ������� (������������ ����������
            '   ��������), ����� ���������� (��� ������ ��������� ������������
            '   ���������� �������� ����������), ��������� ������ �����������,
            '   ���������� ������� ����������� �������� � ���������� ������
            '   ��������� ��������� - ������� ��������
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        imgTime.Visible = True And gAccessDeletion = 1 And gTermOut <> -1 Then
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ��������� ���������� ������� �����������
            '  � ����������� ����� "���������� ��������"
            If frmDemo.cmdOpen(gTermOut).Tag = 0 And Me.Tag = 1 And _
            gAccessDeletion = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
                frmDemo.imgAccessInData(gTermOut).Enabled = False
                frmDemo.imgAccessOutData(gTermOut).Enabled = False
                frmDemo.imgAccessInfoData(gTermOut).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(gTermOut))))
                frmDemo.cmdOpen(gTermOut).Tag = vntAddr
                frmDemo.cmdOpen(gTermOut).Caption = "Addr=" + CStr(vntAddr)
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(gTermOut).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(gTermOut).Enabled = True
            '����������� ������� ����������� "������"
                Call frmDemo.OpenBarrier(gTermOut)
            End If
        End If
        
        
            '������� � ��������� ���������
        cmdCancel_Click
            '����� � ������������ ������������� ���� -
            '   ���������������� �������
    Else
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '������
        gProtocol.strProtocStatus = strStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid AutoDelAccess"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)�������� ������������� ����
        Me.Tag = 2
            
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
            
            '������� ���������� ������ "OK _ +" � "Cancel _ Exit"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
            
            '���� �� ����������� ��������� � ��������� ����� ������� �����
    If Me.Tag = 1 And _
    ((txtPersonCode.Tag = 1 And imgMoneyFree.Visible = True) Or _
    (txtPersonCode.Tag = 1 And imgTime.Visible = True) Or _
    (txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1)) Then
            '���� �������� � �������� ��������� "������� ������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
            '������ �������� ������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" -- """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" -- """, intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '������� ������ �� (����)�������� ������������� ����
            Me.Tag = 2
            '����� �� ���������
            Exit Sub
        End If
    End If
    
                '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
        '������� ������ �� (����)�������� ������������� ����
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
            
            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������, �������� ��������� ����)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessPlus
            
            '������� ��������� ������� ���������� ����� "DataAccessOut"
    txtPersonCode.Enabled = True
            '������� ������������ �������� ���������� ����� "DataAccessOut"
    lblParole.Enabled = False
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ���������� ������ ���������� ������ ����������� ��������
    cmdOutFree.Visible = False
    cmdOutConst.Visible = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '������� ���������� ��������������� "������"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '���������� ����� �� ��������� ���� "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK _ -"
    cmdOK.MousePointer = vbNoDrop
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
            '������� ������������ �������� ���������� ����� "DataAccessOut"
    lblParole.Enabled = False
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� �������� ���������� ����� "DataAccessOut"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtParole.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '����� ������ ��� ��������� (�����)
    intAccessTariffFull = gAccessDN
            '����� ������ ��� ��������� (����)
    intAccessTariffDay = gAccessD
            '����� ������ ��� ��������� (����)
    intAccessTariffNight = gAccessN
            '������� ����������� ������� �� ������ "OK _ -"
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
            '�������� ��������� ���� "����������"
    txtInfo.Text = ""
            '�������� ��������� ���� "����������"
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '������� ���������� ��������������� "������"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            '������� ������������ �������� ���������� ����� "DataAccessOut"
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '������� ����������� ������� �� ������ "OK_-"
    cmdOK.MousePointer = vbNoDrop

End Sub

            '��������� ����� � ������� "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            '������ "����������"
Dim strInfo As String
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '����� ���������� ����������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������� ����������
Dim strHour As String
Dim strMinute As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
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
            '�������� ��������� ���� "����������"
            txtInfo.Text = ""
            '�������� ��������� ���� "����������"
            txtMoneyDate.Text = ""
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
            lblMoneyDate.Visible = False
            '������� ���������� "������"
            imgMoneyFree.Visible = False
            imgCalendar.Visible = False
            fraDayNight.Visible = False
            imgTime.Visible = False
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            '������� ������������ �������� ���������� ����� "DataAccessOut"
            lblLat0.Enabled = False
            lblLat320.Enabled = False
            hsbLat.Enabled = False
            '����� ���������-������� ����������
            '������������� ����
            intAutoFindCode = frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
            strInfo, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
            If intAutoFindCode = 0 Then
            '��������� ��������� ���� "����������"
                txtInfo.Text = strInfo
            
            '���������� ������ "��������"
                Call frmTablePerson.UnPack(strDate, strChecking)
            
            '����������� ������������� ������ "��������"
                txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
                Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/�����/������ ����������
                If Mid(Trim(strChecking), 7, 1) = "0" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
                ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
                ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
                txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
                End If

            '��������� ����� ���������� ����������
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
            '���� ���������� ����������
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
                strDate = Trim(gProtocol.strProtocDate)
            
            '������ ������� ���������� �����������
            
            '������������ ��� ����������� ������ ����������
                If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
                Left(Trim(strStatus), 2) <> "09" Then
            '���� �������� � ��������  ������� ���������� - �� �����
                    intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
                    frmDemo.BeepSound
                    If frmDemo.optEnglish = True Then
                        MsgBox "Status Error", intButtonsAndIcons, "Error"
                    Else
                        MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
                    End If
            '�������� �������  ��������� � ��������� ���� "PersonCode"
                    txtPersonCode.Tag = 0
            '������� ����������� ������� �� ������ "OK_-"
                     cmdOK.MousePointer = vbNoDrop
                     GoTo PersonCodeError
                End If
            '���������� ����������
                If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
                    imgMoneyFree.Visible = True
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
                    gAccessMoneyCell = gAccessMoneyCellHuman
            '������������ ������� � ���� "����������"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            '��� ����������� ���������� �������
                    If txtPersonCode.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK_-"
                        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                        cmdOK.SetFocus
                    End If
            '���������� ����������
                ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
                    imgCalendar.Visible = True
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
                    gAccessMoneyCell = gAccessMoneyCellHuman
            '���������� � ������� ������� ���������������
            '  ������� ����������
                    If Right(Trim(strChecking), 1) = "D" Then
            '������� ����� �������
                        optDay.Value = True
                    ElseIf Right(Trim(strChecking), 1) = "N" Then
            '������ ����� �������
                        optNight.Value = True
                    ElseIf Right(Trim(strChecking), 1) <> "D" And _
                    Right(Trim(strChecking), 1) <> "N" Then
            '�������� ����� �������
                        optDayNight.Value = True
                    End If
            '���������� �� ����� ������������ � �����������
                    If Right(Trim(strChecking), 1) <> "E" Then
                        fraDayNight.Visible = True
            '���������� ����� ������������ � �����������
                    Else
                        Exit Sub
                    End If
            '������������ ������� � ���� "����������"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            '��������� ���������� ���� ��������� ��� �� ��������
                    If CInt(Mid(txtMoneyDate.Text, 17, 4)) = CInt(Mid(txtMoneyDate.Text, 42, 4)) And _
                    ((CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
                    (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
                    CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
                    CInt(Mid(txtMoneyDate.Text, 36, 2)))) Or _
                    CInt(Mid(txtMoneyDate.Text, 17, 4)) > CInt(Mid(txtMoneyDate.Text, 42, 4)) Then
            
            '��� ����������� ���������� �������
                        txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
                        txtMoneyDate.BackColor = vbCyan
            '������� ��������� ������� �� ������ "OK_-"
                        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                        cmdOK.SetFocus
                        Exit Sub
                    End If
            '����������� ������ ����������� (���������� ��� ���������)
            '������� ����� �������
                    If optDay.Value = True Then
                        intAccessTariff = intAccessTariffDay
            '������ ����� �������
                    ElseIf optNight.Value = True Then
                        intAccessTariff = intAccessTariffNight
            '�������� ����� �������
                    ElseIf optDayNight.Value = True Then
                        intAccessTariff = intAccessTariffFull
                    End If
            '������� ������� ����� ���������� ���� "txtMoneyDate"
                    lblMoneyDate.Visible = True
            '������� ���������� �������� ���������� ����� "DataAccessOut"
                    lblLat0.Enabled = True
                    lblLat320.Enabled = True
                    hsbLat.Enabled = True
            '��������� ����������
                ElseIf Left(Trim(strStatus), 2) = "09" Then
            '����� ����������� (���������� ��� ���������) = �������
            '  ��������� ������ ����� (�� ������ ������ ����)
                    intAccessTariff = (gAccessHourD + gAccessHourN) / 2 * 24
            '������� ������� ����� ���������� ���� "txtMoneyDate"
                    lblMoneyDate.Visible = True
            '������� ������� ��������������� "������"
                    imgTime.Visible = True
            '��������
                    If Mid(txtInfo.Text, 5, 1) = "1" Then
                        imgHuman.Visible = True
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
                        gAccessMoneyCell = gAccessMoneyCellHuman
            '����
                    ElseIf Mid(txtInfo.Text, 5, 1) = "2" Then
                        imgBaby.Visible = True
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
                        gAccessMoneyCell = gAccessMoneyCellBaby
            '������
                    ElseIf Mid(txtInfo.Text, 5, 1) = "3" Then
                        imgConvoy.Visible = True
            '������� ����� ����������� ��� ������ (��� ��������� ��������)
                        gAccessMoneyCell = gAccessMoneyCellConvoy
            '�����
                    ElseIf Mid(txtInfo.Text, 5, 1) = "4" Then
                        imgFamily.Visible = True
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
                        gAccessMoneyCell = gAccessMoneyCellFamily
                    End If
            '������� ���������� �������� ���������� ����� "DataAccessOut"
                    lblLat0.Enabled = True
                    lblLat320.Enabled = True
                    hsbLat.Enabled = True
            '������������ ������� � ���� "����������"
                    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
                    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
                End If
    
            '�� ���������� ����������
                If Left(Trim(strStatus), 2) <> "10" Then
            '���� ����������� ���������� (��� ���������� ��� ���������)
                    strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            '���������� ����  ����������� ������� �����������
            '  (��� ���������� ��� �������� ��������)
                    intDayReg = Left(strDate, 2)
                    intMonthReg = Mid(strDate, 4, 2)
                    intYearReg = Right(strDate, 4)
            '����� ��� ���������� ����
                    txtMoneyDate.BackColor = vbWhite
            '��������� ����������
                    If Left(Trim(strStatus), 2) = "09" Then
            '����������� ������� "Scroll" - ��������� ��� �������� "Lat"
                        hsbLat.Value = hsbLat.Max
            '�������������e ������ "��������"
'                        strChecking = strCheckingSafe
                    End If
                End If
                Exit Sub
            End If
            
            '������������ ��� � ������������ ��������� ��� ������ ������
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
            '������� ����������� ������� �� ������ "OK_-"
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
            '������� ����������� ������� �� ������ "OK_-"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub

            '��������� ������� "PersonCode" ��� ������������
            '  ���������� ����� ����������� "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            '������ "����������"
Dim strInfo As String
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '����� ���������� ����������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������� ����������
Dim strHour As String
Dim strMinute As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
             '����� ���������� ����������� ������� �����
    Do While Me.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '������� ������������ ��� � ���������������
            '  ��������� ����
    txtPersonCode.Text = Trim(vntPersonCode)
            '������� ����������� ��������� ���� �������������
            '  ���� ���������� ����� "frmDataAccessOut"
    txtPersonCode.Enabled = False
            '������� ��� ���������� ����
    txtPersonCode.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 1
            '������� ���������� ��������������� "������"
    imgHuman.Visible = False
    imgBaby.Visible = False
    imgConvoy.Visible = False
    imgFamily.Visible = False
            '����� ���������-������� ����������
            '������������� ����
    intAutoFindCode = frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
    strInfo, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
    If intAutoFindCode = 0 Then
            '��������� ��������� ���� "����������"
        txtInfo.Text = strInfo
            
            '���������� ������ "��������"
        Call frmTablePerson.UnPack(strDate, strChecking)
            
            '����������� ������������� ������ "��������"
        txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
        Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
        Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/�����/������ ����������
        If Mid(Trim(strChecking), 7, 1) = "0" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
        ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
        ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
        End If

            '��������� ����� ���������� ����������
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
            '���� ���������� ����������
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
        strDate = Trim(gProtocol.strProtocDate)
            
            '������ ������� ���������� �����������
            
            '������������ ��� ����������� ������ ����������
        If Left(Trim(strStatus), 2) <> "10" And Left(Trim(strStatus), 2) <> "08" And _
        Left(Trim(strStatus), 2) <> "09" Then
            '���� �������� � ��������  ������� ���������� - �� �����
            intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
            frmDemo.BeepSound
            If frmDemo.optEnglish = True Then
                MsgBox "Status Error", intButtonsAndIcons, "Error"
            Else
                MsgBox "Nepareizs statuss", intButtonsAndIcons, "Error"
            End If
            '�������� �������  ��������� � ��������� ���� "PersonCode"
            txtPersonCode.Tag = 0
            '������� ����������� ������� �� ������ "OK_-"
            cmdOK.MousePointer = vbNoDrop
            GoTo PersonCodeError
        End If
            '���������� ����������
        If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
            imgMoneyFree.Visible = True
            '������������ ������� � ���� "����������"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            '��� ����������� ���������� �������
            If txtPersonCode.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK_-"
                cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                cmdOK.SetFocus
            End If
            '���������� ����������
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
            imgCalendar.Visible = True
            '���������� � ������� ������� ���������������
            '  ������� ����������
            If Right(Trim(strChecking), 1) = "D" Then
            '������� ����� �������
                optDay.Value = True
            ElseIf Right(Trim(strChecking), 1) = "N" Then
            '������ ����� ������� �
                optNight.Value = True
            ElseIf Right(Trim(strChecking), 1) <> "D" And _
            Right(Trim(strChecking), 1) <> "N" Then
            '�������� ����� �������
                optDayNight.Value = True
            End If
            '���������� �� ����� ������������ � �����������
            If Right(Trim(strChecking), 1) <> "E" Then
                fraDayNight.Visible = True
            '���������� ����� ������������ � �����������
            Else
                Exit Function
            End If
            '������������ ������� � ���� "����������"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            '��������� ���������� ���� ��������� ��� �� ��������
            If (CInt(Mid(txtMoneyDate.Text, 14, 2)) > CInt(Mid(txtMoneyDate.Text, 39, 2))) Or _
            (CInt(Mid(txtMoneyDate.Text, 14, 2)) = _
            CInt(Mid(txtMoneyDate.Text, 39, 2)) And CInt(Mid(txtMoneyDate.Text, 11, 2)) >= _
            CInt(Mid(txtMoneyDate.Text, 36, 2))) Then
            '��� ����������� ���������� �������
                txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
                txtMoneyDate.BackColor = vbCyan
            '������� ��������� ������� �� ������ "OK_-"
                cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                cmdOK.SetFocus
                Exit Function
            End If
            '����������� ������ ����������� (���������� ��� ���������)
            '������� ����� �������
            If optDay.Value = True Then
                intAccessTariff = intAccessTariffDay
            '������ ����� �������
            ElseIf optNight.Value = True Then
                intAccessTariff = intAccessTariffNight
            '�������� ����� �������
            ElseIf optDayNight.Value = True Then
                intAccessTariff = intAccessTariffFull
            End If
            '������� ������� ����� ���������� ���� "txtMoneyDate"
            lblMoneyDate.Visible = True
            '������� ���������� �������� ���������� ����� "DataAccessOut"
            lblLat0.Enabled = True
            lblLat320.Enabled = True
            hsbLat.Enabled = True
            '��������� ����������
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            '����� ����������� (���������� ��� ���������) = �������
            '  ��������� ������ ����� (�� ������ ������ ����)
            intAccessTariff = (gAccessHourD + gAccessHourN) / 2 * 24
            '������� ������� ����� ���������� ���� "txtMoneyDate"
            lblMoneyDate.Visible = True
            '������� ������� ��������������� "������"
            imgTime.Visible = True
            '��������
            If Mid(txtInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            '����
            ElseIf Mid(txtInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            '������
            ElseIf Mid(txtInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            '�����
            ElseIf Mid(txtInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
            '������� ���������� �������� ���������� ����� "DataAccessOut"
            lblLat0.Enabled = True
            lblLat320.Enabled = True
            hsbLat.Enabled = True
            '������������ ������� � ���� "����������"
            txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
            Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
        End If
    
            '�� ���������� ����������
        If Left(Trim(strStatus), 2) <> "10" Then
            '���� ����������� ���������� (��� ���������� ��� ���������)
            strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            '���������� ����  ����������� ������� �����������
            '  (��� ���������� ��� �������� ��������)
            intDayReg = Left(strDate, 2)
            intMonthReg = Mid(strDate, 4, 2)
            intYearReg = Right(strDate, 4)
            '����� ��� ���������� ����
            txtMoneyDate.BackColor = vbWhite
            '����������� ������� "Scroll" - ��������� ��� �������� "Lat"
            hsbLat.Value = hsbLat.Max
        End If
        Exit Function
    End If
            
            '������������ ��� � ������������ ��������� ��� ������ ������
PersonCodeError:
            '������ �������� ������
    frmDemo.BeepSound
    txtPersonCode.Text = "Error"
            '�������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 0
            '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK_-"
    cmdOK.MousePointer = vbNoDrop
            '���������� ����� �� ��������� ���� "PersonCode"
    If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus

End Function

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
            ' "��������" �������� "������"
    lblParole.Enabled = False
            '�������� �������� ������� ����� ������
    tmrParoleTimeOut.Enabled = False
            '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
            '������� ���������� ������ "OK" � "Cancel"
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    
End Sub

            '��������� ��������� "������ ����" �� ���� ������
Private Sub txtParole_Click()
            '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
            '������� ������������ ������ "OK" � "Cancel"
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
            ' "��������" �������� "������"
    lblParole.Enabled = True
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
            '������ ������� � "���������" �������� "������"
    If KeyAscii = vbKeyReturn And lblParole.Enabled = True Then
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
            '������� ��������� ����� "Document"
            imgDocument.Enabled = True
            chkDocument.Enabled = True
            '������� �������� ������ ���������� ������ ����������� ��������
            cmdOutFree.Visible = True
            cmdOutConst.Visible = True
            '������ ��������
        Else
            '������ �������� ������
            frmDemo.BeepSound
            '������� ����������� ����� "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
            '������� ���������� ������ ���������� ������ ����������� ��������
            cmdOutFree.Visible = False
            cmdOutConst.Visible = False
             '����� ��� ���������� ����
            txtParole.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "Parole"
            If txtParole.Enabled = True Then txtParole.SetFocus
        End If
            '�������� �������� ������� ����� ������
        tmrParoleTimeOut.Enabled = False
            ' "�������" ���� ������ ���������
        txtParole.Text = ""
            ' "��������" �������� "������"
        lblParole.Enabled = False
            '������� ���������� ������ "OK" � "Cancel"
        cmdOK.Enabled = True
        cmdCancel.Enabled = True
    End If

End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Change()
            '����� ����������� ����������
Dim strHourReg As String
Dim strMinuteReg As String
Dim lngTimeReg As Long
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '����� ���������� ����������
Dim intHour As Integer
Dim intMinute As Integer
Dim lngTimeDel As Long
            '��������������� (�� ��� �����) ����� ���������� ����������
Dim strHour As String
Dim strMinute As String
            '��������������� ���������� (��� ����� �����) �� ���� "����������"
Dim strMoneyDate As String
            '������� ��������� ����� ������� �� ����������� (� �������)
Dim lngAccessTimeD As Long
            '������ ��������� ����� ������� �� ����������� (� �������)
Dim lngAccessTimeN As Long
            
            '������� - ����� ���������
    If hsbLat.Tag <> 0 And hsbLat.Value = 0 And _
    hsbSant.Tag <> 0 Then hsbSant.Value = 0
            '��������� ������� ��������� ���������
    hsbLat.Tag = hsbLat.Value
    hsbSant.Tag = hsbSant.Value
            '�� ������ ������������ ��� - ����� �� ���������
    If txtPersonCode.Tag = 0 Then Exit Sub
            
            '���������� ������ "��������"
    Call frmTablePerson.UnPack(strDate, strChecking)
            
            '����������� ������������� ������ "��������"
    txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
    Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
    Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/�����/������ ����������
    If Mid(Trim(strChecking), 7, 1) = "0" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
    ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
    ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
    End If
            '��������� ����� ���������� ����������
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
            '���� ���������� ����������
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Trim(gProtocol.strProtocDate)
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text) + "  ==>  " + _
    Trim(strDate) + "/" + Trim(strHour) + ":" + Trim(strMinute)
            
            '��������� ���������� ���� "����������"
    If hsbLat.Value < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    ElseIf hsbLat.Value < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    ElseIf hsbLat.Value > 99 Then
        txtMoneyDate.Text = Trim(Str(hsbLat.Value)) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If hsbSant.Value < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(hsbSant.Value)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(hsbSant.Value)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            '�������� �������  ��������� ����������
    txtMoneyDate.Tag = 0
           '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK_-"
    cmdOK.MousePointer = vbNoDrop
            '�� ������� ��������� ������ �� ��������� ����� ���������
    If hsbLat.Value > 0 Or hsbSant.Value > 0 Then
            '���������� �������  ��������� ����������
        txtMoneyDate.Tag = 1
            '�������� ����� ������� � ��������
        lngAccessMoney = hsbLat.Value * 100 + hsbSant.Value
            '���������� ������������� ���� ���������
        intAccessDay = Int(lngAccessMoney / intAccessTariff)
            '�������������� ��������� �������� "���������"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            '���� �� ���� "���������" (�� ����������
            '  ��� ��������� ��� ���� ����������� �������)
        For intAccessDay = intAccessDay To 1 Step -1
            '���������� ������������� ���� ��������� ���������
            If frmTableCalendar.comCalendar.Day = _
            Left(strDate, 2) And _
            frmTableCalendar.comCalendar.Month = _
            Mid(strDate, 4, 2) And _
            frmTableCalendar.comCalendar.Year = _
            Right(strDate, 4) Then GoTo EndCycle
            
            '������ �����, ������ � ���� � ���� "����������"
            If frmTableCalendar.comCalendar.Month > 9 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year)) + _
                Right(txtMoneyDate.Text, 31)
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + ".0" + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year)) + _
                Right(txtMoneyDate.Text, 31)
            End If
            '����������� "���������" �� ���� ���� ������
            frmTableCalendar.comCalendar.NextDay
            
        Next
    End If
EndCycle:
           '������������� ������
    If (frmTableCalendar.comCalendar.Day <> Left(strDate, 2) Or _
    frmTableCalendar.comCalendar.Month <> Mid(strDate, 4, 2) Or _
    frmTableCalendar.comCalendar.Year <> Right(strDate, 4)) And _
       (imgCalendar.Visible = True And imgTime.Visible = False Or _
        hsbLat.Value = 320) Or _
           (hsbLat.Value = 0 And hsbSant.Value = 0) Then
          '������ ��������� ����������
        txtMoneyDate.Tag = 0
           '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK_-", "0 Ls" � "XX San"
        cmdOK.MousePointer = vbNoDrop
        cmdOutFree.MousePointer = vbNoDrop
        cmdOutConst.MousePointer = vbNoDrop
    End If
            '��������� (�������� ������� ������ �� �������� ���)
    If intAccessDay > 0 And (hsbLat.Value <> 0 Or hsbSant.Value <> 0) Then
            '���������� ���������� (��� ���������) ���� ���������
        intAccessDay = Int(lngAccessMoney / intAccessTariff) - intAccessDay
           '�������������� ���������� (��� ���������) ����� ������� � ��������
        lngAccessMoney = intAccessDay * intAccessTariff
            
            '��������� ���������� - ���������� ��������� ����� � ���� "����������"
        If imgTime.Visible = True Then
            '����� ���������� ���������� (� �������)
            lngTimeDel = intHour * 60 + intMinute
            '��������������� ���������� (��� ����� �����) �� ���� "����������"
            If Mid(Trim(txtMoneyDate.Text), 12, 1) = "." Then
                strMoneyDate = Left(Trim(txtMoneyDate.Text), 10) + "0" + _
                Trim(Mid(Trim(txtMoneyDate.Text), 11))
            Else
                strMoneyDate = Trim(txtMoneyDate.Text)
            End If
            '����� ����������� ���������� (� �������)
            strHourReg = Mid(Trim(strMoneyDate), 22, 2)
            strMinuteReg = Mid(Trim(strMoneyDate), 25, 2)
            lngTimeReg = CInt(strHourReg) * 60 + CInt(strMinuteReg)
           
           '��������� ����� ������� � ��������
            
            
            '��������� ����� "?"-� ����� - ������� ����� = 0
            '   (�������� ��� ����� � ������� ���� ��� ���������� ����)
            If (lngTimeDel - lngTimeReg) <= 0 And _
            intAccessDay = 0 And (Left(strDate, 2) = Mid(Trim(strMoneyDate), 11, 2) Or _
            CInt(Left(strDate, 2)) < CInt(Mid(Trim(strMoneyDate), 11, 2)) Or _
            CInt(Left(strDate, 2)) > CInt(Mid(Trim(strMoneyDate), 11, 2)) And _
            (CInt(Mid(strDate, 4, 2)) < CInt(Mid(Trim(strMoneyDate), 14, 2)) Or _
            CInt(Mid(strDate, 7, 4)) < CInt(Mid(Trim(strMoneyDate), 17, 4)))) Then
                lngAccessMoney = 0
            '��������� ����� "?"-� ����� - ������� ����� = ����� + . . .
            Else
            '������� ��������� ����� ������� �� ����������� (� �������)
                lngAccessTimeD = CInt(Left(Trim(gAccessTimeD), 2)) * 60 + _
                CInt(Mid(Trim(gAccessTimeD), 4, 2))
            '������ ��������� ����� ������� �� ����������� (� �������)
                lngAccessTimeN = CInt(Mid(Trim(gAccessTimeD), 7, 2)) * 60 + _
                CInt(Right(Trim(gAccessTimeD), 2))
                lngAccessMoney = lngAccessMoney - intAccessTariff
            '���������� ����� �� ����� ������� ����� (����� 00.00 �����)
                If Mid(strMoneyDate, 11, 2) = Mid(strMoneyDate, 36, 2) And _
                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2) Then
            'C���a ������ ����� (����� �� ��������� ��� ������)
                    lngAccessMoney = gAccessMoneyCell
            
'���������� ��� ���������� "ICE HALL"
                        
'�������� ������ - �� ����� �������� ��������� �������
If lngTimeReg >= lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
    lngAccessMoney = lngAccessMoney + _
    Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * gAccessMoneyCell

'''            '�������� ������ - �� ����� �������� ��������� �������
'''                    If lngTimeReg >= lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
'''                        lngAccessMoney = lngAccessMoney + _
'''                        Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * _
'''                        gAccessTimeCell * gAccessHourD / 60
            '�������� ������ - �� ����� ������� ��������� �������
                    ElseIf lngTimeReg > lngAccessTimeN And lngTimeDel <= 24 * 60 Or _
                    lngTimeReg >= 0 And lngTimeDel < lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            '�������� ������ - ��������� (�������� �� ����� ��������, � ��������
            '   �� ����� ������� ���������� �������
                    ElseIf lngTimeReg >= 0 And lngTimeDel <= 24 * 60 Then
            '�������� ������ - ����/����
                        If lngTimeReg < lngAccessTimeD And lngTimeDel <= lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngTimeDel - lngAccessTimeD) / gAccessTimeCell) _
                            * gAccessTimeCell * gAccessHourD / 60 + _
                            Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60
            '�������� ������ - ����/����/����
                        ElseIf lngTimeReg < lngAccessTimeD And lngTimeDel > lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60 + _
                            Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourD / 60 + _
                            Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60
            '�������� ������ - ����/����
                        ElseIf lngTimeReg >= lngAccessTimeD And lngTimeDel > lngAccessTimeN Then
                            lngAccessMoney = lngAccessMoney + _
                            Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourN / 60 + _
                            Int((lngAccessTimeN - lngTimeReg) / gAccessTimeCell) * _
                            gAccessTimeCell * gAccessHourD / 60
                        End If
                    End If
            
            
            '���������� ����� � ���������� ����� (�� 00.00 �����)
                ElseIf (CInt(Mid(strMoneyDate, 36, 2)) - CInt(Mid(strMoneyDate, 11, 2)) = 1 And _
                Mid(strMoneyDate, 14, 2) = Mid(strMoneyDate, 39, 2)) Or _
                (Mid(strMoneyDate, 36, 2) = "01" And _
                CInt(Mid(strMoneyDate, 39, 2)) - CInt(Mid(strMoneyDate, 14, 2)) = 1) Then
            '�������� ������ � ���������� ����� - ����
                    If lngTimeReg >= lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((24 * 60 - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            '�������� ������ � ���������� ����� - ����/����
                    ElseIf lngTimeReg >= lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngAccessTimeN - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int((24 * 60 - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            '�������� ������ � ���������� ����� - ����/����/����
                    ElseIf lngTimeReg < lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngAccessTimeD - lngTimeReg) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60 + _
                        Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int((24 * 60 - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
                    End If
            '�������� ������ � ������� ����� - ����
                    If lngTimeDel <= lngAccessTimeD Then
                        lngAccessMoney = lngAccessMoney + _
                        Int(lngTimeDel / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            '�������� ������ � ������� ����� - ����/����
                    ElseIf lngTimeDel <= lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int(lngAccessTimeD / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
            '�������� ������ � ������� ����� - ����/����/����
                    ElseIf lngTimeDel > lngAccessTimeN Then
                        lngAccessMoney = lngAccessMoney + _
                        Int((lngTimeDel - lngAccessTimeN) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60 + _
                        Int((lngAccessTimeN - lngAccessTimeD) / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourD / 60 + _
                        Int(lngAccessTimeD / gAccessTimeCell) * _
                        gAccessTimeCell * gAccessHourN / 60
                    End If
                End If
                
            End If
                
        End If
        
            '�������������� ����������� ��������� ���������
        hsbSant.Value = lngAccessMoney - Int(lngAccessMoney / 100) * 100
        hsbLat.Value = Int(lngAccessMoney / 100)
        hsbLat_Change
    End If
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Or _
    txtPersonCode.Tag = 1 And _
    Int((lngTimeDel - lngTimeReg) / gAccessTimeCell) = 0 And _
    intAccessDay = 0 Then
            '���������� �������  ��������� ����������
        txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
        txtMoneyDate.BackColor = vbCyan
            '������� ����������� ������ ���������
        hsbLat.Enabled = False
            '������� ��������� ������� �� ������ "OK_-", "0 Ls" � "XX San"
        cmdOK.MousePointer = 0
        cmdOutFree.MousePointer = 0
        cmdOutConst.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If
    
End Sub

            '�������� ����� � ���� �������� ��� ���������� ������� ��� ����� � ���� ������
            '  ��� ����������� ������� (�� ������������ � ������������ ������� ������������
            '  ����������), �� ������� ��� ����� �������� �����
Private Sub Prolong(ByRef strStatus As String)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            '���� (� �����) �������� �������
Dim strDate As String
            '����� �������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� �������� �������
Dim strHour As String
Dim strMinute As String
            '������� ����������� \ 0 - ����� \ 1 - ������ \ 2 - ���������������
Dim strPresent As String * 1
            '������� ("�" - ������������ �����; "D" - ������� ����� �������;
            '  "N" - ������ ����� �������; "������ ������"   - �������� �����
            '  �������)
Dim strExpander As String * 1
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ������������� � "������� ������"
Dim intAutoCorrectionCode  As Integer
            '���������� ����� �������, � ������� �������� �����������
            '  ����������� ������� ���������� ���������� �� �����������
Dim intCellLimit As Integer

            '��������� ����� �������� �������
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
            '���� �������� �������
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
            '������� ����������� �������
    strPresent = "2"
            '������� �������
    strExpander = "P"

            '���� ��� ���������� ������ ����������� ����� ������� ������������ ����������
            '  �� �����������
    If gAcceTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
            '���������� ����� �������, � ������� �������� �����������
            '  ����������� ������� ���������� ���������� �� �����������
'        intCellLimit = gAccessCellLimit
    Else
        intCellLimit = 0
    End If
    
            '��������� "���������" ����� � ���� �����������
            '  �������, �� ������� ��� ����� �������� ���������� �����
        
            '��������� ������� ����
    If (intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell) > 59 Then
        If (gAccessTimeCell * intCellLimit + gAccessTimeCell) > 1440 Then
            intHour = intHour + Int((intMinute + _
            gAccessTimeCell * intCellLimit) / 60)
            intMinute = intMinute + gAccessTimeCell * intCellLimit - _
            Int((intMinute + gAccessTimeCell * intCellLimit) / 60) * 60
        Else
            intHour = intHour + Int((intMinute + gAccessTimeCell + _
            gAccessTimeCell * intCellLimit) / 60)
            intMinute = intMinute + gAccessTimeCell + _
            gAccessTimeCell * intCellLimit - _
            Int((intMinute + gAccessTimeCell + _
            gAccessTimeCell * intCellLimit) / 60) * 60
        End If
            
            '��������� ������� ����
        If intHour >= 24 Then
            intHour = intHour - 24
            '��������� "���������" �� ����, ��������� �� �������
            frmTableCalendar.comCalendar.Today
            frmTableCalendar.comCalendar.NextDay
            '���������  �����
            If frmTableCalendar.comCalendar.Day > 9 Then
                strDate = Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                Right(strDate, 6)
            Else
                strDate = "0" + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + _
                Right(strDate, 6)
            End If
            '���������  ������ �, ��������, ����
            If frmTableCalendar.comCalendar.Day = 1 Then
                If frmTableCalendar.comCalendar.Month > 9 Then
                    strDate = "01" + _
                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                    Right(strDate, 4)
                Else
                    strDate = "010" + _
                    Trim(Str(frmTableCalendar.comCalendar.Month)) + _
                    Right(strDate, 4)
                End If
            End If
        End If
            
            '�� ��������� ������� ����
    Else
        intMinute = intMinute + gAccessTimeCell * intCellLimit + _
        gAccessTimeCell
    End If
            
            '"���������" ����
    If intHour < 10 Then
        strHour = "0" + Trim(Str(intHour))
    Else
        strHour = Trim(Str(intHour))
    End If
            '"���������" ������
    If intMinute < 10 Then
        strMinute = "0" + Trim(Str(intMinute))
    Else
        strMinute = Trim(Str(intMinute))
    End If
    
            '���� ��� ���������� ������ ����������� ����� ������� ������������ ����������
            '  �� �����������
    If gAcceTimeLimit > 0 And imgTime.Visible = True And imgCalendar.Visible = True Then
        strCheckingInfo = ""
            '������������ ����������� ��������� "��������"
        For intCount = 1 To 7 Step 2
            '����
            strCheckingInfo = Trim(strCheckingInfo) + _
            Chr(CByte(CInt(Mid(strDate, intCount, 2))))
        Next
            '����
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strHour, 1, 2))))
            '������
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            '�������� ��������� "��������"
        Call frmTablePerson.Pack(strCheckingInfo)
            
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Extra Payment ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Papildus apmaksa ?", intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "��"
        If strResponse = vbYes Then
            '��������� ���� "txtInfo"
            txtInfo = Left(strCheckingInfo, 6) + Trim(Mid(txtInfo, 7))
        Else
            txtMoneyDate.Text = "000,00 Ls" + Mid(Trim(txtMoneyDate.Text), 10)
            Exit Sub
        End If
            
            '��� ��������� ������ � ���������� ��� ����� �� ����� ����
    ElseIf imgTime.Visible = True And imgCalendar.Visible = False And gAcceInpCellNumb > 0 Then
        strChecking = ""
            '������������ ����������� ������ "��������"
        For intCount = 1 To 7 Step 2
            '����
            strChecking = Trim(strChecking) + _
            Chr(CByte(CInt(Mid(strDate, intCount, 2))))
        Next
            '����
        strChecking = Trim(strChecking) + _
        Chr(CByte(CInt(Mid(strHour, 1, 2))))
            '������
        strChecking = Trim(strChecking) + _
        Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            '�������� ������ "��������"
        Call frmTablePerson.Pack(strChecking)
            
            '������� ����������� ���������� � ������ ��� ����������
        strChecking = Left(strChecking, 6) + strPresent + strExpander
    
    End If


            '����� ���������-������� ������������� ��� �������
            '������������� ����
    intAutoCorrectionCode = frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
    txtInfo.Text, strChecking, strStatus)
            
            '(����)��������� ��� �������� ������������� ���� ��������� -
            '   ���������������� �������
    If intAutoCorrectionCode = 0 Then
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '������
        gProtocol.strProtocStatus = strStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "AutoCorAcce"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
    End If

End Sub
