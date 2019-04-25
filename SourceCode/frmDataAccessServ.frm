VERSION 5.00
Begin VB.Form frmDataAccessServ 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessServData"
   ClientHeight    =   3105
   ClientLeft      =   4665
   ClientTop       =   2745
   ClientWidth     =   7080
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
   ScaleHeight     =   3105
   ScaleWidth      =   7080
   Tag             =   "0"
   Visible         =   0   'False
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
      Height          =   2055
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optNight 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDay 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   186
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessServ.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   99
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   320
      TabIndex        =   6
      Top             =   960
      Width           =   1452
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
      TabIndex        =   5
      Top             =   2400
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF0000&
      Caption         =   "OK"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkDocument 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4200
      TabIndex        =   2
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6960
      Y1              =   2160
      Y2              =   2160
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
      TabIndex        =   19
      Top             =   960
      Width           =   375
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
      TabIndex        =   18
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   600
      Y2              =   2160
   End
   Begin VB.Image imgAccessServ 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessServ.frx":0802
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image imgDocument 
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessServ.frx":0BC4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   1440
      Y2              =   2400
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   2400
      Y2              =   2400
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
      TabIndex        =   17
      Top             =   1080
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
      TabIndex        =   16
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDataAccessServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������ "��������" ��� �����������
Dim strChecking As String * 8
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '�������� ����� ������ � ��������
Dim intAccessMoney As Integer
            '���������� ���� �������� ��������
Dim intAccessDay As Integer
            '����� ������ ��� (�����)
Dim intAccessTariffFull As Integer
            '����� ������ ��� (����)
Dim intAccessTariffDay As Integer
            '����� ������ ��� (����)
Dim intAccessTariffNight As Integer
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

            '�������� ������� ���������� ������ "Alt"+ {"OK" � "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������� ����� ��������
    If Me.Enabled = True Then
            '������������ "������" ���� �� ������ "OK"
        If KeyCode = 79 And Shift = 4 Then
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
        End If
    End If
    
End Sub
            '����������� ������� ������ ��������� - "Document"
Private Sub chkDocument_Click()
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��� ���������� ����
        txtMoneyDate.BackColor = vbCyan
            '������� ��������� ������� �� ������ "OK_+"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub
            '������� � ��������� ��������� (������ "OK")
Private Sub cmdOK_Click()
            '������
Dim strStatus As String
            '���� ���������� ��� (� �����)
            '  ��������� ���������� � ����������
Dim strDate As String
            '����� ����������� ����������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ����������� ����������
Dim strHour As String
Dim strMinute As String
            '������� ����������� ���������� \ 0-����� \1-����� \ 2-�����������
Dim strCarPresent As String * 1
            '������� ���������� ("�" - ������������ �����; "D" - ������� �����;
            '  "N" - ������ �����; "������ ������ - �������� �����)
Dim strExpander As String * 1
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ������������� � "������� ������"
Dim intAutoCorrectionCode  As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer

            '����������� ������� �� ������ "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            '�� ������� ����� ��������� ������
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            '�������� ���� "Tag" ����� "frmMinus"
        frmMinus.Tag = 0
            '������� �� ����� ����� "frmMinus" � ������� ����������� 1
        frmMinus.Show 1
            '����� �� ������ � �� (����)��������� ������������� ����
        If frmMinus.Tag = "Exit" Then
            '������� � ��������� ���������
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            '������� ������ "��������"
    strChecking = ""
            '��������� ����� ��������� ���������� � ����������
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
            
            '������� ��������������� ����������
    strCarPresent = "2"
            '������� ����������
    strExpander = "P"
            '���������� � ������� ������� �������
    If optDay.Value = True Then
        strExpander = "D"
            '���������� � ������ ������� �������
    ElseIf optNight.Value = True Then
        strExpander = "N"
    End If
    
            '���� ���������� ����������� ���
    strDate = Mid(Trim(txtMoneyDate.Text), 11)
    If Mid(Trim(strDate), 3, 1) = "." Then
        strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
        Mid(Trim(strDate), 7, 4)
    Else
        strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
        Mid(Trim(strDate), 6, 4)
    End If
    
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
            
            '�������� ���������� � ��� ����������� �� �����������
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
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
        gProtocol.strProtocReserve = "AutoCorAcce " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            '������� (����)��������� ��� ������� ������������� ����
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
        
            '������� � ��������� ���������
        cmdCancel_Click
            '����� �� ������������� ��� ������� ������������� ���� -
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
        gProtocol.strProtocReserve = "Invalid AutoCorAccess"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '����� ��� ���������� ����
        txtInfo.BackColor = vbWhite
        txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)��������� ��� ������� ������������� ����
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
            
            '���� �� ����������� ��������� � ��������� ����� ������� �����
    If txtPersonCode.Tag = 1 And _
    (txtInfo.Tag = 1 Or txtMoneyDate.Tag = 1) Then
            '���� �������� � �������� ��������� "������� ������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
            '������ �������� ������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Ignore  "" OK """, intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Ignor.  "" OK """, intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '����� �� ���������
            Exit Sub
        End If
    End If
    
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
            
            '������� ������������ �������� ���������� ����� "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� "������"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtMoneyDate.BackColor = vbWhite
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '���������� ����� �� ��������� ���� "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK"
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
            '������� ������������ �������� ���������� ����� "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� �������� ���������� ����� "DataAccessServ"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtMoneyDate.Text = ""
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '����� ������ ��� (�����)
    intAccessTariffFull = gAccessDN
            '����� ������ ��� (����)
    intAccessTariffDay = gAccessD
            '����� ������ ��� (����)
    intAccessTariffNight = gAccessN
            '������� ����������� ������� �� ������ "OK"
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
    txtInfo.BackColor = vbWhite
    txtMoneyDate.BackColor = vbWhite
            '�������� ��������� ���� "����������"
    txtInfo.Text = ""
            '�������� ��������� ���� "����������"
    txtMoneyDate.Text = ""
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� "������"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            '������� ������������ �������� ���������� ����� "DataAccessServ"
    lblInfo.Enabled = False
    txtInfo.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '������� ����������� ������� �� ������ "OK"
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
            '����� ��������� ���������� � ����������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������
            '  ���������� � ����������
Dim strHour As String
Dim strMinute As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            '������� ����� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� "Z_������"
Dim lngZ_Report As Long
            '������� ��������� "Z_������"
Dim strZ_Report As String
            '����� ������� ������ � "��������� �������"
Dim intRowNumSys As Integer
            '��� �������� ��� ���������� "��������� �������"
Dim intSaveTableSystem As Integer
            
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
            imgCalendar.Visible = False
            fraDayNight.Visible = False
            '������� ������������ �������� ���������� ����� "DataAccessServ"
            lblInfo.Enabled = False
            txtInfo.Enabled = False
            lblLat0.Enabled = False
            lblLat320.Enabled = False
            hsbLat.Enabled = False
            
            '�������� "Z_�����"
            If Right(txtPersonCode.Text, 8) = "Z_Report" Then
            '�������� ������� ����� "Z_������"
                lngZ_Report = 0
            '���� �� ���� ������� "������� ���������"
                For intRowNum = 1 To gProtocRowNum - 1 Step 1
            '������ ������ "������� ���������" �� ����� � �����
                    Get gProtocFileNum, intRowNum, gProtocol
            '������� ���������:
            '   - ����������� �������� ������� ����������� ���
            '   - ���������� �������� ������� �����������
            '   - ��������� ����������� ������� �����������
            '   - ����������� �������� ���������� ����������� ���
            '   - ���������� �������� ���������� �����������
            '   - ��������� ����������� ���������� �����������
                    If ((Left(Trim(gProtocol.strProtocStatus), 2) = "05" Or _
                    Left(Trim(gProtocol.strProtocStatus), 2) = "06") And _
                    (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegPark" Or _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelPark") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "05" And _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoCorPark") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "08" Or _
                    Left(Trim(gProtocol.strProtocStatus), 2) = "09") And _
                    (Left(Trim(gProtocol.strProtocReserve), 11) = "AutoRegAcce" Or _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoDelAcce") Or _
                    (Left(Trim(gProtocol.strProtocStatus), 2) = "08" And _
                    Left(Trim(gProtocol.strProtocReserve), 11) = "AutoCorAcce")) And _
                    Left(gProtocol.strProtocName, 1) <> "@" Then
            '�� ���� ����� ������������ "Z_�����" ��� ���������
                        If gZ_Report = Trim(gProtocol.strProtocTime) + _
                        Left(Trim(gProtocol.strProtocDate), 6) Then
            '�������� ������� ����� "Z_������"
                            lngZ_Report = 0
            '�������� ��������� ������� ����� "Z_������"
                            strZ_Report = ""
                        Else
            '�������������� ������� ����� � ��������� ����� "Z_������"
                            If Mid(gProtocol.strProtocReserve, 13, 3) <> "   " Then
                                lngZ_Report = lngZ_Report + _
                                Mid(Trim(gProtocol.strProtocReserve), 13, 3) * 100 + _
                                Mid(Trim(gProtocol.strProtocReserve), 17, 2)
                                strZ_Report = Trim(gProtocol.strProtocTime) + _
                                Left(Trim(gProtocol.strProtocDate), 6)
                            End If
                        End If
                    End If
                Next
            '��������� ���������� ���� "����������"
                txtMoneyDate.Text = Trim(Str(Int(lngZ_Report / 100))) + " Ls " + _
                Trim(Str(lngZ_Report - Int(lngZ_Report / 100) * 100)) + " s"
               
            '����� "������ ���������" �����������
        If chkDocument.Value = 1 Then
            '������� ����������� ������ "Exit_Cancel"
                cmdCancel.Enabled = False
            '������ ��������� (�������� �� �����-�����, ��������
            '  ���� �/��� ��������� ����)
            Call frmDemo.PrintZReport(gProtocol.strProtocName, _
            gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
            gProtocol.strProtocTime, gProtocol.strProtocDate, _
            gProtocol.strProtocReserve, txtMoneyDate.Text, strZ_Report)
            '������� ��������� ������ "Exit_Cancel"
                cmdCancel.Enabled = True
        End If
            
            '��������� ����� ����� "Z_������"
                If strZ_Report <> "" Then
                    gZ_Report = strZ_Report
            '������������� ������� "��������� �������" (������)
                    frmTableSystem.grdTableSystem.Col = 0
            '���� �� ���� ��������������� ������� "��������� �������"
                    For intRowNumSys = 1 To _
                    frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������  "��������� �������"
                        frmTableSystem.grdTableSystem.Row = intRowNumSys
            '������ "��������� �������" � ���������� ����� "Z_������"
                        If Trim(frmTableSystem.grdTableSystem.Text) = "Z_Report" Then
            '������� ������� "��������� �������"=1(���������)
                            frmTableSystem.grdTableSystem.Col = 1
            '���� � ����� ������������ ���������� "Z_������"
                            frmTableSystem.grdTableSystem.Text = gZ_Report
                            Exit For
                        End If
                    Next
            '��������� ����� '��������� �������"
                    intSaveTableSystem = frmTableSystem.SaveTableSystem()
                End If
                
            '���������� ����� �� ������ "Exit_Cancel"
                cmdCancel.SetFocus
                Exit Sub
            
            '�������� "������������ ���"
            Else
            '����� ���������-������� ���������� ������������� ����
                intAutoFindCode = _
                frmTablePerson.AutoFindAccess(txtPersonCode.Text, _
                strInfo, strStatus, strChecking)
            End If
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
            If intAutoFindCode = 0 Then
            
            '������ ������� ���������� �����������
            
            '������������ ��� ����������� ������ ����������
                If Left(Trim(strStatus), 2) <> "08" Then
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
                    GoTo PersonCodeError
                End If
            
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

            '��������� ����� ��������� ���������� � ����������
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
            '���� ��������� ���������� � ����������
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            '������� ������� ��������������� "������"
                imgCalendar.Visible = True
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
                txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
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
            '������� ���������� �������� ���������� ����� "DataAccessServ"
                lblInfo.Enabled = True
                txtInfo.Enabled = True
                lblLat0.Enabled = True
                lblLat320.Enabled = True
                hsbLat.Enabled = True
    
            '���� ���������� ��� �������� ��������
                strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            '����� ���������-������� ���������� ����
            '  ���������� ��� �������� ��������
                intDayReg = Left(strDate, 2)
                intMonthReg = Mid(strDate, 4, 2)
                intYearReg = Right(strDate, 4)
            '���������� ���� , ��������������� ���,
            '  ���������� �� ��������� ���� �������� ��������
                frmTableCalendar.comCalendar.Day = intDayReg
                frmTableCalendar.comCalendar.Month = intMonthReg
                frmTableCalendar.comCalendar.Year = intYearReg
                frmTableCalendar.comCalendar.NextDay
                intDayReg = frmTableCalendar.comCalendar.Day
                intMonthReg = frmTableCalendar.comCalendar.Month
                intYearReg = frmTableCalendar.comCalendar.Year
            '����� ��� ���������� ����
                txtInfo.BackColor = vbWhite
                txtMoneyDate.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "txtPersonCode"
                If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
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
            '��������� ��������� "������ ����" �� ���� ����������
Private Sub txtInfo_Click()
            '����� ��� ���������� ����
    txtInfo.BackColor = vbWhite
            '�������� �������  ��������� � ��������� ���� "Info"
    txtInfo.Tag = 0
            '����������� ����������� ����������
    If txtMoneyDate.Tag = 0 Then
            '������� ����������� ������� �� ������ "OK"
         cmdOK.MousePointer = vbNoDrop
    End If

End Sub
            
            '��������� ����� � ������� ���������� ���� "Info"
Private Sub txtInfo_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '������� ��� ���������� ����
        txtInfo.BackColor = vbCyan
            '���������� � ���������� ���������
        If Len(Trim(txtInfo.Text)) < 17 And Len(Trim(txtInfo.Text)) > 0 Then
            '���������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 1
            '������� ��������� ������� �� ������ "OK"
            cmdOK.MousePointer = 0
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
            '����������� ����������� ����������
            If txtMoneyDate.Tag = 0 Then
            '������� ����������� ������� �� ������ "OK"
                cmdOK.MousePointer = vbNoDrop
            End If
        End If
    End If

End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Change()
            
            '�������� ������ ��������� ����� "������" ������
    If hsbLat.Value > hsbLat.Tag And (hsbLat.Tag * 100 + intAccessTariff) > 32000 Then
            '�������������� ����������� ��������� ���������
        hsbSant.Value = hsbSant.Tag
        hsbLat.Value = hsbLat.Tag
    ElseIf hsbLat.Value = hsbLat.Tag Then
        Exit Sub
    End If
            '�������� ����� ������ � ��������
    intAccessMoney = hsbLat.Value * 100 + hsbSant.Value
            '�������� ����� ��������� ����� � �������� � ������������
            '  ��������� (�������� ����� �� ���������� ����� ����� ����)
    If Int(intAccessMoney / intAccessTariff) * 100 <> intAccessMoney Or _
    hsbLat.Value * 100 > intAccessTariff Then
            '�������� �������� � ������� ���������� �����
        If hsbLat.Value > hsbLat.Tag Then
            intAccessMoney = hsbLat.Tag * 100 + hsbSant.Tag + intAccessTariff
            '�������� �������� � ������� ���������� �����
        ElseIf hsbLat.Value < hsbLat.Tag Then
            intAccessMoney = hsbLat.Tag * 100 + hsbSant.Tag - intAccessTariff
        End If
            '�������������� ����������� ��������� ���������
            hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
            hsbLat.Value = Int(intAccessMoney / 100)
            '��������� ����� ��������� ���������
        hsbSant.Tag = hsbSant.Value
        hsbLat.Tag = hsbLat.Value
    End If
            '��������� ����� ��������� ���������
    hsbSant.Tag = hsbSant.Value
    hsbLat.Tag = hsbLat.Value
            
            '����
    txtMoneyDate.Text = strDate
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
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
            '�� ������� ��������� ������ �� ��������� ����� ���������
    If hsbLat.Value > 0 Or hsbSant.Value > 0 Then
            '���������� �������  ��������� ����������
        txtMoneyDate.Tag = 1
            '���������� ���� �������� ��������
        intAccessDay = Int(intAccessMoney / intAccessTariff)
            '�������������� ��������� �������� "���������"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            '���� �� ���� "���������" (�� ����������
            '  ��� �������� �������� +1)
        For intAccessDay = intAccessDay To 1 Step -1
            
            '������ �����, ������ � ���� � ���� "����������"
            If frmTableCalendar.comCalendar.Month > 9 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year))
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                Trim(Str(frmTableCalendar.comCalendar.Day)) + ".0" + _
                Trim(Str(frmTableCalendar.comCalendar.Month)) + "." + _
                Trim(Str(frmTableCalendar.comCalendar.Year))
            End If
            '����������� "���������" �� ���� ���� ������
            frmTableCalendar.comCalendar.NextDay
            
        Next
            
            '������ ��������� ����������
    Else
        txtMoneyDate.Tag = 0
            '����
        txtMoneyDate.Text = strDate
            '������������ ������� � ���� "����������"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK"
        If txtInfo.Tag = 0 Then
            cmdOK.MousePointer = vbNoDrop
        End If
    End If
EndCycle:
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '������������� ������ ������ ���
    If Int(intAccessMoney / intAccessTariff) = 0 Then
           '������ ��������� ����������
       txtMoneyDate.Tag = 0
           '����
       txtMoneyDate.Text = strDate
           '������������ ������� � ���� "����������"
       txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK"
        If txtInfo.Tag = 0 Then
            cmdOK.MousePointer = vbNoDrop
        End If
    End If
            '��������� (�������� ������ ������ �� ����� ����)
    If intAccessDay > 0 Then
            '���������� ���������� (��� ���������) ���� �������� ��������
        intAccessDay = Int(intAccessMoney / intAccessTariff) - intAccessDay
           '�������������� ���������� (��� ���������) ����� ������ � ��������
        intAccessMoney = intAccessDay * intAccessTariff
            '�������������� ����������� ��������� ���������
        hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
        hsbLat.Value = Int(intAccessMoney / 100)
        hsbLat_Change
    End If
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If
    
End Sub


