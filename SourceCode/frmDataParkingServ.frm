VERSION 5.00
Begin VB.Form frmDataParkingServ 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ParkingServData"
   ClientHeight    =   3120
   ClientLeft      =   4860
   ClientTop       =   2565
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   186
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
   ScaleHeight     =   3120
   ScaleWidth      =   7095
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Frame fraMonth 
      Caption         =   " D   1/2M  1M   2M"
      Height          =   615
      Left            =   5160
      TabIndex        =   22
      Top             =   960
      Width           =   1695
      Begin VB.OptionButton optTwo 
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optOne 
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optHalf 
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optNot 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.VScrollBar vsbDate 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4800
      Max             =   366
      TabIndex        =   21
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtPersonCode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   720
      TabIndex        =   15
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   720
      TabIndex        =   14
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   4200
      TabIndex        =   13
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox chkDocument 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Width           =   255
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
      TabIndex        =   11
      Top             =   1680
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
      TabIndex        =   10
      Top             =   2400
      Width           =   1212
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   320
      TabIndex        =   9
      Top             =   1680
      Width           =   1452
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4920
      Max             =   99
      TabIndex        =   8
      Top             =   2520
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
      Height          =   2055
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1455
      Begin VB.Frame fraDayNight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optDay 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingServ.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   " ""-1"" D/M ""+1"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   27
      Top             =   960
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   2400
      Y2              =   2400
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
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   1440
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Image imgDocument 
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataParkingServ.frx":0802
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Image imgParkingServ 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataParkingServ.frx":0C18
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   600
      Y2              =   2400
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
      Top             =   1680
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
      TabIndex        =   17
      Top             =   1680
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6960
      Y1              =   2400
      Y2              =   2400
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
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "frmDataParkingServ"
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
Dim intParkingMoney As Integer
            '���������� ����������� ����
Dim intParkingDay As Integer
            '����� ������ ������������ ��� (�����)
Dim intParkingTariffFull As Integer
            '����� ������ ������������ ��� (����)
Dim intParkingTariffDay As Integer
            '����� ������ ������������ ��� (����)
Dim intParkingTariffNight As Integer
            '����� ����������� (���������� ��� ���������)
Dim intParkingTariff As Integer
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
            '������ "������� ���������" ��������������� ����
            '  ���������� ������������ ���
Dim intRowNumReg As Integer
            '������� "������� ���������", ��������������� ����
            '  ���������� ������������ ���
Dim intColNumReg As Integer

            '�������� ������� ���������� ������ "Alt"+ {"OK" � "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '����� "frmDataParkingServ" ��������
    If frmDataParkingServ.Enabled = True Then
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
            '���� ���������� ������������ ��� (� �����)
            '  ��������� ���������� � �������
Dim strDate As String
            '����� ����������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ����������� �������
Dim strHour As String
Dim strMinute As String
            '������� ����������� �� \ 0 - ������ \ 1 - ������ \ 2 - ���������������
Dim strCarPresent As String * 1
            '������� �� ("�" - ������������ ������; "D" - ������� ����� �������
            '  �� �����������; "N" - ������ ����� ������� �� �����������; "������
            '  ������"   - �������� ����� ������� �� �����������)
Dim strExpander As String * 1
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ������������� � "������� ������"
Dim intAutoCorrectionCode  As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer

            '����������� ������� �� ������ "OK_-"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            '������� ������ "��������" ��� �����������
    strChecking = ""
            '��������� ����� ��������� ���������� � �������
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
            
            '������� ��������������� �� �������
    strCarPresent = "2"
            '������� �� �������
    strExpander = "P"
            '������ � ������� ������� ������� �� �����������
    If optDay.Value = True Then
        strExpander = "D"
            '������ � ������ ������� ������� �� �����������
    ElseIf optNight.Value = True Then
        strExpander = "N"
    End If
    
            '���� ���������� ����������� ������������ ���
    strDate = Mid(Trim(txtMoneyDate.Text), 11)
    If Mid(Trim(strDate), 3, 1) = "." Then
        strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
        Mid(Trim(strDate), 7, 4)
    Else
        strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
        Mid(Trim(strDate), 6, 4)
    End If
    
            '������������ ����������� ������ "��������" ��� �����������
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
            
            '�������� �� ������� � ��� ����������� �� �����������
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            '����� ���������-������� ������������� ��� �������
            '������������� ���� ��� �����������
    intAutoCorrectionCode = frmTablePerson.AutoCorParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoCorPark " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            '������� (����)��������� ��� ������� ������������� ����
        frmDataParkingServ.Tag = 1
        
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
        gProtocol.strProtocReserve = "Invalid AutoCorParking"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '����� ��� ���������� ����
        txtInfo.BackColor = vbWhite
        txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)��������� ��� ������� ������������� ����
        frmDataParkingServ.Tag = 2
            
            
            
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
    If frmDataParkingServ.Tag = 0 Then frmDataParkingServ.Tag = 2
            '������� ��������� ������� �����
    frmDataParkingServ.Visible = False
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
            
            '������� ����� "optNot"
    optNot.Value = True
            '������� ������������ �������� ���������� ����� "DataParkingServ"
    txtInfo.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    txtMoneyDate.Enabled = False
    fraMonth.Enabled = False
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
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '���������� ����� �� ��������� ���� "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK"
    cmdOK.MousePointer = vbNoDrop
             '���������� ���� ���������� ����������� ������� �����
    frmDataParkingServ.Tag = 1

End Sub

            '������������� ������� �����
Private Sub Form_Deactivate()
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub
            
            '�������� ������� �����
Private Sub Form_Load()
            '������� ����� "optNot"
    optNot.Value = True
            '������� ������������ �������� ���������� ����� "DataParkingServ"
    txtInfo.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    txtMoneyDate.Enabled = False
    fraMonth.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� �������� ���������� ����� "DataParkingServ"
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
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '����� ������ ������������ ��� (�����)
    intParkingTariffFull = gParkingDN
            '����� ������ ������������ ��� (����)
    intParkingTariffDay = gParkingD
            '����� ������ ������������ ��� (����)
    intParkingTariffNight = gParkingN
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
            '�������� ��������� ���� "����������" ��� �����������
    txtInfo.Text = ""
            '�������� ��������� ���� "����������" ��� �����������
    txtMoneyDate.Text = ""
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� "������"
    imgCalendar.Visible = False
    fraDayNight.Visible = False
            '������� ������������ �������� ���������� ����� "DataParkingServ"
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '������� ����������� ������� �� ������ "OK"
    cmdOK.MousePointer = vbNoDrop

End Sub

            '��������� ����� � ������� "PersonCode"
Private Sub txtPersonCode_KeyPress(KeyAscii As Integer)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            '������ "����������" ��� �����������
Dim strInfo As String
            '������
Dim strStatus As String
            '����� ��������� ���������� � �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ���������
            '  ���������� � �������
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
            '���������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 1
            '�������� ��������� ���� "����������" ��� �����������
            txtInfo.Text = ""
            '�������� ��������� ���� "����������" ��� �����������
            txtMoneyDate.Text = ""
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
            lblMoneyDate.Visible = False
            '������� ���������� "������"
            imgCalendar.Visible = False
            fraDayNight.Visible = False
            '������� ������������ �������� ���������� ����� "DataParkingServ"
            hsbLat.Enabled = False
            vsbDate.Enabled = False
            fraMonth.Enabled = False
            
            
            '�������� "Z_�����"
            If Right(txtPersonCode.Text, 8) = "Z_Report" Then
            '�������� ������� ����� "Z_������"
                lngZ_Report = 0
            '�������� ��������� ������� ����� "Z_������"
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
            '����� ���������-������� ����������
            '������������� ���� ��� �����������
                intAutoFindCode = _
                frmTablePerson.AutoFindParking(txtPersonCode.Text, _
                strInfo, strStatus, strChecking)
            End If
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
            If intAutoFindCode = 0 Then
            
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
                If Left(Trim(strStatus), 2) <> "05" Then
            '���� �������� � ��������  ������� ������� ����������� - �� �����
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
            
            '��������� ��������� ���� "����������" ��� �����������
                txtInfo.Text = strInfo
            
            '���������� ������ "��������"
                Call frmTablePerson.UnPack(strDate, strChecking)
            
            '����������� ������������� ������ "��������" ��� �����������
                txtMoneyDate.Text = Left(Trim(strDate), 2) + "." + _
                Mid(Trim(strDate), 3, 2) + "." + Mid(Trim(strDate), 5, 4) + "/" + _
                Mid(Trim(strDate), 9, 2) + ":" + Mid(Trim(strDate), 11, 2) + "/"
            '������� �����������/������/������ �� �������
                If Mid(Trim(strChecking), 7, 1) = "0" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "+"
                ElseIf Mid(Trim(strChecking), 7, 1) = "1" Then
                    txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "_"
                ElseIf Mid(Trim(strChecking), 7, 1) = "2" Then
                txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "?"
                End If

            '��������� ����� ��������� ���������� � �������
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
            '���� ��������� ���������� � �������
                gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            '������� ������� ��������������� "������"
                imgCalendar.Visible = True
            '���������� � ������� ������� ���������������
            '  ������� �� �������
                If Right(Trim(strChecking), 1) = "D" Then
            '������� ����� ������� � �����������
                    optDay.Value = True
                ElseIf Right(Trim(strChecking), 1) = "N" Then
            '������ ����� ������� � �����������
                    optNight.Value = True
                ElseIf Right(Trim(strChecking), 1) <> "D" And _
                Right(Trim(strChecking), 1) <> "N" Then
            '�������� ����� ������� � �����������
                    optDayNight.Value = True
                End If
            '���������� �� ������ ������������ � �����������
                If Right(Trim(strChecking), 1) <> "E" Then
                    fraDayNight.Visible = True
            '���������� ������ ������������ � �����������
                Else
                    Exit Sub
                End If
            '������������ ������� � ���� "����������"
                txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '����������� ������ ����������� (���������� ��� ���������)
            '������� ����� ������� � �����������
                If optDay.Value = True Then
                    intParkingTariff = intParkingTariffDay
            '������ ����� ������� � �����������
                ElseIf optNight.Value = True Then
                    intParkingTariff = intParkingTariffNight
            '�������� ����� ������� � �����������
                ElseIf optDayNight.Value = True Then
                    intParkingTariff = intParkingTariffFull
                End If
            '������� ������� ����� ���������� ���� "txtMoneyDate"
                lblMoneyDate.Visible = True
            '������� ���������� �������� ���������� ����� "DataParkingServ"
                hsbLat.Enabled = True
                vsbDate.Enabled = True
                fraMonth.Enabled = True
    
            '���� ���������� ������������ ���
                strDate = Mid(Trim(txtMoneyDate.Text), 11, 10)
            '���������� ���� ���������� ������������ ���
                intDayReg = Left(strDate, 2)
                intMonthReg = Mid(strDate, 4, 2)
                intYearReg = Right(strDate, 4)
            '���������� ���� , ��������������� ���,
            '  ���������� �� ��������� ����������� ����
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
            '�������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 0
            '����� ��� ���������� ����
            txtInfo.BackColor = vbWhite
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
            '�������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 0
            '����� ��� ���������� ����
            txtInfo.BackColor = vbWhite
            '���������� ����� �� ��������� ���� "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            '������� ����������� ������� �� ������ "OK_-"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub
            
            '������� ����� - "Not"
Private Sub optNot_Click()
            '������� ������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = True
    lblDate.Visible = True
            '������� ���������� ��������� �������� ���������� �����
    hsbLat.Enabled = True
    vsbDate.Enabled = True
    fraMonth.Enabled = True
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            '������� ����� - "Half"
Private Sub optHalf_Click()
            '���������� ���� �� ����� ������
    Dim intToMonthEnd As Integer
            '�������� ����� �������� ����������
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            '��������� ���������� �� 1/2 �����a
    Dim intLat As Integer
    Dim intSant As Integer
            
            '����� ������ ������������ ��� (�����)
    intParkingTariff = intParkingTariffFull
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            '�������������� ��������� ����
            '���� ���������� ������������ ���
    txtMoneyDate.Text = strDate
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
            
            '���� �� ������� "������� ���������" (� ���������� ������������ ���)
    For intRowNum = intRowNumReg To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = intRowNumReg Then
            '�� �������� "������� ���������" (� ���������� ������������ ���)
            intColNum = intColNumReg
        Else
            '�� ���� �������� "������� ���������"
            intColNum = 1
        End If
            '�� ���� �������� "������� ���������"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ���� �� ����� ������
            intToMonthEnd = intToMonthEnd + 1
            '������� �������� ����� � ������� ������ "������� ���������"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              '������� ������
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            Else
              '������� ������
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            '������� ����
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - �� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And 0 = gYear Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And Not (0 = gYear Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            '������� ������ - �� ����
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
            '������� ������ - �� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 7 Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - ����� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
    End If
    
              '������� ���� � ������
    If Mid(txtMoneyDate.Text, 4, 2) = 12 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              '��������� ����� ����� �� ����
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 16)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              '���� �� ����� - �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 2 And intFinishDay <= 15 Then
        If 0 = gYear Mod 4 Then
            txtMoneyDate.Text = Trim(Str(intFinishDay + 14)) + _
            Trim(Mid(txtMoneyDate.Text, 3))
        Else
            txtMoneyDate.Text = Trim(Str(intFinishDay + 13)) + _
            Trim(Mid(txtMoneyDate.Text, 3))
        End If
              '���� �� ����� - �� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 15)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 And intFinishDay > 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay - 15)) + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              '���� �� �����
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 And intFinishDay <= 15 Then
        txtMoneyDate.Text = Trim(Str(intFinishDay + 15)) + _
        Trim(Mid(txtMoneyDate.Text, 3))
    End If
            
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ��������� ���������� �� 1/2 �����a
    intLat = intParkingTariff * 15
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            '��������� ���������� ���� "����������"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub
            
            '������� ����� - "One"
Private Sub optOne_Click()
            '���������� ���� �� ����� ������
    Dim intToMonthEnd As Integer
            '�������� ����� �������� ����������
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            '��������� ��������� ���������� �� 1 �����
    Dim intLat As Integer
    Dim intSant As Integer
            
            '����� ������ ������������ ��� (�����)
    intParkingTariff = intParkingTariffFull
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            '�������������� ��������� ����
    txtMoneyDate.Text = strDate
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
            '���� �� ������� "������� ���������" (� ���������� ������������ ���)
    For intRowNum = intRowNumReg To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = intRowNumReg Then
            '�� �������� "������� ���������" (� ���������� ������������ ���)
            intColNum = intColNumReg
        Else
            '�� ���� �������� "������� ���������"
            intColNum = 1
        End If
            '�� ���� �������� "������� ���������"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ���� �� ����� ������
            intToMonthEnd = intToMonthEnd + 1
            '������� �������� ����� � ������� ������ "������� ���������"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              '������� ������
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            Else
              '������� ������
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            '������� ����
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - �� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And 0 = gYear Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 1 And Not (0 = gYear Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            '������� ������ - �� ����
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
            '������� ������ - �� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 7 Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - ����� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 7 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
    End If
    strFinishDay = Trim(Str(intFinishDay))
    
              '������� ���� � ������
    If Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        txtMoneyDate.Text = strFinishDay + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 9 Then
        txtMoneyDate.Text = strFinishDay + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 9 Then
        txtMoneyDate.Text = strFinishDay + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 1)) + Trim(Mid(txtMoneyDate.Text, 6))
    End If
            
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ��������� ���������� �� 1 �����
    intLat = intParkingTariff * 30
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            '��������� ���������� ���� "����������"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub
            
            '������� ����� - "Two"
Private Sub optTwo_Click()
            '���������� ���� �� ����� ������
    Dim intToMonthEnd As Integer
            '�������� ����� �������� ����������
    Dim intFinishDay As Integer
    Dim strFinishDay As String
            '��������� ��������� ���������� �� 2 �����a
    Dim intLat As Integer
    Dim intSant As Integer
            
            '����� ������ ������������ ��� (�����)
    intParkingTariff = intParkingTariffFull
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
            '�������������� ��������� ����
    txtMoneyDate.Text = strDate
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
            '���� �� ������� "������� ���������" (� ���������� ������������ ���)
    For intRowNum = intRowNumReg To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        frmTableCalendar.grdTableCalendar.Row = intRowNum
        If intRowNum = intRowNumReg Then
            '�� �������� "������� ���������" (� ���������� ������������ ���)
            intColNum = intColNumReg
        Else
            '�� ���� �������� "������� ���������"
            intColNum = 1
        End If
            '�� ���� �������� "������� ���������"
        For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ���� �� ����� ������
            intToMonthEnd = intToMonthEnd + 1
            '������� �������� ����� � ������� ������ "������� ���������"
            intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
            If intPosNum <> 0 Then
              '������� ������
                If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            Else
              '������� ������
                If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������� ���� �� ����� ������ (��� �������� ���)
                    intToMonthEnd = intToMonthEnd - 1
            '���������� ���� �� ����� ������ ���������
                    GoTo EndCycle
                End If
            End If
        Next
    Next
EndCycle:
            
            '������� ����
    If Mid(txtMoneyDate.Text, 4, 2) = 11 Then
        intFinishDay = 31 - intToMonthEnd
            '�������  ���� & ������ - �� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And 0 = (gYear + 1) Mod 4 Then
        intFinishDay = 29 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 And Not (0 = (gYear + 1) Mod 4) Then
        intFinishDay = 28 - intToMonthEnd
            '������� ������ - �� ����
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 6 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 30 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 6 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - �� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 6 Then
        intFinishDay = 31 - intToMonthEnd
            '������� ������ - ����� �������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 6 And _
    0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2 Then
        intFinishDay = 31 - intToMonthEnd
    ElseIf Mid(txtMoneyDate.Text, 4, 2) > 6 And _
    Not (0 = CInt(Mid(txtMoneyDate.Text, 4, 2)) Mod 2) Then
        intFinishDay = 30 - intToMonthEnd
    End If
    strFinishDay = Trim(Str(intFinishDay))
    
              '������� ���� � ������
    If Mid(txtMoneyDate.Text, 4, 2) = 11 Then
        txtMoneyDate.Text = strFinishDay + ".01." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
    ElseIf Mid(txtMoneyDate.Text, 4, 2) = 12 Then
        txtMoneyDate.Text = strFinishDay + ".02." + _
        Trim(Str(Mid(txtMoneyDate.Text, 7) + 1))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) >= 8 Then
        txtMoneyDate.Text = strFinishDay + "." + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 2)) + Trim(Mid(txtMoneyDate.Text, 6))
              '������� ������
    ElseIf Mid(txtMoneyDate.Text, 4, 2) < 8 Then
        txtMoneyDate.Text = strFinishDay + ".0" + _
        Trim(Str(Mid(txtMoneyDate.Text, 4, 2) + 2)) + Trim(Mid(txtMoneyDate.Text, 6))
    End If
            
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ��������� ���������� �� 2 �����a
    intLat = intParkingTariff * 60
    intSant = intLat - Int(intLat / 100) * 100
    intLat = Int(intLat / 100)
            '��������� ���������� ���� "����������"
    If intLat < 10 Then
        txtMoneyDate.Text = "00" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat < 100 Then
        txtMoneyDate.Text = "0" + Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    ElseIf intLat > 99 Then
        txtMoneyDate.Text = Trim(Str(intLat)) + Mid(txtMoneyDate.Text, 4)
    End If
    If intSant < 10 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + Trim(Str(intSant)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Date"
Private Sub vsbDate_Scroll()
            '���������� �������  ��������� ����������
    If txtMoneyDate.Tag = 1 Then vsbDate_Change
    
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Date"
Private Sub vsbDate_Change()
            
            '�� ���������� �������  ��������� ����������
    If txtMoneyDate.Tag = 0 Then
            '�������������� ����������� ��������� �������a
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            
            '�������� ������ ��������� ���� "������" ������
    If vsbDate.Value >= vsbDate.Max And vsbDate.Tag = vsbDate.Max Then
            '�������������� ����������� ��������� �������a
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            '������o� ������ ��������� ���� � ������������ ���������
            '  (����� �� ������� �������� ���������)
    If vsbDate.Value > ((frmTableCalendar.grdTableCalendar.Rows - 1) * 7 - _
    ((gRowNum - 1) * 7 + gColNum)) Then
            '�������������� ����������� ��������� �������a
        vsbDate.Value = vsbDate.Tag
        Exit Sub
    End If
            
            '��������� ����� ��������� �������a
    vsbDate.Tag = vsbDate.Value
            
            '������� ����� "=" � ���� "����������"
    intPosNum = InStr(1, Trim(txtMoneyDate.Text), "=")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = Left(Trim(txtMoneyDate.Text), intPosNum) + _
    Format(Now, "dd/mm/yyyy")
            
            '�� ������� ��������� �������a
    If vsbDate.Value > 0 Then
            '���������� ����������� ����
        intParkingDay = vsbDate.Value
            '���� �� ������� "������� ���������" (� �������� ���)
        For intRowNum = gRowNum To frmTableCalendar.grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
            frmTableCalendar.grdTableCalendar.Row = intRowNum
            If intRowNum = gRowNum Then
            '�� �������� "������� ���������" (� �������� ���)
                intColNum = gColNum
            Else
            '�� ���� �������� "������� ���������"
                intColNum = 1
            End If
            '�� ���� �������� "������� ���������"
            For intColNum = intColNum To frmTableCalendar.grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
                frmTableCalendar.grdTableCalendar.Col = intColNum
            '���������� ����������� ���� ���������
                If intParkingDay < 0 Then GoTo EndCycle
            '���������� ����������� ����
                intParkingDay = intParkingDay - 1
            '������� �������� ����� � ������� ������ "������� ���������"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            '���������  ����� � ������ � ���� "����������"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) + _
                    Right(txtMoneyDate.Text, 8)
            '������� ������
                    If Left(Trim(frmTableCalendar.grdTableCalendar.Text), intPosNum - 1) = "1" _
                    And (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������  ������ � ���� "����������"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
            '������� ����
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            '���������  ������ � ���� � ���� "����������"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                Else
            '���������  ����� � ������ � ���� "����������"
                    txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + _
                    Trim(frmTableCalendar.grdTableCalendar.Text) + _
                    Right(txtMoneyDate.Text, 8)
            '������� ������
                    If Trim(frmTableCalendar.grdTableCalendar.Text) = "1" And _
                    (intRowNum <> gRowNum Or intColNum <> gColNum) Then
            '���������  ������ � ���� "����������"
                        If CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1 > 9 Then
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1." + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        Else
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 10) + "1.0" + _
                            Trim(Str(CInt(Mid(txtMoneyDate.Text, 13, 2)) + 1)) + _
                            Right(txtMoneyDate.Text, 5)
                        End If
            '������� ����
                        If Mid(txtMoneyDate.Text, 13, 2) = "13" Then
            '���������  ������ � ���� � ���� "����������"
                            txtMoneyDate.Text = Left(txtMoneyDate.Text, 12) + "01." + _
                            Trim(Str(CInt(Right(txtMoneyDate.Text, 4)) + 1))
                        End If
                    End If
                End If
            Next
        Next
    End If
EndCycle:
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If
    
End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Scroll()
    hsbLat_Change
    
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Lat"
Private Sub hsbLat_Change()
            
            '�������� ������ ��������� ����� "������" ������
    If hsbLat.Value > hsbLat.Tag And (hsbLat.Tag * 100 + intParkingTariff) > 32000 Then
            '�������������� ����������� ��������� ���������
        hsbSant.Value = hsbSant.Tag
        hsbLat.Value = hsbLat.Tag
    ElseIf hsbLat.Value = hsbLat.Tag Then
        Exit Sub
    End If
            '�������� ����� ������ � ��������
    intParkingMoney = hsbLat.Value * 100 + hsbSant.Value
            '�������� ����� ��������� ����� � �������� � ������������ ���������
            '  (�������� ����� �� ���������� ����� ����� ����������� ����)
    If Int(intParkingMoney / intParkingTariff) * 100 <> intParkingMoney Or _
    hsbLat.Value * 100 > intParkingTariff Then
            '�������� �������� � ������� ���������� �����
        If hsbLat.Value > hsbLat.Tag Then
            intParkingMoney = hsbLat.Tag * 100 + hsbSant.Tag + intParkingTariff
            '�������� �������� � ������� ���������� �����
        ElseIf hsbLat.Value < hsbLat.Tag Then
            intParkingMoney = hsbLat.Tag * 100 + hsbSant.Tag - intParkingTariff
        End If
            '�������������� ����������� ��������� ���������
            hsbSant.Value = intParkingMoney - Int(intParkingMoney / 100) * 100
            hsbLat.Value = Int(intParkingMoney / 100)
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
            '���������� ����������� ����
        intParkingDay = Int(intParkingMoney / intParkingTariff)
            '�������������� ��������� �������� "���������"
        frmTableCalendar.comCalendar.Day = intDayReg
        frmTableCalendar.comCalendar.Month = intMonthReg
        frmTableCalendar.comCalendar.Year = intYearReg
            '���� �� ���� "���������" (�� ����������
            '  ������������ ��� +1)
        For intParkingDay = intParkingDay To 1 Step -1
            
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
    If Int(intParkingMoney / intParkingTariff) = 0 Then
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
    If intParkingDay > 0 Then
            '���������� ���������� (��� ���������) ����������� ����
        intParkingDay = Int(intParkingMoney / intParkingTariff) - intParkingDay
           '�������������� ���������� (��� ���������) ����� ������ � ��������
        intParkingMoney = intParkingDay * intParkingTariff
            '�������������� ����������� ��������� ���������
        hsbSant.Value = intParkingMoney - Int(intParkingMoney / 100) * 100
        hsbLat.Value = Int(intParkingMoney / 100)
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
