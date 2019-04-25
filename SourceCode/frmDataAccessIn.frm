VERSION 5.00
Begin VB.Form frmDataAccessIn 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessInData"
   ClientHeight    =   3945
   ClientLeft      =   4485
   ClientTop       =   2925
   ClientWidth     =   6990
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
   ScaleHeight     =   3945
   ScaleWidth      =   6990
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.Frame fraPeople 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   2175
      Begin VB.OptionButton optFamily 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optConvoy 
         Caption         =   "Option1"
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optBaby 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optHuman 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   19
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   18
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   17
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   99
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   320
      TabIndex        =   15
      Top             =   1920
      Width           =   1452
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
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraStatus 
      Caption         =   "????"
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
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton optTime 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   12
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optCalendar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   720
         TabIndex        =   11
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton optMoneyFree 
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
         Height          =   252
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Frame fraDayNight 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
         Begin VB.OptionButton optDay 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Image imgCalendar 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image imgTime 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessIn.frx":24A4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
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
      TabIndex        =   14
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   2040
      Picture         =   "frmDataAccessIn.frx":28FE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   720
      Picture         =   "frmDataAccessIn.frx":2F30
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   120
      Picture         =   "frmDataAccessIn.frx":356E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   1200
      Picture         =   "frmDataAccessIn.frx":3DA8
      Stretch         =   -1  'True
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line7 
      X1              =   2760
      X2              =   4080
      Y1              =   1440
      Y2              =   1440
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
      TabIndex        =   25
      Top             =   240
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
      TabIndex        =   24
      Top             =   1920
      Width           =   495
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
      Left            =   6360
      TabIndex        =   23
      Top             =   1920
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
      Left            =   4560
      TabIndex        =   22
      Top             =   1920
      Width           =   135
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
      Left            =   4560
      TabIndex        =   21
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessIn.frx":4B2E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2520
      X2              =   2520
      Y1              =   2280
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6840
      X2              =   4080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   720
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1680
      Y2              =   3720
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2520
      Y1              =   2280
      Y2              =   2280
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
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Image imgAccessIn 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataAccessIn.frx":4F44
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
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
      Left            =   4080
      TabIndex        =   20
      Top             =   1920
      Width           =   375
   End
End
Attribute VB_Name = "frmDataAccessIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '�������� ����� ������ � ��������
Dim intAccessMoney As Integer
            '���������� ���� �������
Dim intAccessDay As Integer
            '����� ������ ��� (�����)
Dim intAccessTariffFull As Integer
            '����� ������ ��� (����)
Dim intAccessTariffDay As Integer
            '����� ������ ��� (����)
Dim intAccessTariffNight As Integer
            '����� (���������� ��� ���������)
Dim intAccessTariff As Integer
            '������� ������� ������� - (��� ���������� �����������)
Dim strTime As String
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

            '�������� ������� ���������� ������ "Alt"+ {"+" � "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������� ����� ��������
    If Me.Enabled = True Then
            '������������ "������" ���� �� ������ "+"
        If KeyCode = 187 And Shift = 4 Then
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
            '������� ��������� ������� �� ������ "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub

            '������� � ��������� ��������� (������ "OK _ +")
Private Sub cmdOK_Click()
            '������
Dim strStatus As String
            '������ "��������"
Dim strChecking As String * 8
            '��������� "��������" ���� "txtInfo"
Dim strCheckingInfo As String * 8
            '���� (� �����) ����������� ���������� ���
            '  ���� ���������� ����������� ���
Dim strDate As String
            '��������� "����������� ���� � �����" ���� "txtInfo"
Dim strDateInfo As String
            '����� ����������� ����������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ����������� ����������
Dim strHour As String
Dim strMinute As String
            '������� ����������� \ 0 - ����� \ 1 - ����� \ 2 - ���������������
Dim strPersPresent As String * 1
            '������� ("�" - ������������ �����; "D" - ������� ����� �������;
            '  "N" - ������ ����� �������; "������ ������"   - �������� �����
            '  �������)
Dim strExpander As String * 1
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ��������������� � "������� ������"
Dim intAutoRegistrCode  As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            '��������� ������� ��������� # 1, 2, 3 � 4
Dim intLease As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '���������� ����� �������, � ������� �������� ����������� ����������
            '  ��� ����������� ����������� ���������� ���������� �� �����������
Dim intCellLimit As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer

            '����������� ������� �� ������ "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            
            '���� ���������� ���������
    If optTime.Value = True Then
    
            '�������� ���� "Tag" ����� "frmLease"
        frmLease.Tag = 0
            '������� �� ����� ����� "frmLease" � ������� ����������� 1
        frmLease.Show 1
            
            '������� ������������ ������ "OK _ +" � "Cancel _ Exit"
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
            
            '��������� � ���� "txtInfo" ���������� � ������� ���������
        If frmLease.Tag <> "Exit" Then
            '��������� ������� ��������� # 1, 2, 3 � 4
            intLease = 0
            If Mid(CStr(frmLease.Tag), 1, 1) = "+" Then intLease = gLease1
            If Mid(CStr(frmLease.Tag), 2, 1) = "+" Then intLease = intLease + gLease2
            If Mid(CStr(frmLease.Tag), 3, 1) = "+" Then intLease = intLease + gLease3
            If Mid(CStr(frmLease.Tag), 4, 1) = "+" Then intLease = intLease + gLease4
            intLease = intLease + CInt(Left(txtMoneyDate.Text, 3)) * 100 + _
            CInt(Mid(txtMoneyDate.Text, 5, 2))
            '������������ ������� � ���� "����������"
            txtMoneyDate.Text = "000,00" + Mid(txtMoneyDate.Text, 7)
            '��������� ���������� ���� "����������"
            If Int(intLease / 100) < 10 Then
                txtMoneyDate.Text = "00" + Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            ElseIf Int(intLease / 100) < 100 Then
                txtMoneyDate.Text = "0" + Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            ElseIf Int(intLease / 100) > 99 Then
                txtMoneyDate.Text = Trim(Str(Int(intLease / 100))) + Mid(txtMoneyDate.Text, 4)
            End If
            '��������� ���������� ���� "����������"
            If intLease - Int(intLease / 100) * 100 < 10 Then
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
                Trim(Str(intLease - Int(intLease / 100) * 100)) + _
                Mid(txtMoneyDate.Text, 7)
            Else
                txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
                Trim(Str(intLease - Int(intLease / 100) * 100)) + _
                Mid(txtMoneyDate.Text, 7)
            End If
        End If
            
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
            '���� �������� � ��������� �������� �����������
            '   ������������� ���� - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        If frmDemo.optEnglish = True Then
            strResponse = MsgBox("Addition PersonCode ?", intButtonsAndIcons, "Cancel")
        Else
            strResponse = MsgBox("Papild. person. kods ?", intButtonsAndIcons, "Cancel")
        End If
            '������ ������ "���"
        If strResponse = vbNo Then
            '������� ������ �� (����)����������� ������������� ����
            Me.Tag = 2
            '������� ���������� ������ "OK _ +" � "Cancel _ Exit"
            cmdOK.Enabled = True
            cmdCancel.Enabled = True
            '������������ ����������� ������ ����
            Me.MousePointer = 0
            '���������� ����� �� ������ "Cancel _ Exit"
            cmdCancel.SetFocus
            Exit Sub
        End If
    
    End If
            
            '�� ������� ����� ��������� ������
    If Left(Me.txtMoneyDate.Text, 9) <> "000,00 Ls" Then
            '�������� ���� "Tag" ����� "frmMinus"
        frmMinus.Tag = 0
            '������� �� ����� ����� "frmMinus" � ������� ����������� 1
        frmMinus.Show 1
            '����� �� ������ � �� (����)����������� ������������� ����
        If frmMinus.Tag = "Exit" Then
            '������� � ��������� ���������
            cmdCancel_Click
            Exit Sub
        End If
    End If
            
            '������� ������ � ��������� "��������"
    strChecking = ""
    strCheckingInfo = ""
            '��������� ����� ����������� ����������
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
            '���� ����������� ����������
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
            '������� ����������� ����������
    strPersPresent = "2"
            '������� ����������
    strExpander = "P"
            '������ ������� ����������
    If optMoneyFree.Value = True Then
            '���������� ����������
        strStatus = "10 - Access/Free"
    ElseIf optCalendar.Value = True Then
            '���������� ����������
        strStatus = "08 - Access/Calen."
            '���������� � ������� ������� �������
        If optCalendar.Value = True And optDay.Value = True Then
            strExpander = "D"
            '���������� � ������ ������� �������
        ElseIf optCalendar.Value = True And optNight.Value = True Then
            strExpander = "N"
        End If
            '���� ���������� ����������� ���
        strDate = Mid(Trim(txtMoneyDate.Text), 11)
        If Len(Trim(strDate)) = 10 Then
            strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
            Right(Trim(strDate), 4)
        Else
            strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
            Right(Trim(strDate), 4)
        End If
    ElseIf optTime.Value = True Then
            '��������� ����������
        strStatus = "09 - Access/Time"
            '���������� ����� �������, � ������� �������� ����������� ����������
            '  ���������� ���������� �� ����������� (���������� �����)
        intCellLimit = gAcceInpCellNumb
    
            '��������� "���������" ����� � ���� ����������� ����������,
            '  �� ������� ��� ����� �������� ����-�����
            
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
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
            
            '������� ����������� ���������� � ������ ��� ����������
    strChecking = Left(strChecking, 6) + strPersPresent + strExpander
            
            '���� ���������� ���������
    If optTime.Value = True Then
            '��������� � ���� "txtInfo" ���������� � ������� ���������
        If frmLease.Tag <> "Exit" Then _
        txtInfo.Text = Left(CStr(frmLease.Tag), 4) + Mid(txtInfo.Text, 5)
    End If
            '��������� � ���� "txtInfo" ���������� � ������� (��������)
    If optHuman.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "1" + Mid(txtInfo.Text, 6)
            '��������� � ���� "txtInfo" ���������� � ������� (����)
    ElseIf optBaby.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "2" + Mid(txtInfo.Text, 6)
            '��������� � ���� "txtInfo" ���������� � ������� (������)
    ElseIf optConvoy.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "3" + Mid(txtInfo.Text, 6)
            '��������� � ���� "txtInfo" ���������� � ������� (�����)
    ElseIf optFamily.Value = True Then
        txtInfo.Text = Left(txtInfo.Text, 4) + "4" + Mid(txtInfo.Text, 6)
    End If
            
            '���������� ���������� �� ����������� � ������������ �������
            '  ������������ ����������
    If gAcceTimeLimit > 0 And optCalendar.Value = True Then
            '���������� ����� �������, � ������� �������� ����������� �����������
            '  ���������� ���������� ���������� �� �����������
        intCellLimit = gAccessCellLimit
            
            '��������� "���������" ����� � ���� ����������� �����������
            '  ����������, �� ������� ��� ����� �������� ���������� �����
        
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
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
    
            '������������ ����������� ��������� "��������"
        For intCount = 1 To 7 Step 2
            '����
            strCheckingInfo = Trim(strCheckingInfo) + _
            Chr(CByte(CInt(Mid(strDateInfo, intCount, 2))))
        Next
            '����
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strHour, 1, 2))))
            '������
        strCheckingInfo = Trim(strCheckingInfo) + _
        Chr(CByte(CInt(Mid(strMinute, 1, 2))))
            
            '�������� ��������� "��������"
        Call frmTablePerson.Pack(strCheckingInfo)
            
            '��������� ���� "txtInfo"
        txtInfo = Left(strCheckingInfo, 6) + Trim(txtInfo)
        
    End If
            
            '����� ���������-������� ���������������
            '������������� ����
    intAutoRegistrCode = frmTablePerson.AutoRegAccess(txtPersonCode.Text, _
    txtInfo.Text, strStatus, strChecking, strTime)
            
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
        gProtocol.strProtocReserve = "AutoRegAcce " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            '������� (����)����������� ������������� ����
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
            '   ���������� �������� ����������), ��������� ������ �����������
            '   � ���������� ������ �������� ��������� - ������� ��������
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        optTime.Value = True And gTermInp <> -1 Then
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ��������� ����������� ������� �����������
            If frmDemo.cmdOpen(gTermInp).Tag = 0 And Me.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
                frmDemo.imgAccessInData(gTermInp).Enabled = False
                frmDemo.imgAccessOutData(gTermInp).Enabled = False
                frmDemo.imgAccessInfoData(gTermInp).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(gTermInp))))
                frmDemo.cmdOpen(gTermInp).Tag = vntAddr
                frmDemo.cmdOpen(gTermInp).Caption = "Addr=" + CStr(vntAddr)
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(gTermInp).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(gTermInp).Enabled = True
            '����������� ������� ����������� "������"
                Call frmDemo.OpenBarrier(gTermInp)
            '����� ���������-������� ������������� ��� �������
            '������������� ���� - ������ ����� �� �����������
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
            End If
        End If
            
            
            '������� � ��������� ���������
        cmdCancel_Click
            '����� � ��������������� ������������� ���� -
            '   ���������������� �������
    Else
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '��������� ����� (������)
        gProtocol.strProtocStatus = strStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid AutoRegAccess"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)����������� ������������� ����
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
    ((txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
    (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1)) Then
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
            '������� ������ �� (����)����������� ������������� ����
            Me.Tag = 2
            '���������� ����� �� ������ "Cancel _ Exit"
            cmdCancel.SetFocus
            '����� �� ���������
            Exit Sub
        End If
    End If
    
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)����������� ������������� ����
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
            
            '������� ��������� ��������� ���� ������������� ����
    txtPersonCode.Enabled = True
            '������� ����� "Time"
    optTime.Value = True
            '������� ����� "Human"
    optHuman.Value = True
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
    gAccessMoneyCell = gAccessMoneyCellHuman
            '������� ���������� ��������� �������� ���������� �����
    fraPeople.Enabled = True
            '������� ������������ �������� ���������� ����� "DataAccessIn"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    fraDayNight.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
            '����� ������ ��� (�����)
    intAccessTariff = intAccessTariffFull
            '������� ������� ������� - �������� (��� ���������� �����������)
    strTime = "DayNight"
           '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop
            '���������� ����� �� ��������� ���� "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
             
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
            '������� ������������ �������� ���������� ����� "DataAccessIn"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    optTime.Value = True
    txtMoneyDate.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '����� ������ ��� (�����)
    intAccessTariffFull = gAccessDN
            '����� ������ ��� (����)
    intAccessTariffDay = gAccessD
            '����� ������ ��� (����)
    intAccessTariffNight = gAccessN
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '��������� ��������� "������ ����" �� ����� "Human"
Private Sub optHuman_Click()
            
            '������� ����� ����������� ��� �������� (��� ��������� ��������)
    gAccessMoneyCell = gAccessMoneyCellHuman
            
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ���������� ���� "����������"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            
            '���������� ����� �� ������ "OK _ +"
    cmdOK.SetFocus

End Sub

            '��������� ��������� "������ ����" �� ����� "Baby"
Private Sub optBaby_Click()
            
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
    gAccessMoneyCell = gAccessMoneyCellBaby
            
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ���������� ���� "����������"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            
            '���������� ����� �� ������ "OK _ +"
    cmdOK.SetFocus

End Sub

            '��������� ��������� "������ ����" �� ����� "Convoy"
Private Sub optConvoy_Click()
            
            '������� ����� ����������� ��� ������ (��� ��������� ��������)
    gAccessMoneyCell = gAccessMoneyCellConvoy
            
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ���������� ���� "����������"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            
            '���������� ����� �� ������ "OK _ +"
    cmdOK.SetFocus

End Sub

            '��������� ��������� "������ ����" �� ����� "Family"
Private Sub optFamily_Click()
            
            '������� ����� ����������� ��� ����� (��� ��������� ��������)
    gAccessMoneyCell = gAccessMoneyCellFamily
            
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
           
            '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ���������� ���� "����������"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            
            '���������� ����� �� ������ "OK _ +"
    cmdOK.SetFocus

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
            '����������� ����� ����������� "PersonCode"� ���� "Info"
            '  ��� ����������� ���������� ����������
            If gAccessCodeInfo = 1 Then
            '����������� "PersonCode"� ���� "Info"
                txtInfo = Trim(txtPersonCode)
            '������� ��� ���������� ����
                txtInfo.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
                txtInfo.Tag = 1
            End If
            '����� ������������� ���� ������ 16-� ��������
            If Len(Trim(txtPersonCode.Text)) < 16 Then
            '�������� ����������� ���������� ���������� �����
                txtPersonCode.Text = Left("0000000000000000", _
                16 - Len(Trim(txtPersonCode.Text))) + Trim(txtPersonCode.Text)
            End If
            '���������� �������  ��������� � ��������� ���� "PersonCode"
            txtPersonCode.Tag = 1
            '���������� ����� �� ��������� ���� "txtInfo"
            If txtInfo.Enabled = True Then txtInfo.SetFocus
            '��� ����������� ���������� �������
            If (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
            (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1) Then
            
            '���������� ����  ����������� �������
            '���������� ����  ����������� �������
            frmTableCalendar.comCalendar.Today
            intDayReg = frmTableCalendar.comCalendar.Day
            intMonthReg = frmTableCalendar.comCalendar.Month
            intYearReg = frmTableCalendar.comCalendar.Year
            
            '������� ��������� ������� �� ������ "OK _ +"
                cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
                cmdOK.SetFocus
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

            '��������� ������� "PersonCode" ��� ��������������� ����������
            '  ����� ����������� "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
            
             '����� ���������� ����������� ������� �����
    Do While Me.Tag = 0
            '���������� ��������� �������
        DoEvents
    Loop
            '������� ������������ ��� � ���������������
            '  ��������� ����
    txtPersonCode.Text = Trim(vntPersonCode)
            '������� ����������� ��������� ���� ������������� ����
    txtPersonCode.Enabled = False
            '������� ��� ���������� ����
    txtPersonCode.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 1
            '���������� ����� �� ��������� ���� "Info"
    If txtInfo.Enabled = True Then txtInfo.SetFocus
            '����������� ����� ����������� "PersonCode"� ���� "Info"
            '  ��� ����������� ���������� ����������
    If gAccessCodeInfo = 1 Then
            '����������� "PersonCode"� ���� "Info"
        txtInfo = Trim(txtPersonCode)
            '������� ��� ���������� ����
        txtInfo.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
        txtInfo.Tag = 1
    End If
            '������� ����� "Calendar"
    optCalendar.Value = True
            
            '���������� ����  ����������� �������
    frmTableCalendar.comCalendar.Today
    intDayReg = frmTableCalendar.comCalendar.Day
    intMonthReg = frmTableCalendar.comCalendar.Month
    intYearReg = frmTableCalendar.comCalendar.Year
    
End Function

            '��������� ������������ "PersonCode", "Info" � ������
            '  ������ �� �����-����� (+ ����) ��� ��������������� �������
            '  ����� ����������� "Controller" � ������� "DALLAS"
Public Function DallasButton(ByVal strAddrPortType As String, intIndex As Integer)
            '������
Dim strStatus As String
            '������ "��������" ��� �����������
Dim strChecking As String * 8
            '���� (� �����) ����������� �������
Dim strDate As String
            '����� ����������� �������
Dim intHour As Integer
Dim intMinute As Integer
            '��������������� (�� ��� �����) ����� ����������� �������
Dim strHour As String
Dim strMinute As String
            '������� ����������� \ 0 - ����� \ 1 - ����� \ 2 - ���������������
Dim strPersPresent As String * 1
            '������� ("�" - ������������ �����; "D" - ������� ����� �������;
            '  "N" - ������ ����� �������; "������ ������"   - �������� �����
            '  �������)
Dim strExpander As String * 1
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ��������������� � "������� ������"
Dim intAutoRegistrCode  As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '������� ������� - ����� ����������� �������
Static btCount As Byte
            '��������������� (��� �����) ����� ����������� �������
Dim strCount As String
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer
            '���������� ����� �������, � ������� �������� �����������
            '  ���������� ���������� ��������� ���������� �� �����������
Dim intCellLimit As Integer
            '������ ����������� ���������
Dim strMessage As String

    
            '����� ����������� �������
    If btCount < gMaxCount And btCount > gMinCount - 1 Then
        btCount = btCount + CByte(1)
    Else
        btCount = CByte(gMinCount)
    End If
    strCount = Trim(Str(btCount))
    
            '������� ������ "��������" ��� �����������
    strChecking = ""
                '����� ����������� �������
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
            '���� ����������� �������
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
    strDate = gProtocol.strProtocDate
    If Len(Trim(strDate)) = 10 Then
        strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
        Right(Trim(strDate), 4)
    Else
        strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
        Right(Trim(strDate), 4)
    End If
            
            '������������ � ������� ������������ ���
            '  � ��������������� ��������� ����
    txtPersonCode.Text = "0000" + Trim(strCount) + Trim(strHour) + _
    Trim(strMinute) + Left(Trim(strDate), 4) + Right(Trim(strDate), 2)
    
            '���������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 1
            '����������� "PersonCode"� ���� "Info"
    txtInfo = Trim(txtPersonCode)
            '���������� �������  ��������� � ��������� ���� "Info"
    txtInfo.Tag = 1
            '������� ����� "Time"
    optTime.Value = True
                '������� ������� ������� �� ����������� -
            '  �������� (��� ��������� ��������)
    strTime = "DayNight"
            '������� ����������� �������
    strPersPresent = "2"
            '������� �������
    strExpander = "P"
            '��������� ������
    strStatus = "09 - Access/Time"
            
            '������� ����� ����������� ����� ���� ���
            '  ?���������� ���� - ��� ���������� ���������?
    If gAccessMoneyCell = 0 Or gParkInpCellNumb > 0 Then
            '���������� ����� �������, � ������� �������� ����������� ����������
            '  ���������� ���������� �� ����������� (���������� �����
            '  ��� ���������� ����/�����)
        intCellLimit = gAcceInpCellNumb
    
            '��������� "���������" ����� � ���� ����������� ����������,
            '  �� ������� ��� ����� �������� ����-�����
            
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
            intMinute = intMinute + gAccessTimeCell * intCellLimit + gAccessTimeCell
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
            
            '������� ����������� ������� � ������ ��� ����������
    strChecking = Left(strChecking, 6) + strPersPresent + strExpander
            
            '����� ���������-������� ���������������
            '������������� ���� ��� �����������
    intAutoRegistrCode = frmTablePerson.AutoRegAccess(txtPersonCode.Text, _
    txtInfo.Text, strStatus, strChecking, strTime)
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
        gProtocol.strProtocReserve = "AutoRegAcce " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� (����)����������� ������������� ����
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
            
            '���������� ����� ���������� - ������� ��������
        If intError = 0 And frmDemo.chkSetup.Value = 1 Then
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ��������� ����������� ������� �����������
            If frmDemo.cmdOpen(intIndex).Tag = 0 And Me.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
                frmDemo.imgAccessInData(intIndex).Enabled = False
                frmDemo.imgAccessOutData(intIndex).Enabled = False
                frmDemo.imgAccessInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                vntAddr = CByte(CInt(Trim(gAcceAddrTerm(intIndex))))
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
                Call frmDemo.OpenBarrier(intIndex)
            '������������ ����������� ���������
                strMessage = "AcceFreePlaces-1"
            '����� ���������-������� ������������� ��� �������
            '������������� ���� - ������ ����� �� �����������
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorAccess(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
            '�������� ���������
                Call frmDemo.SendMessage(strMessage)
            End If
        End If
        
            '����� � ��������������� ������������� ���� -
            '   ���������������� �������
    Else
            '��������� ����������
        gProtocol.strProtocName = txtInfo.Text
            '��������� ������������ ���
        gProtocol.strProtocPersonCode = txtPersonCode.Text
            '��������� ����� (������)
        gProtocol.strProtocStatus = strStatus
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid AutoRegAccess"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
    
    End If
    
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
        If (Len(Trim(txtInfo.Text)) < 17 And Len(Trim(txtInfo.Text)) > 0 _
        And (gAcceTimeLimit = 0 Or _
        (gAcceTimeLimit > 0 And optCalendar.Value = False))) Or _
        (Len(Trim(txtInfo.Text)) < 11 And Len(Trim(txtInfo.Text)) > 0 _
        And gAcceTimeLimit > 0 And optCalendar.Value = True) Then
            '���������� �������  ��������� � ��������� ���� "Info"
            txtInfo.Tag = 1
            '���������� ����� �� ��������� ���� "PersonCode"
            If txtPersonCode.Enabled = True Then txtPersonCode.SetFocus
            '��� ����������� ���������� �������
            If (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And optCalendar = False) Or _
            (txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1) Then
            '������� ��������� ������� �� ������ "OK _ +"
                 cmdOK.MousePointer = 0
            '���������� ����� �� ������ "�� _+"
                cmdOK.SetFocus
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
            If txtInfo.Enabled = True Then txtInfo.SetFocus
            '������� ����������� ������� �� ������ "OK _ +"
            cmdOK.MousePointer = vbNoDrop
        End If
    End If

End Sub
            
            '������� ����� - "Calendar"
Private Sub optCalendar_Click()
            '������� ��������� ������� ���������� "fraDayNight"
    fraDayNight.Enabled = True
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = True
            '������� ����� "Human"
    optHuman.Value = True
            '������� ������������ ��������� �������� ���������� �����
    fraPeople.Enabled = False
            '������� ���������� ��������� �������� ���������� �����
    lblLat0.Enabled = True
    lblLat320.Enabled = True
    hsbLat.Enabled = True
    txtPersonCode.Enabled = True
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            '������� ����� - "Day"
Private Sub optDay_Click()
            '����� ������ ��� (����)
    intAccessTariff = intAccessTariffDay
            '������� ������� ������� - ������� (��� ���������� �����������)
    strTime = "Day"
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            '������� ����� - "DayNight"
Private Sub optDayNight_Click()
            '����� ������ ��� (�����)
    intAccessTariff = intAccessTariffFull
            '������� ������� ������� - �������� (��� ���������� �����������)
    strTime = "DayNight"
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            '������� ����� - "Night"
Private Sub optNight_Click()
            '����� ������ ��� (����)
    intAccessTariff = intAccessTariffNight
            '������� ������� ������� - ������ (��� ���������� �����������)
    strTime = "Night"
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ����������� ������� �� ������ "OK _ +"
     cmdOK.MousePointer = vbNoDrop

End Sub
            
            '������� ����� - "MoneyFree"
Private Sub optMoneyFree_Click()
            '������� ����������� ������� ���������� "fraDayNight"
    fraDayNight.Enabled = False
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ����� "Human"
    optHuman.Value = True
            '������� ������������ ��������� �������� ���������� �����
    fraPeople.Enabled = False
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
    txtPersonCode.Enabled = True
            '�������� ��������� ����
    txtMoneyDate.Text = ""
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            '�������� ������� ��������� � ��������� ����
    txtMoneyDate.Tag = 0
            '������� ��������� ������� �� ������ "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If

End Sub
            
            '������� ����� - "Time"
Private Sub optTime_Click()
            '������� ����������� ������� ���������� "fraDayNight"
    fraDayNight.Enabled = False
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
            '������� ���������� ��������� �������� ���������� �����
    fraPeople.Enabled = True
            '������� ����� "Human"
    optHuman.Value = True
            '������� ������������ ��������� �������� ���������� �����
    lblLat0.Enabled = False
    lblLat320.Enabled = False
    hsbLat.Enabled = False
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
            
            '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ���������� ���� "����������"
    If Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 99 Then
        txtMoneyDate.Text = Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 9 Then
        txtMoneyDate.Text = "0" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    ElseIf Int(gAccessMoneyCell * gAcceInpCellNumb / 100) > 0 Then
        txtMoneyDate.Text = "00" + Trim(Str(Int(gAccessMoneyCell * _
        gAcceInpCellNumb / 100))) + Mid(txtMoneyDate.Text, 4)
    End If
            '��������� ���������� ���� "����������"
    If gAccessMoneyCell * gAcceInpCellNumb - _
    Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100 > 9 Then
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    Else
        txtMoneyDate.Text = Left(txtMoneyDate.Text, 4) + "0" + _
        Trim(CStr(gAccessMoneyCell * gAcceInpCellNumb - _
        Int(gAccessMoneyCell * gAcceInpCellNumb / 100) * 100)) + _
        Mid(txtMoneyDate.Text, 7)
    End If
            
            '���������� �������  ��������� ����������
    txtMoneyDate.Tag = 1
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
        cmdOK.MousePointer = 0
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
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
            '������� ���������� ����� "MoneyFree" � "Document"
            imgDocument.Enabled = True
            chkDocument.Enabled = True
            imgMoneyFree.Enabled = True
            optMoneyFree.Enabled = True
            '������ ��������
        Else
            '������ �������� ������
            frmDemo.BeepSound
            '������� ������������ ����� "MoneyFree" � "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
            imgMoneyFree.Enabled = False
            optMoneyFree.Enabled = False
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
            '�������� ����� ��������� ����� � �������� � ������������ ���������
            '  (�������� ����� �� ���������� ����� ����� ����)
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
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
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
            '���������� ����
        intAccessDay = Int(intAccessMoney / intAccessTariff)
            
            '��������� "���������" �� ������� ����
        frmTableCalendar.comCalendar.Today
            '���� �� ���� "���������" (�� ���� ����������� �������)
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
        txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '������� ����������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = vbNoDrop
    End If
EndCycle:
            '������� ��� ���������� ����
    txtMoneyDate.BackColor = vbCyan
            '������������� ������ ������ ���
    If Int(intAccessMoney / intAccessTariff) = 0 Then
           '������ ��������� ����������
        txtMoneyDate.Tag = 0
           '����
        txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
           '������������ ������� � ���� "����������"
        txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ����������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = vbNoDrop
    End If
            '��������� (�������� ������ ������ �� ����� ����)
    If intAccessDay > 0 Then
            '���������� ���������� (��� ���������) ����
        intAccessDay = Int(intAccessMoney / intAccessTariff) - intAccessDay
           '�������������� ���������� (��� ���������) ����� ������ � ��������
        intAccessMoney = intAccessDay * intAccessTariff
            '�������������� ����������� ��������� ���������
        hsbSant.Value = intAccessMoney - Int(intAccessMoney / 100) * 100
        hsbLat.Value = Int(intAccessMoney / 100)
        hsbLat_Change
    End If
            '��� ����������� ���������� �������
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If
    
End Sub
