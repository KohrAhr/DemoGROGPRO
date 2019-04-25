VERSION 5.00
Begin VB.Form frmDataAccessInfo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AccessInfoData"
   ClientHeight    =   4305
   ClientLeft      =   4665
   ClientTop       =   2925
   ClientWidth     =   6960
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
   ScaleHeight     =   4305
   ScaleWidth      =   6960
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.CommandButton cmdCleaning 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Outlook"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1680
      TabIndex        =   21
      Top             =   3600
      Width           =   615
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
      Top             =   3600
      Width           =   1212
   End
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   17
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4680
      TabIndex        =   15
      Tag             =   "0"
      Top             =   2400
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
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   1212
   End
   Begin VB.ListBox lstPersonCode 
      Height          =   1110
      ItemData        =   "frmDataAccessInfo.frx":0000
      Left            =   720
      List            =   "frmDataAccessInfo.frx":0002
      TabIndex        =   12
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ListBox lstInfo 
      Height          =   1110
      ItemData        =   "frmDataAccessInfo.frx":0004
      Left            =   720
      List            =   "frmDataAccessInfo.frx":0006
      TabIndex        =   11
      ToolTipText     =   "Information"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtAccessIn 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3120
      TabIndex        =   10
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtAccessOut 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   9
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtAccessReg 
      Enabled         =   0   'False
      Height          =   288
      Left            =   6000
      TabIndex        =   8
      Tag             =   "0"
      Top             =   3840
      Width           =   735
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
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
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
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
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
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1440
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   0
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   1440
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessInfo.frx":0008
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessInfo.frx":0462
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataAccessInfo.frx":2104
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Image imgFamily 
      Height          =   615
      Left            =   6360
      Picture         =   "frmDataAccessInfo.frx":2906
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgConvoy 
      Height          =   615
      Left            =   5520
      Picture         =   "frmDataAccessInfo.frx":2F38
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image imgHuman 
      Height          =   615
      Left            =   4440
      Picture         =   "frmDataAccessInfo.frx":3CBE
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image imgBaby 
      Height          =   615
      Left            =   5040
      Picture         =   "frmDataAccessInfo.frx":44F8
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image imgAccessInfo 
      Height          =   615
      Left            =   1680
      Picture         =   "frmDataAccessInfo.frx":4B36
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   5640
      X2              =   5640
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   1440
      Y2              =   720
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
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataAccessInfo.frx":4F10
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4320
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4320
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
      TabIndex        =   19
      Top             =   1560
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
      TabIndex        =   18
      Top             =   240
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Image imgAccessIn 
      Height          =   375
      Left            =   2640
      Picture         =   "frmDataAccessInfo.frx":5326
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgAccessOut 
      Height          =   375
      Left            =   4080
      Picture         =   "frmDataAccessInfo.frx":576C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgAccessReg 
      Height          =   375
      Left            =   5520
      Picture         =   "frmDataAccessInfo.frx":5BB2
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
End
Attribute VB_Name = "frmDataAccessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������ "��������"
Dim strChecking As String * 8
             '��������� ������
Dim strPassword As String
            '������� ����� ������� "������� ������"
Dim intColNumCorr As Integer

            '������ ������ � "������� ������" (������ "Pressing")
Private Sub cmdOK_Click()
            '��� �������� ��� ������ ������ � "������� ������"
Dim intAutoPressingCode  As Integer

            '����� ���������-������� ������ ������
            '  � "������� ������" ��� �����������
    intAutoPressingCode = frmTablePerson.AutoPresAccess()
            
            'Ha ����������� ������������ ���������� , ������� ������
            '   ���� ������������ ������ ����� ������
    If intAutoPressingCode = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ���������
        If frmDemo.optEnglish = True Then
            MsgBox ("The Persons for Exit  are Present")
        Else
            MsgBox ("Ir Persona izejai")
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
        gProtocol.strProtocReserve = "Pressing AccessInfo "
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
        gProtocol.strProtocReserve = "Invalid Press. Access"
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
            
            '�������� ������ "lstInfo" � "lstPersonCode"
    lstInfo.Clear
    lstPersonCode.Clear
    
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub

            '������� "������� ������" �� ��������� �������� (������ "Cleaning")
Private Sub cmdCleaning_Click()
            '��� �������� ��� ������� "������� ������"
Dim intCleaningCode  As Integer

            '����� ���������-������� ������� "������� ������"
    intCleaningCode = frmTablePerson.CleaningAccess()
            '������� ��������� - ���������������� �������
    If intCleaningCode = 0 Then
            '���������� - �����������
        gProtocol.strProtocName = "CLEANING TabPers"
            '������������ ��� - �����������
        gProtocol.strProtocPersonCode = "CLEANING TabPers"
            '������
        gProtocol.strProtocStatus = "04 - Operator"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Cleaning AccessInfo "
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������� "������� ������"
        Me.Tag = 1
            '������� � ��������� ���������
        cmdCancel_Click
            '������ ���������� - ���������������� �������
    ElseIf intCleaningCode = 1 Then
            '���������� - �����������
        gProtocol.strProtocName = "CLEANING TabPers"
            '������������ ��� - �����������
        gProtocol.strProtocPersonCode = "CLEANING TabPers"
            '������
        gProtocol.strProtocStatus = "04 - Operator"
            '�����
        gProtocol.strProtocTime = Format(Now, "h:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
        gProtocol.strProtocReserve = "Invalid Clean. Access"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ �� ������� "������� ������"
        Me.Tag = 2
            '������� � ��������� ���������
        cmdCancel_Click
    End If
            
End Sub
            
            '����������� ������� �����
Private Sub Form_Activate()
            '������� ����� ������ "������� ������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            
            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������, �������� ��������� ����)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessPlus
            
            '������� ������������ �������� ���������� ����� "DataAccessInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ����
    txtAccessIn.Text = "0"
    txtAccessOut.Text = "0"
    txtAccessReg.Text = "0"
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             '����� ��� ���������� ����
    txtParole.BackColor = vbWhite
    
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
    Else
            '���� �� ���� ��������������� ������� "������� ������"
        For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
            gTablePerson.Row = intRowNum
            '������� - "Status"
            gTablePerson.Col = 2
            '������ ������� ����������
            If Left(Trim(gTablePerson.Text), 2) = "10" Or _
            Left(Trim(gTablePerson.Text), 2) = "08" Or _
            Left(Trim(gTablePerson.Text), 2) = "09" Then
            '������� - "Person or Terminal"
                gTablePerson.Col = 0
            '���������� ������ "lstInfo" �������� �� "������� ������"
                lstInfo.AddItem gTablePerson.Text
            '������� - "PersonCode"
                gTablePerson.Col = 1
            '���������� ������ "lstPersonCode" �������� �� "������� ������"
                lstPersonCode.AddItem gTablePerson.Text
            '������� - "Reserve"
                gTablePerson.Col = 5
            '���������� ����� ���������� � �����������
                If Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
                    txtAccessIn.Text = Str(CInt(txtAccessIn.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
                    txtAccessOut.Text = Str(CInt(txtAccessOut.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
                    txtAccessReg.Text = Str(CInt(txtAccessReg.Text) + 1)
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
            Exit Sub
        End If
            '�������  �������� �������
        lstInfo.ListIndex = 0
        lstPersonCode.ListIndex = 0
    End If
             
            
            '����� ���������-������� ����������
            '������������� ����
    intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
    If intAutoFindCode = 0 Then
            
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
            '������� ����������
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            '������ ������� ����������
            
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
            GoTo UnknownError
        End If
            '���������� ����������
        If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
            imgMoneyFree.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            '���������� ����������
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
            imgCalendar.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
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
            End If
            '��������� ����������
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            '������� ������� ��������������� "������"
            imgTime.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            '��������
            If Mid(lstInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            '����
            ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            '������
            ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            '�����
            ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
        End If
             '���������� ���� ���������� ����������� ������� �����
        Me.Tag = 1
        Exit Sub
    End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
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
            '������� ������������ �������� ���������� ����� "DataAccessInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ����
    txtAccessIn.Text = ""
    txtAccessOut.Text = ""
    txtAccessReg.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             '����� ��� ���������� ����
    txtParole.BackColor = vbWhite

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
            
            '�������� ������� ���������� ������ "Alt"+ {"P", "E", "^" � "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            
            '������� ����� ��������
    If Me.Enabled = True Then
            '������������ "������" ���� �� ������ "P"
        If KeyCode = 80 And Shift = 4 Then
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
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ���� "����������" ��� �����������
    txtMoneyDate.Text = ""
            '����� ���������-������� ����������
            '������������� ����
    intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
    If intAutoFindCode = 0 Then
            
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
            '������� ����������
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            '������ ������� ����������
            
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
            GoTo UnknownError
        End If
            '���������� ����������
        If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
            imgMoneyFree.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            '���������� ����������
        ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
            imgCalendar.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
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
            End If
            '��������� ����������
        ElseIf Left(Trim(strStatus), 2) = "09" Then
            '������� ������� ��������������� "������"
            imgTime.Visible = True
            '������� ���������� ��������������� "������"
            imgHuman.Visible = False
            imgBaby.Visible = False
            imgConvoy.Visible = False
            imgFamily.Visible = False
            '��������
            If Mid(lstInfo.Text, 5, 1) = "1" Then
                imgHuman.Visible = True
            '����
            ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                imgBaby.Visible = True
            '������
            ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                imgConvoy.Visible = True
            '�����
            ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                imgFamily.Visible = True
            End If
        End If
        Exit Sub
    End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"

End Sub

            '����� ������ "������� ������" ��� "������" �� ������ "PersonCode"
Private Sub lstPersonCode_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            '������� ������
Dim strWork As String
            '������� ����������
Dim intWork As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� ������ "������� ������"
        lstInfo.ListIndex = lstPersonCode.ListIndex
            '������� ���������� "������"
        imgMoneyFree.Visible = False
        imgCalendar.Visible = False
        fraDayNight.Visible = False
        imgTime.Visible = False
            '�������� ��������� ���� "����������"
        txtMoneyDate.Text = ""
            '����� ���������-������� ����������
            '������������� ����
        intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
        If intAutoFindCode = 0 Then
            
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
            '������� ����������
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
            '������ ������� ����������
            
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
                GoTo UnknownError
            End If
            '���������� ����������
            If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
                imgMoneyFree.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            '���������� ����������
            ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
                imgCalendar.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
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
                End If
            '��������� ����������
            ElseIf Left(Trim(strStatus), 2) = "09" Then
            '������� ������� ��������������� "������"
                imgTime.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            '��������
                If Mid(lstInfo.Text, 5, 1) = "1" Then
                    imgHuman.Visible = True
            '����
                ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                    imgBaby.Visible = True
            '������
                ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                    imgConvoy.Visible = True
            '�����
                ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                    imgFamily.Visible = True
                End If
            End If
            
            '����� "������ ���������" �����������
            If chkDocument.Value = 1 Then
            '������ ��������� (�������� �� �����-�����, ��������
            '  ���� �/��� ��������� ����)
                Call frmDemo.PrintDocument(gProtocol.strProtocName, _
                gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
                gProtocol.strProtocTime, gProtocol.strProtocDate, _
                gProtocol.strProtocReserve, intError)
            End If
            
            Exit Sub
        End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
    End If

End Sub
            
            '����� ������ "������� ������" ��� "������" �� ������ "Info"
Private Sub lstInfo_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '������� �������
Dim intCount As Integer
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            '������� ������
Dim strWork As String
            '������� ����������
Dim intWork As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� ������ "������� ������"
        lstPersonCode.ListIndex = lstInfo.ListIndex
            '������� ���������� "������"
        imgMoneyFree.Visible = False
        imgCalendar.Visible = False
        fraDayNight.Visible = False
        imgTime.Visible = False
            '�������� ��������� ���� "����������"
        txtMoneyDate.Text = ""
            '����� ���������-������� ����������
            '������������� ����
        intAutoFindCode = frmTablePerson.AutoFindAccess(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
        If intAutoFindCode = 0 Then
            
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
            '������� �������������� ������ ����������
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
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
                GoTo UnknownError
            End If
            '���������� ����������
            If Left(Trim(strStatus), 2) = "10" Then
            '������� ������� ��������������� "������"
                imgMoneyFree.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            '���������� ����������
            ElseIf Left(Trim(strStatus), 2) = "08" Then
            '������� ������� ��������������� "������"
                imgCalendar.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
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
                End If
            '��������� ����������
            ElseIf Left(Trim(strStatus), 2) = "09" Then
            '������� ������� ��������������� "������"
                imgTime.Visible = True
            '������� ���������� ��������������� "������"
                imgHuman.Visible = False
                imgBaby.Visible = False
                imgConvoy.Visible = False
                imgFamily.Visible = False
            '��������
                If Mid(lstInfo.Text, 5, 1) = "1" Then
                    imgHuman.Visible = True
            '����
                ElseIf Mid(lstInfo.Text, 5, 1) = "2" Then
                    imgBaby.Visible = True
            '������
                ElseIf Mid(lstInfo.Text, 5, 1) = "3" Then
                    imgConvoy.Visible = True
            '�����
                ElseIf Mid(lstInfo.Text, 5, 1) = "4" Then
                    imgFamily.Visible = True
                End If
            End If
            
            '����� "������ ���������" �����������
            If chkDocument.Value = 1 Then
            '������ ��������� (�������� �� �����-�����, ��������
            '  ���� �/��� ��������� ����)
                Call frmDemo.PrintDocument(gProtocol.strProtocName, _
                gProtocol.strProtocPersonCode, gProtocol.strProtocStatus, _
                gProtocol.strProtocTime, gProtocol.strProtocDate, _
                gProtocol.strProtocReserve, intError)
            End If
            
            Exit Sub
        End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
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
            '������� ��������� ����� "Document"
            imgDocument.Enabled = True
            chkDocument.Enabled = True
            '������ ��������
        Else
            '������ �������� ������
            frmDemo.BeepSound
            '������� ����������� ����� "Document"
            imgDocument.Enabled = False
            chkDocument.Enabled = False
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


