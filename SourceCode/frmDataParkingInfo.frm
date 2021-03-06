VERSION 5.00
Begin VB.Form frmDataParkingInfo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ParkingInfoData"
   ClientHeight    =   4320
   ClientLeft      =   4860
   ClientTop       =   2565
   ClientWidth     =   6960
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
   ScaleHeight     =   4320
   ScaleWidth      =   6960
   Tag             =   "0"
   Visible         =   0   'False
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
      TabIndex        =   13
      Top             =   240
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
         Begin VB.OptionButton optDay 
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   16
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optNight 
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   15
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   18
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
      Begin VB.Image imgCalendar 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingInfo.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingInfo.frx":0802
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgMoneyFree 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingInfo.frx":24A4
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
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   0
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.TextBox txtParkingReg 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5160
      TabIndex        =   12
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtParkingOut 
      Enabled         =   0   'False
      Height          =   288
      Left            =   3720
      TabIndex        =   11
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtParkingIn 
      Enabled         =   0   'False
      Height          =   288
      Left            =   2280
      TabIndex        =   10
      Tag             =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.ListBox lstInfo 
      Height          =   1110
      ItemData        =   "frmDataParkingInfo.frx":28FE
      Left            =   720
      List            =   "frmDataParkingInfo.frx":2900
      TabIndex        =   9
      ToolTipText     =   "Information"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox lstPersonCode 
      Height          =   1110
      ItemData        =   "frmDataParkingInfo.frx":2902
      Left            =   720
      List            =   "frmDataParkingInfo.frx":2904
      TabIndex        =   8
      ToolTipText     =   "PersonCode"
      Top             =   240
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
      TabIndex        =   7
      Top             =   2880
      Width           =   1212
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4680
      TabIndex        =   3
      Tag             =   "0"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox chkDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
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
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Image imgParkingReg 
      Height          =   375
      Left            =   4680
      Picture         =   "frmDataParkingInfo.frx":2906
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgParkingOut 
      Height          =   375
      Left            =   3240
      Picture         =   "frmDataParkingInfo.frx":2D58
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgParkingIn 
      Height          =   375
      Left            =   1800
      Picture         =   "frmDataParkingInfo.frx":2F6A
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   4320
      Y1              =   720
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   4320
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4320
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image imgDocument 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      Picture         =   "frmDataParkingInfo.frx":317C
      Stretch         =   -1  'True
      Top             =   360
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
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   1440
      Y2              =   720
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6720
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
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image imgParkingInfo 
      Height          =   615
      Left            =   1680
      Picture         =   "frmDataParkingInfo.frx":3592
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "frmDataParkingInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������ "��������" ��� �����������
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
    intAutoPressingCode = frmTablePerson.AutoPresParking()
            
            'Ha ����������� ������������ AM, , ������� ������
            '   ���� ������������ ����a�� ����� ������ ��������
    If intAutoPressingCode = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ���������
        If frmDemo.optEnglish = True Then
            MsgBox ("The Cars for Exit  are Present")
        Else
            MsgBox ("Ir Automobils izbrauk.")
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
        gProtocol.strProtocReserve = "Pressing Parking Info "
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ ������ � "������� ������"
        frmDataParkingInfo.Tag = 1
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
        gProtocol.strProtocReserve = "Invalid Press. Parking"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� ������ �� ������ ������ � "������� ������"
        frmDataParkingInfo.Tag = 2
            '������� � ��������� ���������
        cmdCancel_Click
    End If
            
End Sub
            
            '������� � ��������� ��������� (������ "Cancel _ Exit")
Private Sub cmdCancel_Click()
        '������� ������ �� ������ ������ � "������� ������"
    If frmDataParkingInfo.Tag = 0 Then frmDataParkingInfo.Tag = 2
            '������� ��������� ������� �����
    frmDataParkingInfo.Visible = False
            
            '�������� ������ "lstInfo" � "lstPersonCode"
    lstInfo.Clear
    lstPersonCode.Clear
    
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
            
            '������� ������������ �������� ���������� ����� "DataParkingInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ����
    txtParkingIn.Text = "0"
    txtParkingOut.Text = "0"
    txtParkingReg.Text = "0"
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
            '������ ������� ������� �����������
            If Left(Trim(gTablePerson.Text), 2) = "07" Or _
            Left(Trim(gTablePerson.Text), 2) = "05" Or _
            Left(Trim(gTablePerson.Text), 2) = "06" Then
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
            '���������� ����� ���������� �� �����������
                If Mid(Trim(gTablePerson.Text), 7, 1) = "0" Then
                    txtParkingIn.Text = Str(CInt(txtParkingIn.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "1" Then
                    txtParkingOut.Text = Str(CInt(txtParkingOut.Text) + 1)
                ElseIf Mid(Trim(gTablePerson.Text), 7, 1) = "2" Then
                    txtParkingReg.Text = Str(CInt(txtParkingReg.Text) + 1)
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
            '������������� ���� ��� �����������
    intAutoFindCode = frmTablePerson.AutoFindParking(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
    If intAutoFindCode = 0 Then
            
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
            '������� �� �������
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
        If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
        Left(Trim(strStatus), 2) <> "06" Then
            '���� �������� � ��������  ������� ������� ����������� - �� �����
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
            '���������� ������
        If Left(Trim(strStatus), 2) = "07" Then
            '������� ������� ��������������� "������"
            imgMoneyFree.Visible = True
            '���������� ������
        ElseIf Left(Trim(strStatus), 2) = "05" Then
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
            End If
            '��������� ������
        ElseIf Left(Trim(strStatus), 2) = "06" Then
            '������� ������� ��������������� "������"
            imgTime.Visible = True
        End If
             '���������� ���� ���������� ����������� ������� �����
        frmDataParkingInfo.Tag = 1
        Exit Sub
    End If
    
            '����������� ������
UnknownError:
            '������ �������� ������
    frmDemo.BeepSound
    txtMoneyDate.Text = "Unknown Error"
             '���������� ���� ���������� ����������� ������� �����
    frmDataParkingInfo.Tag = 1
    
End Sub

            '������������� ������� �����
Private Sub Form_Deactivate()
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub

            '�������� ������� �����
Private Sub Form_Load()
            '������� ������������ �������� ���������� ����� "DataParkingInfo"
    lblParole.Enabled = False
    imgDocument.Enabled = False
    chkDocument.Enabled = False
            '������� ���������� "������"
    imgMoneyFree.Visible = False
    imgCalendar.Visible = False
    fraDayNight.Visible = False
    imgTime.Visible = False
            '�������� ��������� ����
    txtParkingIn.Text = ""
    txtParkingOut.Text = ""
    txtParkingReg.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
             '����� ��� ���������� ����
    txtParole.BackColor = vbWhite

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '�������� ������� ���������� ������ "Alt"+ {"P", "E" , "^" � "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������
Dim strStatus As String
            '���� � ����� � ������ "Reserve" " ������� ������"
Dim strDate As String
            '��� �������� ��� ���������� � "������� ������"
Dim intAutoFindCode  As Integer
            
            '����� "frmDataParkingInfo" ��������
    If frmDataParkingInfo.Enabled = True Then
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
            '������������� ���� ��� �����������
    intAutoFindCode = frmTablePerson.AutoFindParking(lstPersonCode.Text, _
    lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
    If intAutoFindCode = 0 Then
            '���������� ������ "��������" ��� �����������
            
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
            '������� �� �������
        txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
        Right(Trim(strChecking), 1)
    
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
        If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
        Left(Trim(strStatus), 2) <> "06" Then
            '���� �������� � ��������  ������� ������� ����������� - �� �����
        intButtonsAndIcons = vbOKOnly + vbExclamation
            '������ �������� ������
            frmDemo.BeepSound
            MsgBox "Status Error  !!!", intButtonsAndIcons, "Error"
                GoTo UnknownError
        End If
            '���������� ������
        If Left(Trim(strStatus), 2) = "07" Then
            '������� ������� ��������������� "������"
            imgMoneyFree.Visible = True
            '���������� ������
        ElseIf Left(Trim(strStatus), 2) = "05" Then
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
            End If
            '��������� ������
        ElseIf Left(Trim(strStatus), 2) = "06" Then
            '������� ������� ��������������� "������"
            imgTime.Visible = True
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
            '�������� ��������� ���� "����������" ��� �����������
        txtMoneyDate.Text = ""
            '����� ���������-������� ����������
            '������������� ���� ��� �����������
        intAutoFindCode = frmTablePerson.AutoFindParking(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
        If intAutoFindCode = 0 Then
            
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
            '������� �� �������
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
            If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
            Left(Trim(strStatus), 2) <> "06" Then
            '���� �������� � ��������  ������� ������� ����������� - �� �����
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
            '���������� ������
            If Left(Trim(strStatus), 2) = "07" Then
            '������� ������� ��������������� "������"
                imgMoneyFree.Visible = True
            '���������� ������
            ElseIf Left(Trim(strStatus), 2) = "05" Then
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
                End If
            '��������� ������
            ElseIf Left(Trim(strStatus), 2) = "06" Then
            '������� ������� ��������������� "������"
                imgTime.Visible = True
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
            '�������� ��������� ���� "����������" ��� �����������
        txtMoneyDate.Text = ""
            '����� ���������-������� ����������
            '������������� ���� ��� �����������
        intAutoFindCode = frmTablePerson.AutoFindParking(lstPersonCode.Text, _
        lstInfo.Text, strStatus, strChecking)
            '(����)����� ������������� ���� �������� �������
            '   ���������������� �������
        If intAutoFindCode = 0 Then
            
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
            '������� �������������� ������ �� �������
            txtMoneyDate.Text = Trim(txtMoneyDate.Text) + "/" + _
            Right(Trim(strChecking), 1)
    
            '������ ������� ������� �����������
            
            '������������ ��� ����������� ������ �������
            If Left(Trim(strStatus), 2) <> "07" And Left(Trim(strStatus), 2) <> "05" And _
            Left(Trim(strStatus), 2) <> "06" Then
            '���� �������� � ��������  ������� ������� ����������� - �� �����
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
            '���������� ������
            If Left(Trim(strStatus), 2) = "07" Then
            '������� ������� ��������������� "������"
                imgMoneyFree.Visible = True
            '���������� ������
            ElseIf Left(Trim(strStatus), 2) = "05" Then
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
                End If
            '��������� ������
            ElseIf Left(Trim(strStatus), 2) = "06" Then
            '������� ������� ��������������� "������"
                imgTime.Visible = True
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
