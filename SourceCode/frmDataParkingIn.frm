VERSION 5.00
Begin VB.Form frmDataParkingIn 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ParkingInData"
   ClientHeight    =   3960
   ClientLeft      =   4860
   ClientTop       =   2745
   ClientWidth     =   6990
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
   ScaleHeight     =   3960
   ScaleWidth      =   6990
   Tag             =   "0"
   Visible         =   0   'False
   Begin VB.VScrollBar vsbDate 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4680
      Max             =   366
      TabIndex        =   31
      Top             =   1680
      Width           =   255
   End
   Begin VB.Frame fraMonth 
      Caption         =   " D   1/2M  1M   2M"
      Height          =   615
      Left            =   5040
      TabIndex        =   26
      Top             =   1680
      Width           =   1695
      Begin VB.OptionButton optNot 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optHalf 
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optOne 
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optTwo 
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Timer tmrParoleTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   5880
      Top             =   240
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
      TabIndex        =   17
      Top             =   3240
      Width           =   1212
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
   Begin VB.TextBox txtParole 
      Height          =   324
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "*"
      TabIndex        =   14
      Tag             =   """"""
      ToolTipText     =   "Password "
      Top             =   960
      Width           =   972
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
      TabIndex        =   10
      Top             =   360
      Width           =   1455
      Begin VB.Frame fraDayNight 
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
         Begin VB.OptionButton optNight 
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optDayNight 
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   600
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDay 
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lblNight 
            Alignment       =   2  'Center
            Caption         =   "N"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblDayNight 
            Alignment       =   2  'Center
            Caption         =   "DN"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lbDay 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
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
         TabIndex        =   13
         Top             =   600
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
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
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
         TabIndex        =   11
         Top             =   3000
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   1440
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
         Picture         =   "frmDataParkingIn.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   375
      End
      Begin VB.Image imgTime 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingIn.frx":045A
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image imgCalendar 
         Height          =   375
         Left            =   120
         Picture         =   "frmDataParkingIn.frx":20FC
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   375
      End
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
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.HScrollBar hsbLat 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   320
      TabIndex        =   4
      Top             =   2400
      Width           =   1452
   End
   Begin VB.HScrollBar hsbSant 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4800
      Max             =   99
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.TextBox txtInfo 
      Height          =   288
      Left            =   720
      TabIndex        =   2
      Tag             =   "0"
      ToolTipText     =   "Information"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtPersonCode 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "PersonCode"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtMoneyDate 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4560
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Money and Date"
      Top             =   2760
      Width           =   1935
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
      Left            =   4080
      TabIndex        =   32
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMoneyDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Ls"
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
      TabIndex        =   18
      Top             =   2400
      Width           =   375
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4680
      X2              =   5640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image imgParkingIn 
      Height          =   495
      Left            =   1680
      Picture         =   "frmDataParkingIn.frx":28FE
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
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
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1560
      Y2              =   3120
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
      X1              =   4080
      X2              =   4680
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
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5640
      X2              =   6840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   720
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6840
      X2              =   4080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   6840
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
      Left            =   4800
      Picture         =   "frmDataParkingIn.frx":2B10
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
      TabIndex        =   15
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblLat0 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      TabIndex        =   9
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label lblLat320 
      Alignment       =   2  'Center
      Caption         =   "320"
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
      TabIndex        =   8
      Top             =   2400
      Width           =   375
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
      TabIndex        =   7
      Top             =   1920
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
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmDataParkingIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
            '������� ������� ������� � ����������� -
            '   (��� ���������� ��������)
Dim strTime As String
            '������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ������� "������� ���������"
Dim intColNum As Integer
            '����� ������� ��������� ������� � ������
Dim intPosNum As Integer
             '��������� ������
Dim strPassword As String

            '�������� ������� ���������� ������ "Alt"+ {"+" � "E"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '����� "frmDataParkingIn" ��������
    If frmDataParkingIn.Enabled = True Then
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
            '������ "��������" ��� �����������
Dim strChecking As String * 8
            '��������� "��������" ���� "txtInfo"
Dim strCheckingInfo As String * 8
            '���� (� �����) ����������� ���������� ���
            '  ���� ���������� ����������� ���
Dim strDate As String
            '��������� "����������� ���� � �����" ���� "txtInfo"
Dim strDateInfo As String
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
            '��� �������� ��� ��������������� � "������� ������"
Dim intAutoRegistrCode  As Integer
            '����� �������� � ������� "������� ����������", ��������� ������
Dim intRequest As Integer
            '����� �����������
Dim vntAddr As Variant
            '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
Dim intCellLimit As Integer
            '������� ������ ��� ������ �����-���� � ��.
Dim intError As Integer

            '����������� ������� �� ������ "OK _ +"
    If cmdOK.MousePointer = vbNoDrop Then Exit Sub
            
            '������� ������ � ��������� "��������"
    strChecking = ""
    strCheckingInfo = ""
            '��������� ����� ����������� �������
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
    strDate = Left(Trim(gProtocol.strProtocDate), 2) + _
    Mid(Trim(gProtocol.strProtocDate), 4, 2) + _
    Right(Trim(gProtocol.strProtocDate), 4)
    strDateInfo = strDate
            
            '������� ����������� �� �������
    strCarPresent = "2"
            '������� �� �������
    strExpander = "P"
            '������ ������� ������� �����������
    If optMoneyFree.Value = True Then
            '���������� ������
        strStatus = "07 - Parking/Free"
    ElseIf optCalendar.Value = True Then
            '���������� ������
        strStatus = "05 - Parking/Calen."
            '������ � ������� ������� ������� �� �����������
        If optCalendar.Value = True And optDay.Value = True Then
            strExpander = "D"
            '������ � ������ ������� ������� �� �����������
        ElseIf optCalendar.Value = True And optNight.Value = True Then
            strExpander = "N"
        End If
            '���� ���������� ����������� ������������ ���
        strDate = Mid(Trim(txtMoneyDate.Text), 11)
        If Len(Trim(strDate)) = 10 Then
            strDate = Left(Trim(strDate), 2) + Mid(Trim(strDate), 4, 2) + _
            Right(Trim(strDate), 4)
        Else
            strDate = "0" + Left(Trim(strDate), 1) + Mid(Trim(strDate), 3, 2) + _
            Right(Trim(strDate), 4)
        End If
    ElseIf optTime.Value = True Then
            '��������� ������
        strStatus = "06 - Parking/Time"
    
            '�������� ����� ����������� ����� ���� ���
            '  ?���������� ������ - ��� ���������� ���������?
        If gParkingMoneyCell = 0 Or gParkInpCellNumb > 1 Then
            '���������� ����� �������, � ������� �������� ����������� ����������
            '  ������� ���������� �� ����������� (���������� ������
            '  ��� ���������� �����/�����)
            intCellLimit = gParkInpCellNumb
    
            '��������� "���������" ����� � ���� ����������� �������,
            '  �� ������� ��� ����� �������� �����-�����
            
            '��������� ������� ����
            If (intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell) > 59 Then
                If (gParkingTimeCell * intCellLimit + gParkingTimeCell) > 1440 Then
                    intHour = intHour + Int((intMinute + _
                    gParkingTimeCell * intCellLimit) / 60)
                    intMinute = intMinute + gParkingTimeCell * intCellLimit - _
                    Int((intMinute + gParkingTimeCell * intCellLimit) / 60) * 60
                Else
                    intHour = intHour + Int((intMinute + gParkingTimeCell + _
                    gParkingTimeCell * intCellLimit) / 60)
                    intMinute = intMinute + gParkingTimeCell + _
                    gParkingTimeCell * intCellLimit - _
                    Int((intMinute + gParkingTimeCell + _
                    gParkingTimeCell * intCellLimit) / 60) * 60
                End If
            '��������� ������� ����
                If intHour >= 24 Then
                    intHour = intHour - 24
            
            '��������� ������� "������� ���������" (� ������� ����)
                    If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            '��������� ������ "������� ���������" (� ������� ����)
                        If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            '������ �������� ������
                            frmDemo.BeepSound
                            If frmDemo.optEnglish = True Then
                                MsgBox ("New Year TableCalendar Error")
                            Else
                                MsgBox ("Jauna gada kalend. nesask.")
                            End If
            '�������� ��������� ��������� ������ � ������� "������� ���������"
            
            '�� ��������� ������ "������� ���������" (� ������� ����)
                        Else
            '������� ������ "������� ���������"+1 (��������� ����)
                            frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            '������� ������� "������� ���������" =1 (��������� ����)
                            frmTableCalendar.grdTableCalendar.Col = 1
                        End If
            '�� ��������� ������� "������� ���������" (� ������� ����)
                    Else
            '������� ������ "������� ���������" (������� ����)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum
            '������� ������� "������� ���������" +1 (��������� ����)
                        frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                    End If
                
            '������� �������� ����� � ������� ������ "������� ���������"
                    intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                    If intPosNum <> 0 Then
            '���������  �����
                        If intPosNum = 3 Then
                            strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                            Right(strDate, 6)
                        Else
                            strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                            Right(strDate, 6)
                        End If
                    Else
            '���������  �����
                        If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                            strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                            Right(strDate, 6)
                        Else
                            strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                            Right(strDate, 6)
                        End If
                    End If
            '������� ������
                    If Left(strDate, 2) = "01" Then
            '���������  ������
                        If CInt(Mid(strDate, 3, 2)) + 1 > 9 Then
                            strDate = "01" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                            Right(strDate, 4)
                        Else
                            strDateInfo = "01" + "0" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                            Right(strDate, 4)
                        End If
            '������� ����
                        If Mid(strDate, 3, 2) = "13" Then
            '���������  ������ � ����
                            strDate = "01" + "01" + Trim(Str(CInt(Right(strDateInfo, 4)) + 1))
                        End If
                    End If
                End If
            '�� ��������� ������� ����
            Else
                intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            
            '������� ����������� �� ������� � ������ ��� ����������
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            '���������� ������ �� ����������� � ������������ �������
            '  ������������ ���������� �� �� �����������
    If gParkTimeLimit > 0 And optCalendar.Value = True Then
            '���������� ����� �������, � ������� �������� �����������
            '  �� ����������� ������� ���������� ���������� �� �����������
        intCellLimit = gParkingCellLimit
            
            '��������� "���������" ����� � ���� ����������� �����������
            '  �������, �� ������� ��� ����� �������� ���������� �����
        
            '��������� ������� ����
        If (intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell) > 59 Then
            If (gParkingTimeCell * intCellLimit + gParkingTimeCell) > 1440 Then
                intHour = intHour + Int((intMinute + _
                gParkingTimeCell * intCellLimit) / 60)
                intMinute = intMinute + gParkingTimeCell * intCellLimit - _
                Int((intMinute + gParkingTimeCell * intCellLimit) / 60) * 60
            Else
                intHour = intHour + Int((intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit) / 60)
                intMinute = intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit - _
                Int((intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit) / 60) * 60
            End If
            '��������� ������� ����
            If intHour >= 24 Then
                intHour = intHour - 24
            
            '��������� ������� "������� ���������" (� ������� ����)
                If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            '��������� ������ "������� ���������" (� ������� ����)
                    If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            '������ �������� ������
                        frmDemo.BeepSound
                        If frmDemo.optEnglish = True Then
                            MsgBox ("New Year TableCalendar Error")
                        Else
                            MsgBox ("Jauna gada kalend. nesask.")
                        End If
            '�������� ��������� ��������� ������ � ������� "������� ���������"
            
            '�� ��������� ������ "������� ���������" (� ������� ����)
                    Else
            '������� ������ "������� ���������"+1 (��������� ����)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            '������� ������� "������� ���������" =1 (��������� ����)
                        frmTableCalendar.grdTableCalendar.Col = 1
                    End If
            '�� ��������� ������� "������� ���������" (� ������� ����)
                Else
            '������� ������ "������� ���������" (������� ����)
                    frmTableCalendar.grdTableCalendar.Row = gRowNum
            '������� ������� "������� ���������" +1 (��������� ����)
                    frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                End If
                
            '������� �������� ����� � ������� ������ "������� ���������"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            '���������  �����
                    If intPosNum = 3 Then
                        strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDateInfo, 6)
                    Else
                        strDateInfo = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDateInfo, 6)
                    End If
                Else
            '���������  �����
                    If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                        strDateInfo = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDateInfo, 6)
                    Else
                        strDateInfo = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDateInfo, 6)
                    End If
                End If
            '������� ������
                If Left(strDateInfo, 2) = "01" Then
            '���������  ������
                    If CInt(Mid(strDateInfo, 3, 2)) + 1 > 9 Then
                        strDateInfo = "01" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                        Right(strDateInfo, 4)
                    Else
                        strDateInfo = "01" + "0" + Trim(Str(CInt(Mid(strDateInfo, 3, 2)) + 1)) + _
                        Right(strDateInfo, 4)
                    End If
            '������� ����
                    If Mid(strDateInfo, 3, 2) = "13" Then
            '���������  ������ � ����
                        strDateInfo = "01" + "01" + Trim(Str(CInt(Right(strDateInfo, 4)) + 1))
                    End If
                End If
            End If
            '�� ��������� ������� ����
        Else
            intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            '������������� ���� ��� �����������
    intAutoRegistrCode = frmTablePerson.AutoRegParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoRegPark " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '��������� � ��������� ����� ������� �����
            '   ��������� � "������� ������"
        txtPersonCode.Tag = 0
        txtInfo.Tag = 0
        txtMoneyDate.Tag = 0
            '������� (����)����������� ������������� ����
        frmDataParkingIn.Tag = 1
            
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
            '   � ���������� ������ ��������� ��������� - ������� ��������
        If intError = 0 And gTimeShare = 1 And frmDemo.chkSetup.Value = 1 And _
        optTime.Value = True And gTermInp <> -1 Then
            ' ����������� "������" �� ������ ����� "Controller'a",
            '  ��������� ����������� ������� �����������
            If frmDemo.cmdOpen(gTermInp).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
                frmDemo.imgParkingInData(gTermInp).Enabled = False
                frmDemo.imgParkingOutData(gTermInp).Enabled = False
                frmDemo.imgParkingInfoData(gTermInp).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                vntAddr = CByte(CInt(Trim(gParkAddrTerm(gTermInp))))
                frmDemo.cmdOpen(gTermInp).Tag = vntAddr
                frmDemo.cmdOpen(gTermInp).Caption = "Addr=" + CStr(vntAddr)
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(gTermInp).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(gTermInp).Enabled = True
            '����������� ������� ����������� "������"
                Call frmDemo.OpenBarrier(gTermInp)
            '����� ���������-������� ������������� ��� �������
            '������������� ���� - ���������� ������ �� �����������
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "Invalid AutoRegParking"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '����� ��� ���������� ����
        txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)����������� ������������� ����
        frmDataParkingIn.Tag = 2
    
    
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
    If frmDataParkingIn.Tag = 1 And _
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
            '����� �� ���������
            Exit Sub
        End If
    End If
    
            '����� ��� ���������� ����
    txtMoneyDate.BackColor = vbWhite
            '������� ������ �� (����)����������� ������������� ����
    If frmDataParkingIn.Tag = 0 Then frmDataParkingIn.Tag = 2
            '������� ��������� ������� �����
    frmDataParkingIn.Visible = False
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
            '������� ����� "Not"
    optNot.Value = True
            '������� ����� "Time"
    optTime.Value = True
            '������� ������������ �������� ���������� ����� "DataParkingIn"
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    fraDayNight.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    txtMoneyDate.Enabled = False
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
             '����� ��� ���������� ����
    txtPersonCode.BackColor = vbWhite
    txtInfo.BackColor = vbWhite
    txtParole.BackColor = vbWhite
           '����
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
            '������������ ������� � ���� "����������"
    txtMoneyDate.Text = "000,00 Ls=" + Trim(txtMoneyDate.Text)
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '����� ������ ������������ ��� (�����)
    intParkingTariff = intParkingTariffFull
            '������� ������� ������� � ����������� -
            '  �������� (��� ���������� ��������)
    strTime = "DayNight"
           '������� ����������� ������� �� ������ "OK _ +"
    cmdOK.MousePointer = vbNoDrop
            '���������� ����� �� ��������� ���� "txtPersonCode"
    If txtPersonCode.Visible = True Then txtPersonCode.SetFocus
             '���������� ���� ���������� ����������� ������� �����
    frmDataParkingIn.Tag = 1

End Sub

            '������������� ������� �����
Private Sub Form_Deactivate()
            '��������� ������� �������� ��������
            '  ����������� "������� ������"
    gTablePerson.AccessMinus

End Sub
            
            '�������� ������� �����
Private Sub Form_Load()
            '������� ������������ �������� ���������� ����� "DataParkingIn"
    imgDocument.Enabled = False
    chkDocument.Enabled = False
    imgMoneyFree.Enabled = False
    optMoneyFree.Enabled = False
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    optNot.Value = True
    optTime.Value = True
    txtMoneyDate.Enabled = False
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
            '�������� ��������� ����
    txtPersonCode.Text = ""
    txtInfo.Text = ""
    txtParole.Text = ""
    txtMoneyDate.Text = ""
            '��������� ��������� ��������� ��������� ��� ����� ���������
    hsbSant.Tag = 0
    hsbLat.Tag = 0
    vsbDate.Tag = 0
            '�������� ������ ���������
    hsbSant.Value = 0
    hsbLat.Value = 0
    vsbDate.Value = 0
            '�������� �������� ��������� � ��������� �����
    txtPersonCode.Tag = 0
    txtInfo.Tag = 0
    txtMoneyDate.Tag = 0
            '����� ������ ������������ ��� (�����)
    intParkingTariffFull = gParkingDN
            '����� ������ ������������ ��� (����)
    intParkingTariffDay = gParkingD
            '����� ������ ������������ ��� (����)
    intParkingTariffNight = gParkingN
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
            '����������� ����� ����������� "PersonCode"� ���� "Info"
            '  ��� ����������� ���������� ����������
            If gParkingCodeInfo = 1 Then
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

            '��������� ������� "PersonCode" ��� ��������������� �������
            '  ����������� ����� ����������� "Controller"
Public Function Analysis(ByVal vntPersonCode As Variant)
             '����� ���������� ����������� ������� �����
    Do While frmDataParkingIn.Tag = 0
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
    If gParkingCodeInfo = 1 Then
            '����������� "PersonCode"� ���� "Info"
        txtInfo = Trim(txtPersonCode)
            '������� ��� ���������� ����
        txtInfo.BackColor = vbCyan
            '���������� �������  ��������� � ��������� ���� "PersonCode"
        txtInfo.Tag = 1
    End If
            '������� ����� "Calendar"
    optCalendar.Value = True
    
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
            '������� ����������� �� \ 0 - ������ \ 1 - ������ \ 2 - ���������������
Dim strCarPresent As String * 1
            '������� �� ("�" - ������������ ������; "D" - ������� ����� �������
            '  �� �����������; "N" - ������ ����� ������� �� �����������; "������
            '  ������"   - �������� ����� ������� �� �����������)
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
            '  �� ���������� ������� ��������� ���������� �� �����������
Dim intCellLimit As Integer
            '������ ����������� ���������
Dim strMessage As String
    
            '����� ����������� �������
    If btCount < 99 And btCount > 9 Then
        btCount = btCount + CByte(1)
    Else
        btCount = CByte(10)
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
    txtPersonCode.Text = "000000" + Trim(strCount) + Trim(strHour) + _
    Trim(strMinute) + Left(Trim(strDate), 4)
    
            '���������� �������  ��������� � ��������� ���� "PersonCode"
    txtPersonCode.Tag = 1
            '����������� "PersonCode"� ���� "Info"
    txtInfo = Trim(txtPersonCode)
            '���������� �������  ��������� � ��������� ���� "Info"
    txtInfo.Tag = 1
            '������� ����� "Time"
    optTime.Value = True
                '������� ������� ������� � ����������� -
            '  �������� (��� ��������� ��������)
    strTime = "DayNight"
            '������� ����������� �� �������
    strCarPresent = "2"
            '������� �� �������
    strExpander = "P"
            '��������� ������
    strStatus = "06 - Parking/Time"
    
            '�������� ����� ����������� ����� ���� ���
            '  ?���������� ������ - ��� ���������� ���������?
    If gParkingMoneyCell = 0 Or gParkInpCellNumb > 0 Then
            '���������� ����� �������, � ������� �������� ����������� ����������
            '  ������� ���������� �� ����������� (���������� ������
            '  ��� ���������� �����/�����)
        intCellLimit = gParkInpCellNumb
    
            '��������� "���������" ����� � ���� ����������� �������,
            '  �� ������� ��� ����� �������� �����-�����
            
            '��������� ������� ����
        If (intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell) > 59 Then
            If (gParkingTimeCell * intCellLimit + gParkingTimeCell) > 1440 Then
                intHour = intHour + Int((intMinute + _
                gParkingTimeCell * intCellLimit) / 60)
                intMinute = intMinute + gParkingTimeCell * intCellLimit - _
                Int((intMinute + gParkingTimeCell * intCellLimit) / 60) * 60
            Else
                intHour = intHour + Int((intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit) / 60)
                intMinute = intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit - _
                Int((intMinute + gParkingTimeCell + _
                gParkingTimeCell * intCellLimit) / 60) * 60
            End If
            '��������� ������� ����
            If intHour >= 24 Then
                intHour = intHour - 24
            
            '��������� ������� "������� ���������" (� ������� ����)
                If gColNum = frmTableCalendar.grdTableCalendar.Cols - 1 Then
            '��������� ������ "������� ���������" (� ������� ����)
                    If gRowNum = frmTableCalendar.grdTableCalendar.Rows - 1 Then
            '������ �������� ������
                        frmDemo.BeepSound
                        If frmDemo.optEnglish = True Then
                            MsgBox ("New Year TableCalendar Error")
                        Else
                            MsgBox ("Jauna gada kalend. nesask.")
                        End If
            '�������� ��������� ��������� ������ � ������� "������� ���������"
                
            '�� ��������� ������ "������� ���������" (� ������� ����)
                    Else
            '������� ������ "������� ���������"+1 (��������� ����)
                        frmTableCalendar.grdTableCalendar.Row = gRowNum + 1
            '������� ������� "������� ���������" =1 (��������� ����)
                        frmTableCalendar.grdTableCalendar.Col = 1
                    End If
            '�� ��������� ������� "������� ���������" (� ������� ����)
                Else
            '������� ������ "������� ���������" (������� ����)
                    frmTableCalendar.grdTableCalendar.Row = gRowNum
            '������� ������� "������� ���������" +1 (��������� ����)
                    frmTableCalendar.grdTableCalendar.Col = gColNum + 1
                End If
                
            '������� �������� ����� � ������� ������ "������� ���������"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            '���������  �����
                    If intPosNum = 3 Then
                        strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDate, 6)
                    Else
                        strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDate, 6)
                    End If
                Else
            '���������  �����
                    If Len(Trim(frmTableCalendar.grdTableCalendar.Text)) = 2 Then
                        strDate = Left(Trim(frmTableCalendar.grdTableCalendar.Text), 2) + _
                        Right(strDate, 6)
                    Else
                        strDate = "0" + Left(Trim(frmTableCalendar.grdTableCalendar.Text), 1) + _
                        Right(strDate, 6)
                    End If
                End If
            '������� ������
                If Left(strDate, 2) = "01" Then
            '���������  ������
                    If CInt(Mid(strDate, 3, 2)) + 1 > 9 Then
                        strDate = "01" + Trim(Str(CInt(Mid(strDate, 3, 2)) + 1)) + _
                        Right(strDate, 4)
                    Else
                        strDate = "01" + "0" + Trim(Str(CInt(Mid(strDate, 3, 2)) + 1)) + _
                        Right(strDate, 4)
                    End If
            '������� ����
                    If Mid(strDate, 3, 2) = "13" Then
            '���������  ������ � ����
                        strDate = "01" + "01" + Trim(Str(CInt(Right(strDate, 4)) + 1))
                    End If
                End If
            End If
            '�� ��������� ������� ����
        Else
            intMinute = intMinute + gParkingTimeCell * intCellLimit + gParkingTimeCell
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
            
            '������� ����������� �� ������� � ������ ��� ����������
    strChecking = Left(strChecking, 6) + strCarPresent + strExpander
            
            '����� ���������-������� ���������������
            '������������� ���� ��� �����������
    intAutoRegistrCode = frmTablePerson.AutoRegParking(txtPersonCode.Text, _
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
        gProtocol.strProtocReserve = "AutoRegPark " + Left(Trim(txtMoneyDate.Text), 9)
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
            '������� (����)����������� ������������� ����
        frmDataParkingIn.Tag = 1
            
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
            If frmDemo.cmdOpen(intIndex).Tag = 0 And frmDataParkingIn.Tag = 1 Then
            '������� ������������ �������� ���������� (�����������
            '  � ���������� ��������, ����������) ��� ��������� �����������
                frmDemo.imgParkingInData(intIndex).Enabled = False
                frmDemo.imgParkingOutData(intIndex).Enabled = False
                frmDemo.imgParkingInfoData(intIndex).Enabled = False
            '�������� � ����������� "������" ����� "Controller'a",
            '  ���������� ������� ������������� �������� ���������
                vntAddr = CByte(CInt(Trim(gParkAddrTerm(intIndex))))
                frmDemo.cmdOpen(intIndex).Tag = vntAddr
                frmDemo.cmdOpen(intIndex).Caption = "Addr=" + CStr(vntAddr)
            '����� "N_?" - (������� ���)
                frmDemo.lblInform(intIndex).BackColor = vbGreen
            '�������� �������� "TimeOut" ����������� "������"
                frmDemo.tmrButton(intIndex).Enabled = True
            '����������� ������� ����������� "������"
                Call frmDemo.OpenBarrier(intIndex)
            '������������ ����������� ���������
                strMessage = "ParkFreePlaces-1"
            '�������� ���������
                Call frmDemo.SendMessage(strMessage)
            '����� ���������-������� ������������� ��� �������
            '������������� ���� - ���������� ������ �� �����������
                strChecking = Left(Trim(strChecking), 6) + "0" + Right(Trim(strChecking), 1)
                Call frmTablePerson.AutoCorParking(txtPersonCode.Text, _
                txtInfo.Text, strChecking, strStatus)
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
        gProtocol.strProtocReserve = "Invalid AutoRegParking"
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
        And (gParkTimeLimit = 0 Or _
        (gParkTimeLimit > 0 And optCalendar.Value = False))) Or _
        (Len(Trim(txtInfo.Text)) < 11 And Len(Trim(txtInfo.Text)) > 0 _
        And gParkTimeLimit > 0 And optCalendar.Value = True) Then
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
            '���������� � ������������ ���������
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
            
            '������� ����� - "Day"
Private Sub optDay_Click()
            '����� ������ ������������ ��� (����)
    intParkingTariff = intParkingTariffDay
            '������� ������� ������� � ����������� -
            '  ������� (��� ���������� ��������)
    strTime = "Day"
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
            
            '������� ����� - "DayNight"
Private Sub optDayNight_Click()
            '����� ������ ������������ ��� (�����)
    intParkingTariff = intParkingTariffFull
            '������� ������� ������� � ����������� -
            '  �������� (��� ���������� ��������)
    strTime = "DayNight"
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
            
            '������� ����� - "Night"
Private Sub optNight_Click()
            '����� ������ ������������ ��� (����)
    intParkingTariff = intParkingTariffNight
            '������� ������� ������� � ����������� -
            '  ������ (��� ���������� ��������)
    strTime = "Night"
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
            
            '������� ����� - "Not"
Private Sub optNot_Click()
            '������� ������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = True
    lblDate.Visible = True
            '������� ���������� ��������� �������� ���������� �����
    hsbLat.Enabled = True
    vsbDate.Enabled = True
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
            '������� ������� ������� � ����������� -
            '  �������� (��� ���������� ��������)
    strTime = "DayNight"
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
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

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
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
            
            '���������� ��������� ����� �������� ����������
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
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
            '������� ������� ������� � ����������� -
            '  �������� (��� ���������� ��������)
    strTime = "DayNight"
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
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

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
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
            
            '���������� ��������� ����� �������� ����������
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
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
            '������� ������� ������� � ����������� -
            '  �������� (��� ���������� ��������)
    strTime = "DayNight"
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
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

            '���������� ���� �� ����� ������
    intToMonthEnd = -1
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
            
            '���������� ��������� ����� �������� ����������
    txtMoneyDate.Text = Format(Now, "dd/mm/yyyy")
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
            
            '������� ����� - "MoneyFree"
Private Sub optMoneyFree_Click()
            '������� ����������� ������� ���������� "fraDayNight"
    fraDayNight.Enabled = False
            '������� ����� "optDayNight"
    optDayNight.Value = True
            '������� ��������� ����� ���������� ���� "txtMoneyDate"
    lblMoneyDate.Visible = False
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
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
    lblDate.Visible = False
            '������� ������������ ��������� �������� ���������� �����
    hsbLat.Enabled = False
    vsbDate.Enabled = False
    fraMonth.Enabled = False
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
            '������� ��������� ������� �� ������ "OK _ +"
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 Then
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        If cmdOK.Visible = True Then cmdOK.SetFocus
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
            '�������� �������o� ���e
    txtParole.Text = ""
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
            '���������� ����������� ����
        intParkingDay = Int(intParkingMoney / intParkingTariff)
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
                If intParkingDay < 1 Then GoTo EndCycle
            '���������� ����������� ����
                intParkingDay = intParkingDay - 1
            '������� �������� ����� � ������� ������ "������� ���������"
                intPosNum = InStr(1, Trim(frmTableCalendar.grdTableCalendar.Text), "/")
                If intPosNum <> 0 Then
            '���������  ����� � ������� ���� "����������"
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
    If Int(intParkingMoney / intParkingTariff) = 0 Then
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
    If txtPersonCode.Tag = 1 And txtInfo.Tag = 1 And txtMoneyDate.Tag = 1 Then
            '������� ��������� ������� �� ������ "OK _ +"
        cmdOK.MousePointer = 0
            '���������� ����� �� ������ "��_+"
        cmdOK.SetFocus
    End If
    
End Sub

