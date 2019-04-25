VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmTableCalendar 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_calendar"
   ClientHeight    =   4800
   ClientLeft      =   2400
   ClientTop       =   3015
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   186
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6615
   Begin MSACAL.Calendar comCalendar 
      Height          =   495
      Left            =   1320
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
      _Version        =   524288
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   1
      Day             =   1
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   0   'False
      ShowDays        =   0   'False
      ShowHorizontalGrid=   0   'False
      ShowTitle       =   0   'False
      ShowVerticalGrid=   0   'False
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.Timer tmrMinute 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Tag             =   "0"
      Top             =   1560
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2520
      Max             =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Frame fraDayType 
      Caption         =   "Day Correction"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   240
      Width           =   4212
      Begin VB.OptionButton optWorkDay 
         Caption         =   " - Workday"
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
         Height          =   192
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton optSpecDay 
         Caption         =   "/^ - Specday"
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
         Height          =   192
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1332
      End
      Begin VB.OptionButton optHoliday 
         Caption         =   "/* - Holyday"
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
         Height          =   192
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.ListBox lstWeekNum 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableCalendar.frx":0000
      Left            =   120
      List            =   "frmTableCalendar.frx":0002
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Frame fraDay 
      Caption         =   "Day"
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
      Height          =   492
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   4212
      Begin VB.OptionButton optSun 
         Enabled         =   0   'False
         Height          =   192
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optSat 
         Enabled         =   0   'False
         Height          =   192
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optFri 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optThu 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optWed 
         Enabled         =   0   'False
         Height          =   192
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optTue 
         Enabled         =   0   'False
         Height          =   192
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton optMon 
         Enabled         =   0   'False
         Height          =   192
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCorrection 
      Caption         =   "Correction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   1092
   End
   Begin VB.CommandButton cmdNewCalen 
      Caption         =   "NewCalen"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "SaveAs..."
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
      Left            =   3360
      TabIndex        =   3
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   1212
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableCalendar 
      Height          =   1815
      Left            =   2160
      TabIndex        =   0
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   54
      Cols            =   8
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Calendars variant"
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
      TabIndex        =   24
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<="
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
      TabIndex        =   23
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblVariant99 
      Alignment       =   2  'Center
      Caption         =   "V0"
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
      Left            =   3720
      TabIndex        =   22
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblVariant0 
      Alignment       =   2  'Center
      Caption         =   "V0"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   1800
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2520
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3840
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2040
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6480
      X2              =   2040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   4080
      Y2              =   1560
   End
   Begin VB.Label lblWeekNum 
      Caption         =   "Week Number "
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
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1332
   End
End
Attribute VB_Name = "frmTableCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '"������" ����� �������� "������� ���������"
Dim intVariantOld As Integer
            '"�����" ����� �������� "������� ���������"
Dim intVariantNew As Integer
            '������� ����� �����
Dim intFileNum As Integer
           '������� ����� �������������� ������ "������� ���������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "������� ���������"
Dim intColNumCorr As Integer
            '������ "��������� �������"
Dim gSystem As SystemInfo
            '������ "������� ������"
Dim gPerson As PersonInfo
            '������ "������� ���������"
Dim gCalendar As CalendarInfo

            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            '"������" ����� �������� "������� ���������" �� �������
    If hsbVariant.Value <> 0 Then
            
            '���� �� ����������� ��������� � "������� ���������"
        If gChangesTableCalendar = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ���������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� ������� ��������� � ����� �� ���������
                cmdSave_Click
            End If
        End If

    Else
            
            '���� �� ����������� ��������� � "������� ���������"
        If gChangesTableCalendar = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ���������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� ������� ��������� � ����� �� ���������
                cmdSave_Click
            End If
        End If

            '������� ������������ �������� ���������� ���������� "������� ���������"
        fraDayType.Enabled = False
        optHoliday.Enabled = False
        optSpecDay.Enabled = False
        optWorkDay.Enabled = False
        fraDay.Enabled = False
        optMon.Enabled = False
        optTue.Enabled = False
        optWed.Enabled = False
        optThu.Enabled = False
        optFri.Enabled = False
        optSat.Enabled = False
        optSun.Enabled = False
        lblWeekNum.Enabled = False
        lstWeekNum.Enabled = False
    End If
    
                '�������� ������ ������� ������
    lstWeekNum.Clear
        
            '������� ������� ������� ������� "������� ����������"
    intVariantOld = 0
    hsbVariant.Value = 0
            '�������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = False

            '������� ��������� ������� �����
    frmTableCalendar.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub

            '����� ��������� (�� ��������� ���)
Private Sub cmdNewCalen_Click()
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer

            '������� ������������ �������� ���������� ���������� "������� ���������"
    fraDayType.Enabled = False
    optHoliday.Enabled = False
    optSpecDay.Enabled = False
    optWorkDay.Enabled = False
    fraDay.Enabled = False
    optMon.Enabled = False
    optTue.Enabled = False
    optWed.Enabled = False
    optThu.Enabled = False
    optFri.Enabled = False
    optSat.Enabled = False
    optSun.Enabled = False
    lblWeekNum.Enabled = False
    lstWeekNum.Enabled = False
            
            '"�������" ������� "������� ���������"
    hsbVariant.Value = 0
            '��������� ������ �������� �������� "������� ���������"
    grdTableCalendar.Row = grdTableCalendar.Rows - 1
            '������ ������� ��������� ������
            ' �������� �������� "������� ���������"
    grdTableCalendar.Col = 1
            '���� ������������ ������ ���������
    comCalendar.Today
            '���� "������ ���� ������ ������ � ��������� ������ ����"
    If Left(Trim(grdTableCalendar.Text), 2) = "25" Then
        comCalendar.Value = Str(CInt(comCalendar.Year) + 1) + ".01.01."
    Else
        comCalendar.Value = Str(CInt(comCalendar.Year)) + ".12." + _
        Left(Trim(grdTableCalendar.Text), 2) + "."
    End If
            '"������" ������� "������� ���������"
    hsbVariant.Value = 1
    
            '���� �� ���� ��������������� ������� "������� ���������" - (����)
    For intRowNum = 1 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
       grdTableCalendar.Row = intRowNum
            '�� ���� �������� "������� ���������"
       For intColNum = 1 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '������ ���� � "������� ���������"
            grdTableCalendar.Text = comCalendar.Day
            '���� ���������� ���
            comCalendar.NextDay

            '���������� ������� ������������� ��������������� ������������
            '  ����� �������� ���� � "������� ���������" ��� ������ ����
            If gHolidays = 1 Then
            '�������� ���� - ������� ��� �����������
                If intColNum >= 6 Then _
                grdTableCalendar.Text = grdTableCalendar.Text + "/*"
            End If
                
        Next
    Next
    
              '���������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = True
            '���������� ����� �� ������ "Correction"
    If frmTableCalendar.Visible = True Then cmdCorrection.SetFocus

End Sub
            
            '���������
Private Sub cmdCorrection_Click()

            ' "������� ���������" �� �������� ��������������� �����
    If grdTableCalendar.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        MsgBox ("The table is empty")
    
    Else
            '������� ���������� �������� ���������� ���������� "������� ���������"
        fraDayType.Enabled = True
        optHoliday.Enabled = True
            '���� ������� ������� "T������ ���������"
        If hsbVariant.Value = 0 Then optSpecDay.Enabled = True
        optWorkDay.Enabled = True
        optWorkDay.Value = True
        fraDay.Enabled = True
        optMon.Enabled = True
        optMon.Value = True
        optTue.Enabled = True
        optWed.Enabled = True
        optThu.Enabled = True
        optFri.Enabled = True
        optSat.Enabled = True
        optSun.Enabled = True
        lblWeekNum.Enabled = True
        lstWeekNum.Enabled = True
            '�������� ������ ����
        lstWeekNum.Clear
    
            '������� "Week"
        grdTableCalendar.Col = 0
            '���� �� ���� ��������������� ������� "������� ���������"
        For intRowNumCorr = 1 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
            grdTableCalendar.Row = intRowNumCorr
            '���������� ������ "lstWeekNum" �������� �� "������� ���������"
            lstWeekNum.AddItem grdTableCalendar.Text
        Next
            '�������  ������� ������
        lstWeekNum.ListIndex = 0
            '����� �������������� ������ - (Week - 1)
        intRowNumCorr = 1
        grdTableCalendar.Row = intRowNumCorr
            '�������� �����
        optMon.Value = True
        optWorkDay.Value = True
    
    End If
    
End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "������� ���������"
Private Sub grdTableCalendar_Click()
            '��������� "��������"
    If lstWeekNum.Enabled = True Then
            '����� �������������� ������ "������� ���������"
        intRowNumCorr = grdTableCalendar.RowSel
        grdTableCalendar.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstWeekNum.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "������� ���������"
        intColNumCorr = grdTableCalendar.ColSel
        grdTableCalendar.Col = intColNumCorr
            '����� �������������� ������ "������� ���������"
        lstWeekNum_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstWeekNum.Left, _
        Y:=lstWeekNum.Top
            '����� ��������������� ������� "������� ���������"
        Select Case intColNumCorr
            Case 1
            optMon.Value = True
            Case 2
            optTue.Value = True
            Case 3
            optWed.Value = True
            Case 4
            optThu.Value = True
            Case 5
            optFri.Value = True
            Case 6
            optSat.Value = True
            Case 7
            optSun.Value = True
        End Select
    End If

End Sub

            '����� �������������� ������ "������� ���������"
Private Sub lstWeekNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� ��������������� ������� "������� ���������" �� ����� ���� (���������, �
            '  �� ���������� ������ ������� ������ - "lstWeekNum")
        If intColNumCorr <> 0 Then
            '����� �������������� ������ "������� ���������"
            intRowNumCorr = lstWeekNum.ListIndex + 1
        End If
    End If

End Sub

            '������� ����� ��� - "�������� ����"
Private Sub optHoliday_GotFocus()
            '����� ��������������� ������� "������� ���������"
    grdTableCalendar.Col = intColNumCorr
            '����� �������������� ������ "������� ���������"
    grdTableCalendar.Row = intRowNumCorr
            '��������� ����� ��� � "������� ���������"
    If InStr(1, Trim(grdTableCalendar.Text), "/") = 0 Then
        grdTableCalendar.Text = Trim(grdTableCalendar.Text) + "/*"
    Else
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1) + "/*"
    End If
    
            '���������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = True

End Sub

            '������� ����� ��� - "����������� ������� ����"
Private Sub optSpecDay_GotFocus()
            '����� ��������������� ������� "������� ���������"
    grdTableCalendar.Col = intColNumCorr
            '����� �������������� ������ "������� ���������"
    grdTableCalendar.Row = intRowNumCorr
            '��������� ����� ��� � "������� ���������"
    If InStr(1, Trim(grdTableCalendar.Text), "/") = 0 Then
        grdTableCalendar.Text = Trim(grdTableCalendar.Text) + "/^"
    Else
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1) + "/^"
    End If
    
            '���������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = True

End Sub

            '������� ����� ��� - "������� ����"
Private Sub optWorkDay_GotFocus()
            '����� ��������������� ������� "������� ���������"
    grdTableCalendar.Col = intColNumCorr
            '����� �������������� ������ "������� ���������"
    grdTableCalendar.Row = intRowNumCorr
            '��������� ����� ��� � "������� ���������"
    If InStr(1, Trim(grdTableCalendar.Text), "/") <> 0 Then
        grdTableCalendar.Text = _
        Left(Trim(grdTableCalendar.Text), InStr(1, Trim(grdTableCalendar.Text), "/") - 1)
    End If
    
              '���������� ������� ��������� ��������� � "������� ���������"
  gChangesTableCalendar = True

End Sub

            '������� ����� "�����������"
Private Sub optMon_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 1
            
End Sub

            '������� ����� "�������"
Private Sub optTue_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 2

End Sub

            '������� ����� "�����"
Private Sub optWed_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 3

End Sub

            '������� ����� "�������"
Private Sub optThu_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 4

End Sub

            '������� ����� "�������"
Private Sub optFri_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 5

End Sub

            '������� ����� "�������"
Private Sub optSat_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 6

End Sub

            '������� ����� "�����������"
Private Sub optSun_Click()
            '����� ��������������� ������� "������� ���������"
    intColNumCorr = 7

End Sub

            '��������� ������� "Change" - ��������� ��� �������� "Variant"
Private Sub hsbVariant_Change()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '���������� ������ ��������
    SetColWidth
            '������� ������������ �������� ���������� ���������� "������� ���������"
    fraDayType.Enabled = False
    optHoliday.Enabled = False
    optSpecDay.Enabled = False
    optWorkDay.Enabled = False
    fraDay.Enabled = False
    optMon.Enabled = False
    optTue.Enabled = False
    optWed.Enabled = False
    optThu.Enabled = False
    optFri.Enabled = False
    optSat.Enabled = False
    optSun.Enabled = False
    lblWeekNum.Enabled = False
    lstWeekNum.Enabled = False
            '�������� ������ ������� ������
    lstWeekNum.Clear
            
            '���� ������������� ��������� � "������� ���������"
    If gChangesTableCalendar = True Then
            '������ �������� ������
        frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ���������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '��������� "�����" ����� �������� "������� ���������"
            intVariantNew = hsbVariant.Value
            '"������" ����� �������� "������� ���������"
            hsbVariant.Value = intVariantOld
            '���������� "������� ���������" � ����� �� ���������
            cmdSave_Click
            '������������ "�����" ����� �������� "������� ���������"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"������" ����� �������� "������� ���������"
    intVariantOld = hsbVariant.Value
            '�������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = False

            '���������� �������� "������� ���������" �� �����
            
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gCalendar)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            '���� ����������� - ?
    On Error GoTo ErrorTableCalendar
            '���� (�������)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������� ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� �������� "������� ���������"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        grdTableCalendar.Row = intRowNum
            '������ ������ "������� ���������" �� ����� � �����
        Get intFileNum, intRowNum + 1, gCalendar
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '���������� ������� ������ "������� ���������" �� ������
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
        Next
    Next
            '������� ����
    Close intFileNum
    
            '���������� ����� �� ������ "Correction"
    If frmTableCalendar.Visible = True Then cmdCorrection.SetFocus
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = hsbVariant.Value
    
    Exit Sub
ErrorTableCalendar:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")

End Sub
            
            '���������� "������� ���������" � ����� �� ���������
Private Sub cmdSave_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gCalendar)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ���� ������� "������� ���������"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        grdTableCalendar.Row = intRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '���������� ������� ������ "������� ���������" �� ������
            Select Case intColNum
                Case 0
                gCalendar.strWeekNum = Trim(grdTableCalendar.Text)
                Case 1
                gCalendar.strMonday = Trim(grdTableCalendar.Text)
                Case 2
                gCalendar.strTuesday = Trim(grdTableCalendar.Text)
                Case 3
                gCalendar.strWednesday = Trim(grdTableCalendar.Text)
                Case 4
                gCalendar.strThursday = Trim(grdTableCalendar.Text)
                Case 5
                gCalendar.strFriday = Trim(grdTableCalendar.Text)
                Case 6
                gCalendar.strSaturday = Trim(grdTableCalendar.Text)
                Case 7
                gCalendar.strSunday = Trim(grdTableCalendar.Text)
           End Select
        Next
            '�������� ������ "������� ���������" �� ������ � ����
            Put intFileNum, intRowNum + 1, gCalendar
    Next
            '������� ������������ ����
    Close intFileNum
            '�������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = False
            '���������� ����� �� ������ "Cancel"
    If frmTableCalendar.Visible = True Then cmdCancel.SetFocus
            
End Sub
            '���������� ������� ��������� � ���������� �����
Private Sub cmdSaveAs_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer

            '��������� (�� ���������) ����� "frmGetFile"
    Load frmGetFile
            '��������� ������ ���������������� ���� "cboFileType
    frmGetFile.cboFileType.AddItem "All files (*.*)"
    frmGetFile.cboFileType.AddItem "Text files (*.TXT)"
    frmGetFile.cboFileType.AddItem "Word document(*.DOC)"
            '������� ������� ������ "��� �����"
    frmGetFile.cboFileType.ListIndex = 0
            '������� �� ����� ����� "frmGetFile" � ������� ����������� 1
    frmGetFile.Show 1
            '���� �� ������
    If frmGetFile.Tag = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The file isn't selected !"
            '������ "������� ������" � ��������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
    strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gCalendar)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '������� ���������� ���� ��� ������������� ������� ���
            '  ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ���� ������� "������� ���������"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        grdTableCalendar.Row = intRowNum
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '���������� ������� ������ "������� ���������" �� ������
            Select Case intColNum
                Case 0
                gCalendar.strWeekNum = Trim(grdTableCalendar.Text)
                Case 1
                gCalendar.strMonday = Trim(grdTableCalendar.Text)
                Case 2
                gCalendar.strTuesday = Trim(grdTableCalendar.Text)
                Case 3
                gCalendar.strWednesday = Trim(grdTableCalendar.Text)
                Case 4
                gCalendar.strThursday = Trim(grdTableCalendar.Text)
                Case 5
                gCalendar.strFriday = Trim(grdTableCalendar.Text)
                Case 6
                gCalendar.strSaturday = Trim(grdTableCalendar.Text)
                Case 7
                gCalendar.strSunday = Trim(grdTableCalendar.Text)
           End Select
        Next
            '�������� ������ "������� ���������" �� ������ � ����
            Put intFileNum, intRowNum + 1, gCalendar
    Next
            '�������  ��������� ����
    Close intFileNum
                '�������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = False

    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '���������� ����� �� ������ "Cancel"
    If frmTableCalendar.Visible = True Then cmdCancel.SetFocus
    
End Sub

            '�������� ����� "������� ���������"
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer
            '������� �����
Dim intMonth As Integer
            '������� ������ ������� ���� � ������ "������� ���������" - (����� ������� �����)
Dim intRightPositionDay

            '���������� ������ ��������
    SetColWidth
            '���������� ��������� "������� ���������"
    lblVariant99.Caption = "V" + Str(gVarNumCalendar)
    hsbVariant.Max = gVarNumCalendar
            '�������������� ����������� �������
ReDim gToday(gVarNumCalendar) As String * 4
            '��������� "������" ����� �������� "������� ���������"
    intVariantOld = hsbVariant.Value
            
            '���������� "������� ���������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gCalendar)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(hsbVariant.Value)) + ".dat"
    
    
            '���� ����������� - ?
    On Error GoTo ErrorTableCalendar
            '���� (�������)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������� �����
    intMonth = 0
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        grdTableCalendar.Row = intRowNum
            '������ ������ "������� ���������" �� ����� � �����
        Get intFileNum, intRowNum + 1, gCalendar
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '���������� ������� ������ "������� ���������" �� ������
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
            '���� ��������������� ������ � ������� "������� ���������"
            If grdTableCalendar.Row <> 0 And grdTableCalendar.Col <> 0 Then
            '���������� ������� ������ ������� ���� � ������ "������� ���������"
                intRightPositionDay = InStr(1, Trim(grdTableCalendar.Text), "/")
                If intRightPositionDay = 0 Then
            '����� ��� ����������� � ������ "������� ���������" - (���������� ������� ����)
                    intRightPositionDay = 2
            '����� ��� ������������ � ������ "������� ���������"
                Else
                    intRightPositionDay = intRightPositionDay - 1
                End If
            '���� 1-�� �����, �� ����� ������ +1
                If CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) = 1 _
                Then intMonth = intMonth + 1
            '���� "�����������" ������� �����
                If intMonth = Mid(gProtocol.strProtocDate, 4, 2) Then
            '���� ���� �������
                    If CInt(Left(gProtocol.strProtocDate, 2)) = _
                    CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) Then
            '���� (�������) � �������
                        gToday(0) = Trim(grdTableCalendar.Text)
            '����� ������ "������� ���������", ��� ����������� ������ �������� ���
                        gRowNum = intRowNum
            '����� ������� "������� ���������", ��� ����������� ������ �������� ���
                        gColNum = intColNum
                    End If
                End If
            End If
        Next
    Next
            '������� ������������ ����
    Close intFileNum
                
            '���� �� ���� ��������� ��������� "������� ���������"
    For intVariantNew = 1 To gVarNumCalendar Step 1
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableCalendar" + Trim(Str(intVariantNew)) + ".dat"
    
            '������� ������������ ���� ��� ������������� �������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '������ ������ ���������� �������� "������� ���������" �� ����� � �����
        Get intFileNum, gRowNum + 1, gCalendar
            '������� ���� (� �������) ��� ���������� ��������  "������� ���������"
        Select Case gColNum
            Case 1
            gToday(intVariantNew) = gCalendar.strMonday
            Case 2
            gToday(intVariantNew) = gCalendar.strTuesday
            Case 3
            gToday(intVariantNew) = gCalendar.strWednesday
            Case 4
            gToday(intVariantNew) = gCalendar.strThursday
            Case 5
            gToday(intVariantNew) = gCalendar.strFriday
            Case 6
            gToday(intVariantNew) = gCalendar.strSaturday
            Case 7
            gToday(intVariantNew) = gCalendar.strSunday
        End Select
            '������� ������������ ����
        Close intFileNum
    Next
                
                '�������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = False
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableCalendar:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")
    
End Sub

            '�������� ����� "������� ���������" � ����� ����
Private Sub NewYear()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ���������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ���������"
Dim intColNum As Integer
            '������� �����
Dim intMonth As Integer
            '������� ������ ������� ���� � ������ "������� ���������" - (����� ������� �����)
Dim intRightPositionDay

            '���������� ������ ��������
    SetColWidth
            '���������� ��������� "������� ���������"
    lblVariant99.Caption = "V" + Str(gVarNumCalendar)
    hsbVariant.Max = gVarNumCalendar
            '�������������� ����������� �������
ReDim gToday(gVarNumCalendar) As String * 4
            '��������� "������" ����� �������� "������� ���������"
    intVariantOld = hsbVariant.Value
            
            '���������� "������� ���������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� ���������"
    lngRecordLen = Len(gCalendar)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableCalendar" + "1" + ".dat"
    
    
            '���� ����������� - ?
    On Error GoTo ErrorTableCalendar
            '���� (�������)
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������� �����
    intMonth = 0
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ������� "������� ���������"
    For intRowNum = 0 To grdTableCalendar.Rows - 1 Step 1
            '������� ������ "������� ���������"
        grdTableCalendar.Row = intRowNum
            '������ ������ "������� ���������" �� ����� � �����
        Get intFileNum, intRowNum + 1, gCalendar
            '�� ���� �������� "������� ���������"
        For intColNum = 0 To grdTableCalendar.Cols - 1 Step 1
            '������� ������� "������� ���������"
            grdTableCalendar.Col = intColNum
            '���������� ������� ������ "������� ���������" �� ������
            Select Case intColNum
                Case 0
                If intRowNum = 0 Then grdTableCalendar.Text = "Week Numb."
                If intRowNum <> 0 Then grdTableCalendar.Text = gCalendar.strWeekNum
                Case 1
                grdTableCalendar.Text = gCalendar.strMonday
                Case 2
                grdTableCalendar.Text = gCalendar.strTuesday
                Case 3
                 grdTableCalendar.Text = gCalendar.strWednesday
                Case 4
                grdTableCalendar.Text = gCalendar.strThursday
                Case 5
                grdTableCalendar.Text = gCalendar.strFriday
                Case 6
                grdTableCalendar.Text = gCalendar.strSaturday
                Case 7
                grdTableCalendar.Text = gCalendar.strSunday
            End Select
            '���� ��������������� ������ � ������� "������� ���������"
            If grdTableCalendar.Row <> 0 And grdTableCalendar.Col <> 0 Then
            '���������� ������� ������ ������� ���� � ������ "������� ���������"
                intRightPositionDay = InStr(1, Trim(grdTableCalendar.Text), "/")
                If intRightPositionDay = 0 Then
            '����� ��� ����������� � ������ "������� ���������" - (���������� ������� ����)
                    intRightPositionDay = 2
            '����� ��� ������������ � ������ "������� ���������"
                Else
                    intRightPositionDay = intRightPositionDay - 1
                End If
            '���� 1-�� �����, �� ����� ������ +1
                If CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) = 1 _
                Then intMonth = intMonth + 1
            '���� "�����������" ������� �����
                If intMonth = Mid(gProtocol.strProtocDate, 4, 2) Then
            '���� ���� �������
                    If CInt(Left(gProtocol.strProtocDate, 2)) = _
                    CInt(Left(Trim(grdTableCalendar.Text), intRightPositionDay)) Then
            '���� (�������) � �������
                        gToday(0) = Trim(grdTableCalendar.Text)
            '����� ������ "������� ���������", ��� ����������� ������ �������� ���
                        gRowNum = intRowNum
            '����� ������� "������� ���������", ��� ����������� ������ �������� ���
                        gColNum = intColNum
                    End If
                End If
            End If
        Next
    Next
            '������� ������������ ����
    Close intFileNum
            '��������������� ������ ������ ��������� �� ������� "������� ���������"
    hsbVariant.Value = 0
                '���������� ������� ��������� ��������� � "������� ���������"
    gChangesTableCalendar = True
            '��������� "������� ���������" � ������������ ����� (��� ������ ����)
    Call cmdSave_Click
            
    
    Exit Sub
ErrorTableCalendar:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("New Year TableCalendar Error !")
    
End Sub

            '��������� �������� ������� (����� �����, ����, ��� � ����)
Private Sub tmrMinute_Timer()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '��� �������� ��� ������ ������� "Shell"
Dim vntShell As Variant
            '����� ������ "��������� �������", "������� ���������" ��� "������� ���������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "��������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "��������� �������"
Dim intColNum As Integer
            '����
Dim strDate As String
            '�����
Dim strTime As String
Dim intHour As Integer
Dim intMinute As Integer
Dim strHour As String
Dim strMinute As String
            '������� �������
Dim intCount As Integer
            '������ ����������� ���������
Dim strMessage As String
            
            '������� ����
    strDate = Trim(Format(Now, "dd/mm/yyyy"))
            
            '������� �����
    strTime = Format(Now, "h:mm:ss")
            '����
    intHour = Hour(strTime)
    If intHour < 10 Then
        strHour = "0" + Trim(Str(intHour))
    Else
        strHour = Trim(Str(intHour))
    End If
            '������
    intMinute = Minute(strTime)
    If intMinute < 10 Then
        strMinute = "0" + Trim(Str(intMinute))
    Else
        strMinute = Trim(Str(intMinute))
    End If
    
            '����� ���� � �������
    frmDemo.lblTime.Caption = "   " + strDate + "   " + strHour + ":" + strMinute

        
            '���� ���������� ������� ������������� ������ "������� ������":
            '   ��������������� ������ � "Host Computer'e" � � ��� �������
            '   � "�������������", ����� ��������� ���������� ����
            '   ����������� "������� ������" - "���������� ������a ������"
    If gCompresTablPers = 1 Then
            '���� ����� "����������"
        If frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
            '���� ���������� ������ �� �������� �������� ����� ��
            '   "������� ������", "������a ������" �������� ���
            '   ������������ ����� � ���������� ��������� ��������
            If gRealDelPerson = True And gTablePerson.Access < 1 And _
            frmDemo.lblInform(frmDemo.tmrTermContr.Tag).Tag = 0 Then
            '��������� ������� ��������� ��������� ������ '������� ������"
                Call frmTablePerson.RealDelPerson
            End If
            '���� ���������� ������� ��������� ��������� �
            '   "������� ������" - ��������� ������� � ������������ �����
            If gChangesTablePerson = True Then
                Call frmTablePerson.SaveTablePerson
            End If
        End If
    End If
            
            '���� ��� �� "Host Computer" � ����� "����������"
    If gPreprocName <> "" And frmDemo.Enabled = True And _
    frmDemo.chkSetup.Value = 1 Then
            
            '���� ���� ������� ������� ������������ ��� ��������� �
            '   ����������� - �������/������� "������� ���������"
        If gMSBase = 0 Then
            '������� ���� "������� ���������"
            Close gProtocFileNum
            '��������� ����� ������ (������) "������� ���������"
            lngRecordLen = Len(gProtocol)
            '�������� ��������� ����� ����� "������� ���������"
            gProtocFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            strPathFileName = strPathFileName + "TableProtocol.dat"
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
            Open strPathFileName For Random As gProtocFileNum Len = lngRecordLen
            '����� ������ ��������� ������ "������� ���������"
            gProtocRowNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    
    End If
    
    
            '��� ����� � ����� "����������"
    If strHour <> frmDemo.lblTime.Tag And _
    frmDemo.Enabled = True And frmDemo.chkSetup = 1 Then
            '��������� ����� ����� (���) ������������ ���
        frmDemo.lblTime.Tag = strHour
            
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� ����������� ��� �� �����������
        If gParkingPlaceNum <> 0 Or gAccessPlaceNum <> 0 Then
            
            If gParkingPlaceNum <> 0 Then
            '������� ��������������� ������� � ������
            '  �����������
                If Not (Left(gDefaultParkTime, 2) = "00" And _
                Mid(gDefaultParkTime, 4, 2) = "00" And _
                Mid(gDefaultParkTime, 7, 2) = "24" And _
                Mid(gDefaultParkTime, 10, 2) = "00") And _
                Left(gDefaultParkTime, 2) = intHour And _
                Mid(gDefaultParkTime, 4, 2) >= intMinute Then
            '�������� ���������� ��������� ���� �� �����������
                    gParkFreePlaces = gParkingPlaceNum
                ElseIf Not (Left(gDefaultParkTime, 2) = "00" And _
                Mid(gDefaultParkTime, 4, 2) = "00" And _
                Mid(gDefaultParkTime, 7, 2) = "24" And _
                Mid(gDefaultParkTime, 10, 2) = "00") And _
                Mid(gDefaultParkTime, 7, 2) = intHour Then
            '������� ���������� ��������� ���� �� �����������
                    gParkFreePlaces = 0
                End If
            '���������� ������������� �������-���������
                strMessage = "ParkFreePlaces=" + CStr(gParkFreePlaces)
            End If
            
            If gAccessPlaceNum <> 0 Then
            '������� ��������������� ������� � ������
            '  �����������
                If Not (Left(gDefaultAcceTime, 2) = "00" And _
                Mid(gDefaultAcceTime, 4, 2) = "00" And _
                Mid(gDefaultAcceTime, 7, 2) = "24" And _
                Mid(gDefaultAcceTime, 10, 2) = "00") And _
                Left(gDefaultAcceTime, 2) = intHour And _
                Mid(gDefaultAcceTime, 4, 2) >= intMinute Then
            '�������� ���������� ��������� ���� �� �����������
                    gAcceFreePlaces = gAccessPlaceNum
                ElseIf Not (Left(gDefaultAcceTime, 2) = "00" And _
                Mid(gDefaultAcceTime, 4, 2) = "00" And _
                Mid(gDefaultAcceTime, 7, 2) = "24" And _
                Mid(gDefaultAcceTime, 10, 2) = "00") And _
                Mid(gDefaultAcceTime, 7, 2) = intHour Then
            '������� ���������� ��������� ���� �� �����������
                    gAcceFreePlaces = 0
                End If
            '���������� ������������� �������-���������
                strMessage = "AcceFreePlaces=" + CStr(gAcceFreePlaces)
            End If
            
            '������� ���������� �� �������
            Call frmDemo.Display(strMessage)
            '��������� ����� '������� ������"
            Call frmTablePerson.SaveTablePerson
        End If
            
            '���������� ������� ������������ ��� ��������� � �����������
        If gMSBase = 1 Then
            '������������ ��� ��������� � ����������� � ������� ACCESS"
            Call frmDemo.BasesConvert
        End If
    
    End If
            
            '����� ������� ��� ���������� ��������� �������� ����
            '  � "��������� �������"
    If strDate <> frmTableCalendar.Tag Or _
    Right(Trim(frmTableCalendar.Tag), 4) <> Trim(Str(gYear)) Then
    
            '���� ��� �� "Host Computer" � ����� "����������"
        If gPreprocName <> "" And frmDemo.Enabled = True And _
        frmDemo.chkSetup.Value = 1 Then
            '������������ ��������� "Host Computer'�" � �������������
            '  ������������� ������� ��� ������� �������������
            qMsgOutput.Body = "Time"
            ' ���������� ���� � ������� ������������ ���������
            qInfoOutput.FormatName = "DIRECT=OS:" + gHost + "\Private$\GeneralQueue"
            ' ������� ������� ��������� � ����������� (��� ��������
            '   ���������, ������ � ������� �������� ����)
            Set qQueueOutput = qInfoOutput.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            ' �������� ���������
            qMsgOutput.Send qQueueOutput
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
            If gParkingPlaceNum <> 0 Then
            ' ������������ ����������� ���������
                qMsgOutput.Body = "ParkFreePlaces "
            ' �������� ���������
                qMsgOutput.Send qQueueOutput
            End If
            ' ���� ������� �������-��������� ���������� ��������� ����
            '   �� �����������
            If gAccessPlaceNum <> 0 Then
            ' ������������ ����������� ���������
                qMsgOutput.Body = "AcceFreePlaces "
            ' �������� ���������
                qMsgOutput.Send qQueueOutput
            End If
            ' ������� ������� ���������
            qQueueOutput.Close
        End If
    
            '����� ��� �������� ��� ���������� ��������� �������� ����
        If Right(Trim(strDate), 4) <> Right(Trim(frmTableCalendar.Tag), 4) Or _
        Right(Trim(strDate), 4) <> Trim(Str(gYear)) Then
            '��������� "������� ���������" ������ ����
            Call NewYear
            '��������� ����� ����
            frmTableCalendar.Tag = strDate
            
            '���� �� ���� ��������������� ������� "��������� �������"
            For intRowNum = 1 To frmTableSystem.grdTableSystem.Rows - 1 Step 1
            '������� ������  "��������� �������"
                frmTableSystem.grdTableSystem.Row = intRowNum
            '������������� ������� "��������� �������" (������)
                frmTableSystem.grdTableSystem.Col = 0
                If Trim(frmTableSystem.grdTableSystem.Text) = "gYear" Then
            '������� ������� "��������� �������"=1(���������)
                    frmTableSystem.grdTableSystem.Col = 1
            '���������� ������� ���
                    gYear = Right(Trim(strDate), 4)
                    frmTableSystem.grdTableSystem.Text = gYear
                End If
            Next
            '��������� ����� '��������� �������"
            Call frmTableSystem.SaveTableSystem
            
            '������������ ����� "������� ���������" ��� ���������� ������ ����
            Call cmdNewCalen_Click
            '��������������� ������ ������ ��������� �� ������
            '  "������� ���������" (��� ������ ����)
            hsbVariant.Value = 1
            '��������� "������� ���������" � ������������ ����� (��� ������ ����)
            Call cmdSave_Click
            '��������������� ������ ������ ��������� �� ������� "������� ���������"
            hsbVariant.Value = 0
        
        Else
            '��������� ����� ����
            frmTableCalendar.Tag = strDate
            '����� ������� "������� ���������", ��� ����������� ������ �������� ��� = 7
            If gColNum = 7 Then
            '����� ������� "������� ���������", ��� ����������� ������ �������� ���
                gColNum = 1
            '����� ������ "������� ���������", ��� ����������� ������ �������� ���
            gRowNum = gRowNum + 1
            Else
            '����� ������� "������� ���������", ��� ����������� ������ �������� ���
                gColNum = gColNum + 1
            End If
        End If
            
            '��������� ����� ������ (������) "������� ���������"
        lngRecordLen = Len(gCalendar)
            
            '���� �� ���� ��������� "������� ���������"
        For intVariantNew = 0 To gVarNumCalendar Step 1
            '�������� ��������� ����� �����
            intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
            strPathFileName = App.Path
            If Right(strPathFileName, 1) <> "\" Then
                strPathFileName = strPathFileName + "\"
            End If
            strPathFileName = strPathFileName + "TableCalendar" + _
            Trim(Str(intVariantNew)) + ".dat"
    
            '���� ����������� - ?
            On Error GoTo ErrorTableCalendar
            '������� ������������ ���� ��� ������������� �������
            Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '������ ������ �������� "������� ���������" �� ����� � �����
            Get intFileNum, gRowNum + 1, gCalendar
            '������� ���� (� �������) ��� ��������  "������� ���������"
            Select Case gColNum
                Case 1
                gToday(intVariantNew) = gCalendar.strMonday
                Case 2
                gToday(intVariantNew) = gCalendar.strTuesday
                Case 3
                gToday(intVariantNew) = gCalendar.strWednesday
                Case 4
                gToday(intVariantNew) = gCalendar.strThursday
                Case 5
                gToday(intVariantNew) = gCalendar.strFriday
                Case 6
                gToday(intVariantNew) = gCalendar.strSaturday
                Case 7
                gToday(intVariantNew) = gCalendar.strSunday
            End Select
            '������� ������������ ����
            Close intFileNum
        Next
    End If
    
    Exit Sub
ErrorTableCalendar:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableCalendar Error !")

End Sub
            
            '��������� ��������� ������ � ������������ �������� "������� ���������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '������������� �������
    grdTableCalendar.ColWidth(intColNumber) = 1070
    grdTableCalendar.ColAlignment(intColNumber) = 0
            '���� �� ���� ��������������� ��������
    For intColNumber = 1 To grdTableCalendar.Cols - 1 Step 1
        grdTableCalendar.ColWidth(intColNumber) = 415
        grdTableCalendar.ColAlignment(intColNumber) = 0
    Next
    
End Sub


