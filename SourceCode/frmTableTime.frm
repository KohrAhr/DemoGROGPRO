VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableTime 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_time"
   ClientHeight    =   4935
   ClientLeft      =   2640
   ClientTop       =   2745
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
   ScaleHeight     =   4935
   ScaleWidth      =   6615
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4080
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2400
      Max             =   0
      TabIndex        =   20
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox txtExpander 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4320
      TabIndex        =   19
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Frame fraColName 
      Caption         =   "Options"
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
      Height          =   1452
      Left            =   1920
      TabIndex        =   15
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optTime 
         Caption         =   "Time"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   -1  'True
         Width           =   732
      End
      Begin VB.OptionButton optExpander 
         Caption         =   "Expander"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1452
      End
      Begin VB.CheckBox chkFromToTime 
         Caption         =   "v From     To"
         Enabled         =   0   'False
         Height          =   492
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.HScrollBar hsbMinute 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4320
      Max             =   59
      TabIndex        =   12
      Top             =   840
      Width           =   1452
   End
   Begin VB.HScrollBar hsbHour 
      Enabled         =   0   'False
      Height          =   252
      Left            =   4320
      Max             =   23
      TabIndex        =   9
      Top             =   480
      Width           =   1452
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete..."
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
      TabIndex        =   8
      Top             =   3480
      Width           =   1092
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
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
      TabIndex        =   7
      Top             =   2880
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
      Left            =   3720
      TabIndex        =   6
      Top             =   4320
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4320
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
      TabIndex        =   4
      Top             =   4320
      Width           =   1212
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
      TabIndex        =   2
      Top             =   360
      Width           =   1092
   End
   Begin VB.ListBox lstInterval 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableTime.frx":0000
      Left            =   120
      List            =   "frmTableTime.frx":0002
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableTime 
      Height          =   1695
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   3
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
      Left            =   3600
      TabIndex        =   24
      Top             =   2040
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
      Left            =   2040
      TabIndex        =   23
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<==="
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
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Intervals variant"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6480
      X2              =   1920
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1920
      X2              =   1920
      Y1              =   1800
      Y2              =   4200
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1920
      X2              =   120
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   4200
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   6480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   1800
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2640
      Y2              =   120
   End
   Begin VB.Label lblMinute59 
      Alignment       =   2  'Center
      Caption         =   "59min"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblMinute0 
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
      Left            =   4080
      TabIndex        =   13
      Top             =   840
      Width           =   135
   End
   Begin VB.Label lblHour23 
      Alignment       =   2  'Center
      Caption         =   "23h"
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
      Left            =   5880
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblHour0 
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
      Left            =   4080
      TabIndex        =   10
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lblIntervals 
      Alignment       =   2  'Center
      Caption         =   "Intervals "
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmTableTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������� ����� �������������� ������ "������� �������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "������� �������"
Dim intColNumCorr As Integer
            '"������" ����� �������� "������� �������"
Dim intVariantOld As Integer
            '"�����" ����� �������� "������� �������"
Dim intVariantNew As Integer
            '������� ����� �����
Dim intFileNum As Integer
            '������ "������� ������� - ����������"
Dim gTime As TimeInfo

            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            '"������" ����� �������� "������� �������" �� �������
    If hsbVariant.Value <> 0 Then

            '���� ������������� ��������� � "������� �������"
        If gChangesTableTime = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� �������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� "������� �������" � ����� �� ���������
                cmdSave_Click
            End If
        End If
            
    Else

            '���� ������������� ��������� � "������� �������"
        If gChangesTableTime = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� �������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� "������� �������" � ����� �� ���������
                cmdSave_Click
            End If
        End If
            
            '������� ������������ �������� ���������� ���������� "������� �������"
        lblIntervals.Enabled = False
        lstInterval.Enabled = False
        fraColName.Enabled = False
        optTime.Enabled = False
        chkFromToTime.Enabled = False
        optExpander.Enabled = False
        lblHour0.Enabled = False
        lblHour23.Enabled = False
        lblMinute0.Enabled = False
        lblMinute59.Enabled = False
        hsbHour.Enabled = False
        hsbMinute.Enabled = False
        txtExpander.Enabled = False
    End If
    
            '�������� ��������� ����
    txtExpander.Text = ""
            '�������� ������ ����������
    lstInterval.Clear
            
            '������� ������� ������� ������� "������� ����������"
    intVariantOld = 0
    hsbVariant.Value = 0
            '�������� ������� ��������� ��������� � "������� �������"
    gChangesTableTime = False

            '������� ��������� ������� �����
    frmTableTime.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '���������
Private Sub cmdCorrection_Click()

            ' "������� �������" �� �������� ��������������� �����
    If grdTableTime.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        MsgBox ("The table is empty")
    
    Else
            '������� ���������� ��������� �������� ���������� ���������� "������� �������"
        lblIntervals.Enabled = True
        lstInterval.Enabled = True
        fraColName.Enabled = True
        optTime.Enabled = True
        optTime.Value = True
        chkFromToTime.Enabled = True
        optExpander.Enabled = True
        lblHour0.Enabled = True
        lblHour23.Enabled = True
        lblMinute0.Enabled = True
        lblMinute59.Enabled = True
        hsbHour.Enabled = True
        hsbMinute.Enabled = True
            '�������� ��������� ����
        txtExpander.Text = ""
            '�������� ������ ����
        lstInterval.Clear
    
            '������� "Intervals"
        grdTableTime.Col = 0
            '���� �� ���� ��������������� ������� "������� �������"
        For intRowNumCorr = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
            grdTableTime.Row = intRowNumCorr
            '���������� ������ "lstInterval" �������� �� "������� �������"
            lstInterval.AddItem grdTableTime.Text
        Next
            '�������  ������� ������
        lstInterval.ListIndex = 0
            '����� �������������� ������ - (1)
        intRowNumCorr = 1
        grdTableTime.Row = intRowNumCorr
            '�������� �����
        optTime_Click
    End If
    
End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "������� �������"
Private Sub grdTableTime_Click()
            '��������� "��������"
    If lstInterval.Enabled = True Then
            '����� �������������� ������ "������� �������"
        intRowNumCorr = grdTableTime.RowSel
        grdTableTime.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstInterval.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "������� �������"
        intColNumCorr = grdTableTime.ColSel
        grdTableTime.Col = intColNumCorr
            '����� �������������� ������ "������� �������"
        lstInterval_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstInterval.Left, Y:=lstInterval.Top
            '����� ��������������� ������� "������� �������"
        Select Case intColNumCorr
            Case 1
            optTime.Value = True
            Case 2
            optExpander.Value = True
            '���������� ����� �� ��������� ���� ��� ���������
            txtExpander.SetFocus
        End Select
    End If
        
End Sub

            '����� �������������� ������ "������� �������"
Private Sub lstInterval_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� �������������� ������ "������� �������"
        intRowNumCorr = lstInterval.ListIndex + 1
        grdTableTime.Row = intRowNumCorr
        grdTableTime.Col = 2
            '����������� ������ "������� �������" � ��������� ���� ��� ���������
        txtExpander.Text = grdTableTime.Text
            '������������ ����� ��������������� ������� "������� �������"
        grdTableTime.Col = intColNumCorr
    End If

End Sub

            '������� ����� - "Time"
Private Sub optTime_Click()
            '����� ��������������� ������� "������� �������"
    intColNumCorr = 1
    grdTableTime.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� �������"
    lblHour0.Enabled = True
    lblHour23.Enabled = True
    lblMinute0.Enabled = True
    lblMinute59.Enabled = True
    hsbHour.Enabled = True
    hsbMinute.Enabled = True
    txtExpander.Enabled = False

End Sub

            '������� ����� "Expander"
Private Sub optExpander_Click()
            '����� ��������������� ������� "������� �������"
    intColNumCorr = 2
    grdTableTime.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� �������"
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
            '����������� ������ "������� �������" � ��������� ���� ��� ���������
    txtExpander.Text = grdTableTime.Text
    txtExpander.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtExpander.SetFocus

End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Hour"
Private Sub hsbHour_Scroll()
    hsbHour_Change
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Hour"
Private Sub hsbHour_Change()
            '������ ���������� ���������
    If chkFromToTime.Value = 1 Then
            '��������� ������ "Time" � "������� �������"
        If hsbHour.Value < 10 Then
            grdTableTime.Text = "0" + Trim(Str(hsbHour.Value)) + Mid(grdTableTime.Text, 3)
        Else
            grdTableTime.Text = Trim(Str(hsbHour.Value)) + Mid(grdTableTime.Text, 3)
        End If
            '����� ���������� ���������
    Else
            '��������� ������ "Time" � "������� �������"
        If hsbHour.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 6) + "0" + Trim(Str(hsbHour.Value)) _
            + Mid(grdTableTime.Text, 9)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 6) + Trim(Str(hsbHour.Value)) _
            + Mid(grdTableTime.Text, 9)
        End If
    End If
            '���������� �������  ��������� ��������� � "������� �������"
    gChangesTableTime = True
    
End Sub
            
            '��������� ������� "Scroll" - ��������� ��� �������� "Minute"
Private Sub hsbMinute_Scroll()
    hsbMinute_Change
End Sub
            
            '��������� ������� "Change" - ��������� ��� �������� "Minute"
Private Sub hsbMinute_Change()
            '������ ���������� ���������
    If chkFromToTime.Value = 1 Then
            '��������� ������ "Time" � "������� �������"
        If hsbMinute.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 3) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 6)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 3) + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 6)
        End If
            '����� ���������� ���������
    Else
            '��������� ������ "Time" � "������� �������"
        If hsbMinute.Value < 10 Then
            grdTableTime.Text = Left(grdTableTime.Text, 9) + "0" + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 12)
        Else
            grdTableTime.Text = Left(grdTableTime.Text, 9) + Trim(Str(hsbMinute.Value)) _
            + Mid(grdTableTime.Text, 12)
        End If
    End If
            '���������� �������  ��������� ��������� � "������� �������"
    gChangesTableTime = True

End Sub
            
            '��������� ����� � ������� ��������������� ���� "Expander"
Private Sub txtExpander_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtExpander.Text)) < 9 Then
            '��������� ���� "Expander" � "������� �������"
            grdTableTime.Text = Trim(txtExpander.Text)
            '���������� �������  ��������� ��������� � "������� �������"
            gChangesTableTime = True
            '�������� ������ ������
            '���������� ����� �� ������ "Save"
            cmdSave.SetFocus
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            '��������� ������� "Change" - ��������� ��� �������� "Variant"
Private Sub hsbVariant_Change()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "������� �������"
Dim intColNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '���������� ������ ��������
    SetColWidth
            '������� ������������ �������� ���������� ���������� "������� �������"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            '�������� ��������� ����
    txtExpander.Text = ""
            '�������� ������ ����������
    lstInterval.Clear
            
            '���� ������������� ��������� � "������� �������"
    If gChangesTableTime = True Then
            '������ �������� ������
        frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� �������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '��������� "�����" ����� �������� "������� �������"
            intVariantNew = hsbVariant.Value
            '"������" ����� �������� "������� �������"
            hsbVariant.Value = intVariantOld
            '���������� "������� �������" � ����� �� ���������
            cmdSave_Click
            '������������ "�����" ����� �������� "������� �������"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"������" ����� �������� "������� �������"
    intVariantOld = hsbVariant.Value
            '���������� ��������/���������� ����� � "������� �������"
    gAddDelRowTableTime = 0
            '�������� ������� ��������� ��������� � "������� �������"
    gChangesTableTime = False

            '���������� �������� "������� �������" �� �����
            
            '��������� ����� ������ (������) "������� �������"
    lngRecordLen = Len(gTime)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
                
                                
            '���� ����������� - ?
    On Error GoTo ErrorTableTime
            '���������� ����� "������� �������" ����� ������� ����� �� ��������� +1
    grdTableTime.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� �������  �������� "������� �������"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        grdTableTime.Row = intRowNum
            '������ ������ "������� �������" �� ����� � �����
        Get intFileNum, intRowNum, gTime
            '�� ���� �������� "������� �������"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
            grdTableTime.Col = intColNum
            '���������� ������� ������ "������� �������" �� ������
            Select Case intColNum
                Case 0
                grdTableTime.Text = gTime.strIntervalNum
                Case 1
                 grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                Case 2
                grdTableTime.Text = gTime.strExpander
            End Select
        Next
    Next
            '������� ����
    Close intFileNum
    
            '���������� ����� �� ������ "Correction"
    cmdCorrection.SetFocus
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = hsbVariant.Value
    
    Exit Sub
ErrorTableTime:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableTime Error !")
    
End Sub
            
            '���������� ������ � "������� �������"
Private Sub cmdAdd_Click()
    
            '������� ������������ �������� ���������� ���������� "������� �������"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            '�������� ������ ����������
    lstInterval.Clear
    
            '������������ ������ ��������� �������
    gTime.strIntervalNum = "Interval-" + Str(grdTableTime.Rows)
            '���������� ������ � ����� "������� �������"
    grdTableTime.AddItem gTime.strIntervalNum
            '������������ ������� ��� ��������� �������
    grdTableTime.Row = grdTableTime.Rows - 1
    grdTableTime.Col = 1
    grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
    "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
            '���������� ��������/���������� ����� � "������� �������"
    gAddDelRowTableTime = gAddDelRowTableTime + 1
            '���������� ������� ��������� ��������� � "������� �������"
    gChangesTableTime = True
            '���������� ����� �� ������ "Add"
    cmdAdd.SetFocus
    
End Sub
            
            '�������� ������ �� "������� �������"
Private Sub cmdDelete_Click()
            '������� ����� ��������������� ������ "������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "������� �������"
Dim intColNum As Integer
    
            '������� ������������ �������� ���������� ���������� "������� �������"
    lblIntervals.Enabled = False
    lstInterval.Enabled = False
    fraColName.Enabled = False
    optTime.Enabled = False
    chkFromToTime.Enabled = False
    optExpander.Enabled = False
    lblHour0.Enabled = False
    lblHour23.Enabled = False
    lblMinute0.Enabled = False
    lblMinute59.Enabled = False
    hsbHour.Enabled = False
    hsbMinute.Enabled = False
    txtExpander.Enabled = False
            '�������� ������ ����������
    lstInterval.Clear
    
            '��������� (�� ���������) ����� "frmSelectRow"
    Load frmSelectRow
            '���������������� �������� "lblColName" ����� "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Interval"
    
            '������� "Intervals"
    grdTableTime.Col = 0
             '�������� ������ ��������
    frmSelectRow.lstSelectRow.Clear
            '���� �� ���� ��������������� ������� "������� �������"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        grdTableTime.Row = intRowNum
            '���������� ������ "lstSelectRow" �������� �� "������� �������"
        frmSelectRow.lstSelectRow.AddItem grdTableTime.Text
    Next
            '������� ������� ������
    frmSelectRow.lstSelectRow.ListIndex = 0
            '������� �� ����� ����� "frmSelectRow" � ������� ����������� 1
    frmSelectRow.Show 1
            '������ �� �������
    If frmSelectRow.Tag = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The row isn't selected !"
            '��������� ������ �� ����� ���� �������
    ElseIf frmSelectRow.lstSelectRow.ListCount = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The last row isn't selected !"
            '�������� ������ �� "������� �������"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            '����� ��������� ������
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            '�������� ������
        grdTableTime.RemoveItem intRowNum
            '��������� ������ �� ��������� � "������� �������"
        If intRowNum < grdTableTime.Rows Then
             '���� �� ���� ������� "������� �������", ������� �� ��������� �������
            For intRowNum = intRowNum To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
                grdTableTime.Row = intRowNum
            '���������� ������ "lstSelectRow" �������� �� "������� �������"
                grdTableTime.Text = "interval-" + Str(intRowNum)
            Next
        End If
           '���������� ��������/���������� ����� � "������� �������"
        gAddDelRowTableTime = gAddDelRowTableTime - 1
            '���������� ������� ��������� ��������� � "������� �������"
        gChangesTableTime = True
    End If
            '��������� ����� "frmSelectRow"
    UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
    Set frmSelectRow = Nothing
            '���������� ����� �� ������ "Delete"
    cmdDelete.SetFocus
    
End Sub
            
            '���������� "������� �������" � ����� �� ���������
Private Sub cmdSave_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "������� �������"
Dim intColNum As Integer
            '��������� ����� ������ (������) "������� �������"
    lngRecordLen = Len(gTime)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
   
            '�����, ��������� �� "������� �������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
    If gAddDelRowTableTime < 0 Then
            '������� "������" ������������ ����
        Kill strPathFileName
    End If
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� �������"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        grdTableTime.Row = intRowNum
            '�� ���� �������� "������� �������"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
            grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� �������" � ����
            Select Case intColNum
                Case 0
                gTime.strIntervalNum = grdTableTime.Text
                Case 1
                gTime.strTime = Left(grdTableTime.Text, 2) + Mid(grdTableTime.Text, 4, 2) + _
                Mid(grdTableTime.Text, 7, 2) + Mid(grdTableTime.Text, 10, 2)
                Case 2
                gTime.strExpander = grdTableTime.Text
            End Select
        Next
            '�������� ������ "������� �������" � ����
        Put intFileNum, intRowNum, gTime
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� �������"
    gAddDelRowTableTime = 0
            '�������� ������� ��������� ��������� � "������� �������"
    gChangesTableTime = False
            '���������� ����� �� ������ "Cancel"
    cmdCancel.SetFocus
            
End Sub
            
            '���������� "������� �������" � ���������� �����
Private Sub cmdSaveAs_Click()
            '������ ��� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "������� �������"
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
            '������ "������� �������" � ��������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� �������"
        lngRecordLen = Len(gTime)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '�����, ��������� �� "������� �������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
        If gAddDelRowTableTime < 0 Then
            '������� "������" ����, ���� �� ����������
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� �������"
        For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
            grdTableTime.Row = intRowNum
            '�� ���� �������� "������� �������"
            For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
                grdTableTime.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� �������" � ����
                Select Case intColNum
                    Case 0
                    gTime.strIntervalNum = grdTableTime.Text
                    Case 1
                    gTime.strTime = Left(grdTableTime.Text, 2) + Mid(grdTableTime.Text, 4, 2) + _
                    Mid(grdTableTime.Text, 7, 2) + Mid(grdTableTime.Text, 10, 2)
                    Case 2
                    gTime.strExpander = grdTableTime.Text
                End Select
            Next
            '�������� ������ "������� �������" � ����
        Put intFileNum, intRowNum, gTime
        Next
            '������� ������������ ����
        Close intFileNum
            '���������� ��������/���������� ����� � "������� �������"
        gAddDelRowTableTime = 0
            '�������� ������� ��������� ��������� � "������� �������"
        gChangesTableTime = False
    End If
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '���������� ����� �� ������ "Cancel"
    cmdCancel.SetFocus
    
End Sub

            '�������� ����� "������� �������"
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "������� �������"
Dim intColNum As Integer
            '���������� �������� � ������� ���������� �������
Dim intIntervalNum As Integer

            '���������� ������ ��������
    SetColWidth
            '���������� ��������� "������� �������"
    lblVariant99.Caption = "V" + Str(gVarNumTime)
    hsbVariant.Max = gVarNumTime
            '��������� "������" ����� �������� "������� �������"
    intVariantOld = hsbVariant.Value
    
            '������� ������ = 0 (��������� ��������)
    grdTableTime.Row = 0
    grdTableTime.Col = 0
    grdTableTime.Text = "Intervals"
            '�������� � ������ (������ 0, ������� 1)
    grdTableTime.Col = 1
    grdTableTime.Text = "Time"
            '�������� � ������ (������ 0, ������� 2)
    grdTableTime.Col = 2
    grdTableTime.Text = "Expander"
            
            '���������� "������� �������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� �������"
    lngRecordLen = Len(gTime)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTime" + Trim(Str(hsbVariant.Value)) + ".dat"
                
            '���� ����������� - ?
    On Error GoTo ErrorTableTime
                '���������� ����� "������� �������" ����� ������� ����� �� ��������� +1
    grdTableTime.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� �������"
    For intRowNum = 1 To grdTableTime.Rows - 1 Step 1
            '������� ������ "������� �������"
        grdTableTime.Row = intRowNum
            '������ ������ "������� �������" �� ����� � �����
        Get intFileNum, intRowNum, gTime
            '�� ���� �������� "������� �������"
        For intColNum = 0 To grdTableTime.Cols - 1 Step 1
            '������� ������� "������� �������"
            grdTableTime.Col = intColNum
            '���������� ������� ������ "������� �������" �� ������
            Select Case intColNum
                Case 0
                grdTableTime.Text = gTime.strIntervalNum
                Case 1
                 grdTableTime.Text = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                Case 2
                grdTableTime.Text = gTime.strExpander
            End Select
        Next
    Next
            '������� ������������ ����
    Close intFileNum
    
            '������������ ���������� �������� � ������e ���������� �������
    intIntervalNum = grdTableTime.Rows
            '���� �� ���� ��������� ��������� "������� �������"
    For intVariantNew = 1 To gVarNumTime Step 1
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTime" + Trim(Str(intVariantNew)) + ".dat"
    
            '������������ ���������� �������� � ������� ���������� �������
        If FileLen(strPathFileName) / lngRecordLen + 1 > intIntervalNum Then
            intIntervalNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    Next
            '�������������� ����������� ������� ���������� �������
ReDim gInterval(gVarNumTime + 1, intIntervalNum) As String * 11
            '�������������� ����������� ������� �������������� ����������
            '  � ���������� ��� ���� ��������� "������� �������"
ReDim gTerCal(gVarNumTime + 1, intIntervalNum) As String * 12
    
            '���� �� ���� ��������� "������� �������"
    For intVariantNew = 0 To gVarNumTime Step 1
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTime" + Trim(Str(intVariantNew)) + ".dat"
    
            '���������� "��������" �������� � ������� ������ �������
            '  ���������� �������
        intIntervalNum = FileLen(strPathFileName) / lngRecordLen + 1
        gInterval(intVariantNew, 0) = intIntervalNum
            '���������� "��������" �������� � ������� ������ �������
            '  �������������� ���������� � ����������
        gTerCal(intVariantNew, 0) = intIntervalNum
    
            '������� ���� ��� ������������� �������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� �������  �������� "������� �������"
        For intRowNum = 1 To intIntervalNum - 1 Step 1
            '������ ������ "������� �������" �� ����� � �����
            Get intFileNum, intRowNum, gTime
                 gInterval(intVariantNew, intRowNum) = Left(gTime.strTime, 2) + "." + Mid(gTime.strTime, 3, 2) + _
                 "-" + Mid(gTime.strTime, 5, 2) + "." + Mid(gTime.strTime, 7, 2)
                 gTerCal(intVariantNew, intRowNum) = gTime.strIntervalNum
        Next
            '������� ����
        Close intFileNum
    Next
    
            '���������� ��������/���������� ����� � "������� �������"
    gAddDelRowTableTime = 0
            '�������� ������� ��������� ��������� � "������� �������"
    gChangesTableTime = False
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableTime:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableTime Error !")
    
End Sub
            
            '��������� ��������� ������ � ������������ �������� "������� �������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '���� �� ���� ��������
    For intColNumber = 0 To grdTableTime.Cols - 1 Step 1
        grdTableTime.ColWidth(intColNumber) = 970
        grdTableTime.ColAlignment(intColNumber) = 0
    Next
    
End Sub


