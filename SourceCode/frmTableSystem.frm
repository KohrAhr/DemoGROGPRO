VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableSystem 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_system"
   ClientHeight    =   6810
   ClientLeft      =   1035
   ClientTop       =   1095
   ClientWidth     =   9060
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
   ScaleHeight     =   6810
   ScaleWidth      =   9060
   Begin VB.OptionButton optReset 
      Caption         =   "Reset"
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
      Left            =   6000
      TabIndex        =   32
      Top             =   2520
      Value           =   -1  'True
      Width           =   732
   End
   Begin VB.ListBox lstIndex 
      Enabled         =   0   'False
      Height          =   900
      ItemData        =   "frmTableSystem.frx":0000
      Left            =   6840
      List            =   "frmTableSystem.frx":0002
      TabIndex        =   31
      Top             =   2400
      Width           =   1815
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
      Height          =   3252
      Left            =   2040
      TabIndex        =   19
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optIndex 
         Caption         =   "Index"
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
         TabIndex        =   29
         Top             =   2280
         Width           =   1452
      End
      Begin VB.OptionButton optObject 
         Caption         =   "Object "
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
         Height          =   312
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.OptionButton optConsAddrTerm 
         Caption         =   "Constant or Address and Terminal"
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
         Height          =   612
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1452
      End
      Begin VB.OptionButton optType 
         Caption         =   "Type"
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
         TabIndex        =   21
         Top             =   1560
         Width           =   1452
      End
      Begin VB.OptionButton optAppendix 
         Caption         =   "Appendix"
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
         TabIndex        =   20
         Top             =   2880
         Width           =   1452
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "Type"
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
      Height          =   612
      Left            =   3960
      TabIndex        =   14
      Top             =   1560
      Width           =   4812
      Begin VB.OptionButton optConstant 
         Caption         =   "Constant"
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
         Top             =   240
         Value           =   -1  'True
         Width           =   1092
      End
      Begin VB.OptionButton optReader 
         Caption         =   "Reader"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   852
      End
      Begin VB.OptionButton optProcessor 
         Caption         =   "Processor"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optWriter 
         Caption         =   "Writer"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox txtObject 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5160
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtConstant 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5160
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtAddress 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7680
      TabIndex        =   11
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox txtAppendix 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5280
      TabIndex        =   10
      Top             =   3120
      Width           =   1212
   End
   Begin VB.TextBox txtTerm 
      Enabled         =   0   'False
      Height          =   288
      Left            =   8280
      TabIndex        =   9
      Top             =   960
      Width           =   252
   End
   Begin VB.ListBox lstObject 
      Enabled         =   0   'False
      Height          =   1950
      ItemData        =   "frmTableSystem.frx":0004
      Left            =   120
      List            =   "frmTableSystem.frx":0006
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
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
      TabIndex        =   6
      Top             =   1440
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
      Top             =   6240
      Width           =   1212
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
      Left            =   6600
      TabIndex        =   3
      Top             =   6240
      Width           =   1092
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
      Left            =   7800
      TabIndex        =   2
      Top             =   6240
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
      TabIndex        =   1
      Top             =   6240
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
      TabIndex        =   0
      Top             =   6240
      Width           =   1092
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableSystem 
      Height          =   2295
      Left            =   2160
      TabIndex        =   5
      Top             =   3840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   9
      Cols            =   5
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
   Begin VB.Line Line19 
      X1              =   3960
      X2              =   3960
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line Line18 
      X1              =   3960
      X2              =   3960
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   3960
      X2              =   6720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6720
      Y1              =   3000
      Y2              =   3480
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   6720
      X2              =   8760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   8760
      X2              =   8760
      Y1              =   3480
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   3960
      X2              =   8760
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   8880
      X2              =   8880
      Y1              =   3720
      Y2              =   120
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   120
      X2              =   8880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   8880
      X2              =   2040
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   6120
      Y2              =   3720
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   1560
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   5160
      X2              =   7440
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblIndex 
      Alignment       =   2  'Center
      Caption         =   "Index "
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
      Left            =   4080
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblObject 
      Alignment       =   2  'Center
      Caption         =   "Object "
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
      Left            =   4080
      TabIndex        =   28
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblConstant 
      Alignment       =   2  'Center
      Caption         =   "Constant "
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
      Left            =   4080
      TabIndex        =   27
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      Caption         =   "01-15 Addr. "
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
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   7440
      X2              =   7440
      Y1              =   1440
      Y2              =   360
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8760
      X2              =   8760
      Y1              =   1560
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   7440
      X2              =   8760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lblAppendix 
      Alignment       =   2  'Center
      Caption         =   "Appendix "
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
      Left            =   4080
      TabIndex        =   25
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblTerm 
      Alignment       =   2  'Center
      Caption         =   " 0-3 Term. "
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
      Height          =   495
      Left            =   8160
      TabIndex        =   24
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblObjects 
      Alignment       =   2  'Center
      Caption         =   "Objects "
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
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   5880
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   6120
      Y2              =   6120
   End
End
Attribute VB_Name = "frmTableSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������� ����� �������������� ������ "��������� �������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "��������� �������"
Dim intColNumCorr As Integer
            '������� ����� �����
Dim intFileNum As Integer
            '������ "��������� �������"
Dim gSystem As SystemInfo

            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            '���� �� ����������� ��������� � "��������� �������"
    If gChangesTableSystem = True Then
            '������ �������� ������
       frmDemo.BeepSound
            '���� �������� � �������� ���������� "��������� �������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '���������� "��������� �������" � ����� �� ���������
            cmdSave_Click
        End If
    End If
    
            '������� ������������ �������� ���������� ���������� "��������� �������"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            '�������� ��������� ����
    txtObject.Text = ""
    txtConstant.Text = ""
    txtAddress.Text = ""
    txtTerm.Text = ""
    txtAppendix.Text = ""
            '�������� ������ ��������
    lstObject.Clear
    lstIndex.Clear
    
            '�������� ������� ��������� ��������� � "��������� �������"
    gChangesTableSystem = False
            '������� ��������� ������� �����
    frmTableSystem.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '���������
Private Sub cmdCorrection_Click()

            ' "��������� �������" �� �������� ��������������� �����
    If grdTableSystem.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        MsgBox ("The table is empty")
    
    Else
            '������� ���������� ��������� ����. ���������� ���������� "��������� �������"
        fraColName.Enabled = True
        optObject.Enabled = True
        optObject.Value = True
        optConsAddrTerm.Enabled = True
        optType.Enabled = True
        optIndex.Enabled = True
        optAppendix.Enabled = True
        lblObject.Enabled = True
        txtObject.Enabled = True
        lblObjects.Enabled = True
        lstObject.Enabled = True
            '�������� ��������� ����
        txtObject.Text = ""
        txtConstant.Text = ""
        txtAddress.Text = ""
        txtTerm.Text = ""
        txtAppendix.Text = ""
            '�������� ������
        lstObject.Clear
        lstIndex.Clear
    
            '������� "Objects"
        grdTableSystem.Col = 0
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNumCorr = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
            grdTableSystem.Row = intRowNumCorr
            '���������� ������ "lstObject" �������� �� "��������� �������"
            lstObject.AddItem grdTableSystem.Text
            '���������� ������ "lstIndex" �������� �� "��������� �������"
            lstIndex.AddItem grdTableSystem.Text
        Next
            '�������  ������� ������
        lstObject.ListIndex = 0
        lstIndex.ListIndex = 0
            '����� �������������� ������ - (1)
        intRowNumCorr = 1
        grdTableSystem.Row = intRowNumCorr
            '�������� �����
        optObject_Click
    
    End If
    
End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "��������� �������"
Private Sub grdTableSystem_Click()
            '��������� "��������"
    If lstObject.Enabled = True Then
            '����� �������������� ������ "��������� �������"
        intRowNumCorr = grdTableSystem.RowSel
        grdTableSystem.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstObject.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "��������� �������"
        intColNumCorr = grdTableSystem.ColSel
        grdTableSystem.Col = intColNumCorr
            '����� �������������� ������ "��������� �������"
        lstObject_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstObject.Left, Y:=lstObject.Top
            '����� ��������������� ������� "��������� �������"
        Select Case intColNumCorr
            Case 1
            If optConsAddrTerm.Value = True Then
                optConsAddrTerm_Click
            Else
                optConsAddrTerm.Value = True
            End If
            Case 2
            optType.Value = True
            Case 3
            optIndex.Value = True
            Case 4
            optAppendix.Value = True
        End Select
    End If
        
End Sub

            '����� �������������� ������ "��������� �������"
Private Sub lstObject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� �������������� ������ "��������� �������"
        intRowNumCorr = lstObject.ListIndex + 1
        grdTableSystem.Row = intRowNumCorr
        grdTableSystem.Col = 0
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
        txtObject.Text = grdTableSystem.Text
            '����� �������������� ������� "��������� �������" - "Type"
        grdTableSystem.Col = 2
            '����������� ����� "Constant"
        If Left(grdTableSystem.Text, 2) = "00" Then
            '����� ��������������� ������� "��������� �������"
            grdTableSystem.Col = 1
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
            txtConstant.Text = grdTableSystem.Text
            '������ ����������� ��� ��������� ��������� �����
            txtAddress.Text = ""
            txtTerm.Text = ""
            '�� ����������� ����� "Constant"
        Else
            '����� ��������������� ������� "��������� �������"
            grdTableSystem.Col = 1
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
            txtAddress.Text = Left(grdTableSystem.Text, 2)
            txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
            '������ ������������ ��� ��������� ���������� ����
            txtConstant.Text = ""
        End If
        grdTableSystem.Col = 3
            '����� ������ "��������� �������", �� ������� ��������� ������ "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            lstIndex.ListIndex = grdTableSystem.Text - 1
            '������ "Index" ������
        Else
            lstIndex.ListIndex = 0
        End If
        grdTableSystem.Col = 4
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
        txtAppendix.Text = grdTableSystem.Text
            '������������ ����� ��������������� �������"��������� �������"
        grdTableSystem.Col = intColNumCorr
    End If
    
End Sub

            '������� ����� - "Object"
Private Sub optObject_Click()
            '����� ��������������� ������� "��������� �������"
    intColNumCorr = 0
    grdTableSystem.Col = intColNumCorr
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
    txtObject.Text = grdTableSystem.Text
            '������� (��)���������� ��������� �������� ������. ���������� "��������� �������"
    lblObject.Enabled = True
    txtObject.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtObject.SetFocus
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
            '����� �������������� ������� "��������� �������" - "Type"
    grdTableSystem.Col = 2
            '����������� ����� "Constant"
    If Left(grdTableSystem.Text, 2) = "00" Then
            '����� ������������� ������� "��������� �������"
        grdTableSystem.Col = 1
            '����������� ������ "��������� �������" � ��������� ���� ��� �����������
        txtConstant.Text = grdTableSystem.Text
            '������� ����������� ��� ����������� ��������� �����
        txtAddress.Text = ""
        txtTerm.Text = ""
            '�� ����������� ����� "Constant"
    Else
            '����� ������������� ������� "��������� �������"
        grdTableSystem.Col = 1
            '����������� ������ "��������� �������" � ��������� ���� ��� �����������
        txtConstant.Text = grdTableSystem.Text
            '����������� ������ "��������� �������" � ��������� ���� ��� �����������
        txtAddress.Text = Left(grdTableSystem.Text, 2)
        txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
            '������� ������������ ��� ����������� ���������� ����
        txtConstant.Text = ""
    End If
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            '����� ������������� ������� "��������� �������"
    grdTableSystem.Col = 4
            '����������� ������ "��������� �������" � ��������� ���� ��� �����������
    txtAppendix.Text = grdTableSystem.Text
    
            '������������ ����� ��������������� ������� "��������� �������"
    grdTableSystem.Col = intColNumCorr

End Sub
            
            '������� ����� - "ConsAddrTerm"
Private Sub optConsAddrTerm_Click()
            '����� ��������������� ������� "��������� �������"
    intColNumCorr = 1
    grdTableSystem.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "��������� �������"
    lblObject.Enabled = False
    txtObject.Enabled = False
            '����� �������������� ������� "��������� �������" - "Type"
    grdTableSystem.Col = 2
            '����������� ����� "Constant"
    If Left(grdTableSystem.Text, 2) = "00" Then
            '����� ��������������� ������� "��������� �������"
        grdTableSystem.Col = intColNumCorr
        lblConstant.Enabled = True
        txtConstant.Enabled = True
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblTerm.Enabled = False
        txtTerm.Enabled = False
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
        txtConstant.Text = grdTableSystem.Text
                '���������� ����� �� ��������� ���� ��� ���������
        txtConstant.SetFocus
            '������ ����������� ��� ��������� ��������� �����
        txtAddress.Text = ""
        txtTerm.Text = ""
            '�� ����������� ����� "Constant"
    Else
            '����� ��������������� ������� "��������� �������"
        grdTableSystem.Col = intColNumCorr
        lblConstant.Enabled = False
        txtConstant.Enabled = False
        lblAddress.Enabled = True
        txtAddress.Enabled = True
        lblTerm.Enabled = True
        txtTerm.Enabled = True
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
        txtAddress.Text = Left(grdTableSystem.Text, 2)
        txtTerm.Text = Mid(grdTableSystem.Text, 3, 1)
                '���������� ����� �� ��������� ���� ��� ���������
        txtAddress.SetFocus
            '������ ������������ ��� ��������� ���������� ����
        txtConstant.Text = ""
    End If
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub

            '������� ����� - "Type"
Private Sub optType_Click()
            '����� ��������������� ������� "��������� �������"
    intColNumCorr = 2
    grdTableSystem.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "��������� �������"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = True
    optConstant.Enabled = True
            '�������� ����� "Constant"
    optConstant.Value = True
    optReader.Enabled = True
    optWriter.Enabled = True
    optProcessor.Enabled = True
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub
            
            '������� ����� - "Index"
Private Sub optIndex_Click()
            '����� ��������������� ������� "��������� �������"
    intColNumCorr = 3
    grdTableSystem.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "��������� �������"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = True
            '����� ������ "��������� �������", �� ������� ��������� ������ "Index"
    If Trim(grdTableSystem.Text) <> "" Then
        lstIndex.ListIndex = grdTableSystem.Text - 1
            '������ "Index" ������
    Else
        lstIndex.ListIndex = 0
    End If
    optReset.Enabled = True
    lstIndex.Enabled = True
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False

End Sub

            '������� ����� "Appendix"
Private Sub optAppendix_Click()
            '����� ��������������� ������� "��������� �������"
    intColNumCorr = 4
    grdTableSystem.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "��������� �������"
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
            '����������� ������ "��������� �������" � ��������� ���� ��� ���������
    txtAppendix.Text = grdTableSystem.Text
    lblAppendix.Enabled = True
    txtAppendix.Enabled = True
                '���������� ����� �� ��������� ���� ��� ���������
    txtAppendix.SetFocus

End Sub
            
            '��������� ����� � ������� ��������������� ����� "Object"
Private Sub txtObject_KeyPress(KeyAscii As Integer)
            '��� �������
    If KeyAscii = vbKeyReturn Then
            '��� � ���������� ���������
        If Len(Trim(txtObject.Text)) < 17 Then
            '��������� ����� "Object" � "��������� �������"
          grdTableSystem.Text = Trim(txtObject.Text)
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '�������� ����� "optConsAddrTerm"
            optConsAddrTerm.Value = True
            '��� � ������������ ���������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Constant"
Private Sub txtConstant_KeyPress(KeyAscii As Integer)
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '��������� � ���������� ���������
        If Len(Trim(txtConstant.Text)) < 17 Then
            '��������� ������ "Cons.,Addr.,Term." � "��������� �������"
            grdTableSystem.Text = Trim(txtConstant.Text)
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '�������� ����� "optType"
            optType.Value = True
            '��������� � ������������ ���������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            '����� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo AddressError
            '����� � ���������� ��������� ������� (01/15,  00 - ��������� �����)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text > 0 And txtAddress.Text < 16 Then
            '��������� ������ "Cons.,Addr.,Term." � "��������� �������"
            If Len(Trim(grdTableSystem.Text)) < 3 Then
            '��������� ������ "Cons.,Addr.,Term." � "��������� �������"
                txtTerm.Text = "0"
                grdTableSystem.Text = Trim(txtAddress.Text) + Trim(txtTerm.Text)
            Else
                grdTableSystem.Text = Trim(txtAddress.Text) + Mid(grdTableSystem.Text, 3)
            End If
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '���������� ����� �� ��������� ���� "txtTerm"
            txtTerm.SetFocus
            Exit Sub
            '������ � ������������ ��������� �������
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            '��������� ����� � ������� ��������������� "Term"
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
            '����� ��������� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo TermError
            '����� ��������� � ���������� ��������� (0/3)
        If Len(Trim(txtTerm.Text)) = 1 And txtTerm.Text >= 0 And txtTerm.Text < 4 Then
            '��������� ������ "Cons.,Addr.,Term." � "��������� �������"
            
            If Len(Trim(grdTableSystem.Text)) < 3 Then
            '��������� ������ "Cons.,Addr.,Term." � "��������� �������"
                txtAddress.Text = "01"
                grdTableSystem.Text = Trim(txtAddress.Text) + Trim(txtTerm.Text)
            Else
                grdTableSystem.Text = Left(grdTableSystem.Text, 2) + Trim(txtTerm.Text) + _
                Mid(grdTableSystem.Text, 4)
            End If
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '�������� ����� "optType"
            optType.Value = True
            Exit Sub
            '����� ��������� � ������������ ���������
TermError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            '������� ����� - "Constant"
Private Sub optConstant_GotFocus()
            '��������� ������ "Type" � "��������� �������"
    grdTableSystem.Text = "00 - Constant"
            '���������� �������  ��������� ��������� � "��������� �������"
    gChangesTableSystem = True

End Sub

            '������� ����� - "Reader"
Private Sub optReader_GotFocus()
            '��������� ������ "Type" � "��������� �������"
    grdTableSystem.Text = "01 - Reader"
            '���������� �������  ��������� ��������� � "��������� �������"
    gChangesTableSystem = True

End Sub

            '������� ����� - "Writer"
Private Sub optWriter_GotFocus()
            '��������� ������ "Type" � "��������� �������"
    grdTableSystem.Text = "02 - Writer"
            '���������� �������  ��������� ��������� � "��������� �������"
    gChangesTableSystem = True

End Sub

            '������� ����� - "Processor"
Private Sub optProcessor_GotFocus()
            '��������� ������ "Type" � "��������� �������"
    grdTableSystem.Text = "03 - Processor"
            '���������� �������  ��������� ��������� � "��������� �������"
    gChangesTableSystem = True

End Sub

            '������� ����� - "Reset"
Private Sub optReset_GotFocus()
            '�� ������ ������� ������ "Index"
    If Trim(grdTableSystem.Text) <> "" Then
            '����� ������ "��������� �������", �� ������� ��������� ������� ������ "Index"
        grdTableSystem.Row = grdTableSystem.Text
            '������ "��������� �������", �� ������� ��������� ������� ������ "Index"
            '  ���� �������� �� ������ ������ "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            '���������� ����������� "�������" � ��������� ���������
            lstIndex.ListIndex = grdTableSystem.Text - 1
            grdTableSystem.Row = intRowNumCorr
            '��������� ����������� ������� ������ "Index" � "��������� �������"
            grdTableSystem.Text = lstIndex.ListIndex + 1
        Else
            '������� ������� ������ "Index" � "��������� �������"
            grdTableSystem.Row = intRowNumCorr
            grdTableSystem.Text = ""
            lstIndex.ListIndex = 0
        End If
            '���������� �������  ��������� ��������� � "��������� �������"
        gChangesTableSystem = True
    End If

End Sub

            '����� ������ "��������� �������", �� ������� ����� ��������� ������ "Index"
Private Sub lstIndex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = 1 Then
            '������ ����� ������ - ������ �� ��������� ��� �� ����
        If lstIndex.ListIndex + 1 <> grdTableSystem.Row Then
            '��������� ������ "Index" � "��������� �������"
            grdTableSystem.Text = lstIndex.ListIndex + 1
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '�������� ������ ������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� ���� "Appendix"
Private Sub txtAppendix_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtAppendix.Text)) < 9 Then
            '��������� ���� "Appendix" � "��������� �������"
            grdTableSystem.Text = Trim(txtAppendix.Text)
            '���������� �������  ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '���������� ����� �� ������ "Save"
            cmdSave.SetFocus
            '�������� ������ ������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '���������� ������ � "��������� �������"
Private Sub cmdAdd_Click()
            '������ ������ "��������� �������"
Dim strSystem As String
    strSystem = ""
    
            '������� ������������ �������� ���������� ���������� "��������� �������"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            '�������� ������ ��������
    lstObject.Clear
    lstIndex.Clear
    
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ ��� �������
    strSystem = InputBox("Objects Name:", "Add ...")
            '��� ������� �� �������
    If strSystem = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox " The object isn't selected !"
            '������� ������� ��� ���������
    Else
            '���������� ������ � ����� "��������� �������"
        grdTableSystem.AddItem strSystem
            '���������� ��������/���������� ����� � "��������� �������"
        gAddDelRowTableSystem = gAddDelRowTableSystem + 1
            '���������� ������� ��������� ��������� � "��������� �������"
        gChangesTableSystem = True
    End If
            '���������� ����� �� ������ "Add"
    cmdAdd.SetFocus
    
End Sub
            
            '�������� ������ �� "��������� �������"
Private Sub cmdDelete_Click()
            '������� ����� ��������������� ������ "��������� �������"
Dim intRowNum As Integer
Dim intRowNumSys As Integer
            '������� ����� ������� "��������� �������"
Dim intColNum As Integer

            '������� ������������ �������� ���������� ���������� "��������� �������"
    fraColName.Enabled = False
    optObject.Enabled = False
    optConsAddrTerm.Enabled = False
    optType.Enabled = False
    optIndex.Enabled = False
    optAppendix.Enabled = False
    lblObjects.Enabled = False
    lstObject.Enabled = False
    lblObject.Enabled = False
    txtObject.Enabled = False
    lblConstant.Enabled = False
    txtConstant.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblTerm.Enabled = False
    txtTerm.Enabled = False
    fraType.Enabled = False
    optConstant.Enabled = False
    optReader.Enabled = False
    optWriter.Enabled = False
    optProcessor.Enabled = False
    lblIndex.Enabled = False
    optReset.Enabled = False
    lstIndex.Enabled = False
    lblAppendix.Enabled = False
    txtAppendix.Enabled = False
            '�������� ������ ��������
    lstObject.Clear
    lstIndex.Clear
            
            '��������� (�� ���������) ����� "frmSelectRow"
    Load frmSelectRow
            '���������������� �������� "lblColName" ����� "frmSelectRow"
    frmSelectRow.lblColName.Caption = "System"
    
            '������� "Objects"
    grdTableSystem.Col = 0
             '�������� ������ ��������
    frmSelectRow.lstSelectRow.Clear
           '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        grdTableSystem.Row = intRowNum
            '���������� ������ "lstSelectRow" �������� �� "��������� �������"
        frmSelectRow.lstSelectRow.AddItem grdTableSystem.Text
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
            '�������� ������ �� "��������� �������"
            '��������� ������ �� ����� ���� �������
    ElseIf frmSelectRow.lstSelectRow.ListCount = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The last row isn't selected !"
            '�������� ������ �� "������� �������"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            '����� ��������� ������
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            
            '������� "Index"
        grdTableSystem.Col = 3
            '�� ������ ������� ������ "Index"
        If Trim(grdTableSystem.Text) <> "" Then
            '������ �������� ������
           frmDemo.BeepSound
            '�������� ����������
            MsgBox "Deletion impossible. The Index isn't empty !"
        Else
            '�������� ������
            grdTableSystem.RemoveItem intRowNum
            '���������� ��������/���������� ����� � "��������� �������"
            gAddDelRowTableSystem = gAddDelRowTableSystem - 1
            '���������� ������� ��������� ��������� � "��������� �������"
            gChangesTableSystem = True
            '���� �� ���� ��������. ������� "��������� �������" - ��������� ����� "Index"
            For intRowNumSys = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
                grdTableSystem.Row = intRowNumSys
            '������ �������� ���������
                If Trim(grdTableSystem.Text) <> "" Then
            '������� ������ "Index", ������� ��������� �� ��������� ������ "��������� �������"
                    If grdTableSystem.Text = intRowNum Then grdTableSystem.Text = ""
            '��������� �� -1 ���������� �������� ����� "Index", ������� ��������� �� ������
            '  ����� "��������� �������" �������, ��� ����� ��������� ������
                    If Trim(grdTableSystem.Text) > intRowNum _
                    Then grdTableSystem.Text = grdTableSystem.Text - 1
                End If
            Next
        End If
            
    End If
            '��������� ����� "frmSelectRow"
    UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
    Set frmSelectRow = Nothing
            '���������� ����� �� ������ "Delete"
    cmdDelete.SetFocus
    
End Sub
            
            '���������� "��������� �������" � ����� �� ���������
Public Function SaveTableSystem()
    Call cmdSave_Click
    SaveTableSystem = 0
    
End Function
            
            '���������� "��������� �������" � ����� �� ���������
Private Sub cmdSave_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "��������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "��������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "��������� �������"
Dim intColNum As Integer
            '��������� ����� ������ (������) "��������� �������"
    lngRecordLen = Len(gSystem)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableSystem.dat"
    
            '�����, ��������� �� "��������� �������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
    If gAddDelRowTableSystem < 0 Then
            '������� "������" ������������ ����
        Kill strPathFileName
    End If
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        grdTableSystem.Row = intRowNum
            '�� ���� �������� "��������� �������"
        For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "��������� �������" � ����
            Select Case intColNum
                Case 0
                gSystem.strObject = grdTableSystem.Text
                Case 1
                gSystem.strConsAddrTerm = grdTableSystem.Text
                Case 2
                gSystem.strType = Left(grdTableSystem.Text, 2)
                Case 3
                gSystem.strIndex = Left(grdTableSystem.Text, 5)
                Case 4
                gSystem.strAppendix = grdTableSystem.Text
            End Select
        Next
            '�������� ������ "��������� �������" � ����
        Put intFileNum, intRowNum, gSystem
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "��������� �������"
    gAddDelRowTableSystem = 0
            '�������� ������� ��������� ��������� � "��������� �������"
    gChangesTableSystem = False
            
End Sub
            
            '���������� "��������� �������" � ���������� �����
Private Sub cmdSaveAs_Click()
            '������ ��� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "��������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "��������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "��������� �������"
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
            '������ "��������� �������" � ��������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "��������� �������"
        lngRecordLen = Len(gSystem)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '�����, ��������� �� "��������� �������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
        If gAddDelRowTableSystem < 0 Then
            '������� "������" ����, ���� �� ����������
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            '������� ��������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "��������� �������"
        For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
            grdTableSystem.Row = intRowNum
            '�� ���� �������� "��������� �������"
            For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
                grdTableSystem.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "��������� �������" � ����
                Select Case intColNum
                    Case 0
                    gSystem.strObject = grdTableSystem.Text
                    Case 1
                    gSystem.strConsAddrTerm = grdTableSystem.Text
                    Case 2
                    gSystem.strType = Left(grdTableSystem.Text, 2)
                    Case 3
                    gSystem.strIndex = Left(grdTableSystem.Text, 5)
                    Case 4
                    gSystem.strAppendix = grdTableSystem.Text
                End Select
            Next
            '�������� ������ "��������� �������" � ����
            Put intFileNum, intRowNum, gSystem
        Next
            '������� ��������� ����
        Close intFileNum
             '���������� ��������/���������� ����� � "��������� �������"
        gAddDelRowTableSystem = 0
               '�������� ������� ��������� ��������� � "��������� �������"
        gChangesTableSystem = False
    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '���������� ����� �� ������ "Cancel"
    cmdCancel.SetFocus
    
End Sub

            '�������� ����� "��������� �������"
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "��������� �������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "��������� �������"
Dim intRowNum As Integer
            '������� ����� ������� "��������� �������"
Dim intColNum As Integer

            '���������� ������ ��������
    SetColWidth
            '������� ������ = 0 (��������� ��������)
    grdTableSystem.Row = 0
    grdTableSystem.Col = 0
    grdTableSystem.Text = "Objects"
            '�������� � ������ (������ 0, ������� 1)
    grdTableSystem.Col = 1
    grdTableSystem.Text = "Cons.,Addr.,Term."
            '�������� � ������ (������ 0, ������� 2)
    grdTableSystem.Col = 2
    grdTableSystem.Text = "Type"
            '�������� � ������ (������ 0, ������� 3)
    grdTableSystem.Col = 3
    grdTableSystem.Text = "Index"
            '�������� � ������ (������ 0, ������� 4)
    grdTableSystem.Col = 4
    grdTableSystem.Text = "Appendix"
    
            
            '���������� "��������� �������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "��������� �������"
    lngRecordLen = Len(gSystem)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableSystem.dat"
                
            '���� ����������� - ?
    On Error GoTo ErrorTableSystem
                '���������� ����� "��������� �������" ����� ������� ����� �� ��������� +1
    grdTableSystem.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "��������� �������"
    For intRowNum = 1 To grdTableSystem.Rows - 1 Step 1
            '������� ������ "��������� �������"
        grdTableSystem.Row = intRowNum
            '������ ������ "��������� �������" �� ����� � �����
        Get intFileNum, intRowNum, gSystem
            '�� ���� �������� "��������� �������"
        For intColNum = 0 To grdTableSystem.Cols - 1 Step 1
            '������� ������� "��������� �������"
            grdTableSystem.Col = intColNum
            '���������� ������� ������ "��������� �������" �� ������
            Select Case intColNum
                Case 0
                grdTableSystem.Text = gSystem.strObject
                Case 1
                grdTableSystem.Text = gSystem.strConsAddrTerm
                Case 2
                grdTableSystem.Text = gSystem.strType
                If gSystem.strType = "00" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Constant"
                If gSystem.strType = "01" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Reader"
                If gSystem.strType = "02" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Writer"
                If gSystem.strType = "03" Then grdTableSystem.Text = grdTableSystem.Text + _
                " - Processor"
                Case 3
                grdTableSystem.Text = Left(gSystem.strIndex, 5)
                Case 4
                grdTableSystem.Text = gSystem.strAppendix
            End Select
        Next
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "��������� �������"
    gAddDelRowTableSystem = 0
            '�������� ������� ��������� ��������� � "��������� �������"
    gChangesTableSystem = False
    
    Exit Sub
ErrorTableSystem:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableSystem Error !")
    
End Sub
            
            '��������� ��������� ������ � ������������ �������� "��������� �������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '���� �� ���� ��������
    For intColNumber = 0 To grdTableSystem.Cols - 1 Step 1
        grdTableSystem.ColWidth(intColNumber) = 1600
        grdTableSystem.ColAlignment(intColNumber) = 0
    Next
    
End Sub
