VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableTerminal 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_terminal"
   ClientHeight    =   6390
   ClientLeft      =   2400
   ClientTop       =   1485
   ClientWidth     =   7110
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
   ScaleHeight     =   6390
   ScaleWidth      =   7110
   Begin VB.TextBox txtVariant 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4200
      TabIndex        =   31
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmdDefaultTerm 
      Cancel          =   -1  'True
      Caption         =   "DefaultTerm"
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
      Left            =   5760
      TabIndex        =   30
      Top             =   5760
      Width           =   1212
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
      Left            =   5880
      TabIndex        =   29
      Top             =   5040
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
      Left            =   4680
      TabIndex        =   28
      Top             =   5040
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
      TabIndex        =   27
      Top             =   5040
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
      TabIndex        =   26
      Top             =   5040
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
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
      TabIndex        =   25
      Top             =   5760
      Width           =   1212
   End
   Begin VB.HScrollBar hsbVariant 
      Height          =   252
      Left            =   2520
      Max             =   0
      TabIndex        =   20
      Top             =   3000
      Width           =   1092
   End
   Begin VB.TextBox txtExpander 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5520
      TabIndex        =   17
      Top             =   2280
      Width           =   1212
   End
   Begin VB.TextBox txtDescription 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5520
      TabIndex        =   15
      Top             =   1800
      Width           =   1212
   End
   Begin VB.TextBox txtPort 
      Enabled         =   0   'False
      Height          =   288
      Left            =   6480
      TabIndex        =   12
      Top             =   1320
      Width           =   252
   End
   Begin VB.TextBox txtAddress 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
      Width           =   372
   End
   Begin VB.TextBox txtTerminal 
      Enabled         =   0   'False
      Height          =   288
      Left            =   5280
      TabIndex        =   9
      Top             =   480
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
      Height          =   2412
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optDescription 
         Caption         =   "Description"
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
         TabIndex        =   7
         Top             =   1560
         Width           =   1452
      End
      Begin VB.OptionButton optAddrPort 
         Caption         =   "Address and Port"
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
         Height          =   372
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1452
      End
      Begin VB.OptionButton optTerminal 
         Caption         =   "Terminal "
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
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1452
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
         TabIndex        =   4
         Top             =   2040
         Width           =   1452
      End
   End
   Begin VB.ListBox lstTerminal 
      Enabled         =   0   'False
      Height          =   1320
      ItemData        =   "frmTableTerminal.frx":0000
      Left            =   120
      List            =   "frmTableTerminal.frx":0002
      TabIndex        =   1
      Top             =   3480
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1092
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableTerminal 
      Height          =   1455
      Left            =   2160
      TabIndex        =   19
      Top             =   3480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   1
      Cols            =   4
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
   Begin VB.Label lblPointer 
      Alignment       =   2  'Center
      Caption         =   "<==="
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
      Left            =   4800
      TabIndex        =   24
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblVariant 
      Alignment       =   2  'Center
      Caption         =   "Terminals variant"
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
      Left            =   5280
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
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
      Top             =   3000
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
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   2760
      Y2              =   4920
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4680
      Y2              =   4920
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3480
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   6960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2040
      X2              =   6960
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblExpander 
      Alignment       =   2  'Center
      Caption         =   "Expander "
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
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Description "
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
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      Caption         =   " 2-8 Port "
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
      Left            =   6480
      TabIndex        =   14
      Top             =   840
      Width           =   375
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
      Left            =   5880
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblAddrAndPort 
      Alignment       =   2  'Center
      Caption         =   "Addr.  and  Port "
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
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblTerminal 
      Alignment       =   2  'Center
      Caption         =   "Terminal "
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
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblTerminals 
      Alignment       =   2  'Center
      Caption         =   "Terminals "
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmTableTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������� ����� �������������� ������ "������� ����������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "������� ����������"
Dim intColNumCorr As Integer
            '"������" ����� �������� "������� ����������"
Dim intVariantOld As Integer
            '"�����" ����� �������� "������� ����������"
Dim intVariantNew As Integer
            '������� ����� �����
Dim intFileNum As Integer
            '������ "������� ����������"
Dim gTerminal As TerminalInfo

            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            '"������" ����� �������� "������� ����������" �� �������
    If hsbVariant.Value <> 0 Then
    
            '���� �� ����������� ��������� � "������� ����������"
        If gChangesTableTerminal = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ����������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� ������� ���������� � ����� �� ���������
                cmdSave_Click
            End If
        End If
            
    Else
    
            '���� �� ����������� ��������� � "������� ����������"
        If gChangesTableTerminal = True Then
            '������ �������� ������
            frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ����������" - �� �����
            intButtonsAndIcons = vbYesNo + vbQuestion
            strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
            If strResponse = vbYes Then
            '���������� ������� ���������� � ����� �� ���������
                cmdSave_Click
            End If
        End If
            
            '������� ������������ �������� ����������
            '  ���������� "������� ����������"
        fraColName.Enabled = False
        optTerminal.Enabled = False
        optAddrPort.Enabled = False
        optDescription.Enabled = False
        optExpander.Enabled = False
        lblTerminal.Enabled = False
        txtTerminal.Enabled = False
        lblAddrAndPort.Enabled = False
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblDescription.Enabled = False
        txtDescription.Enabled = False
        lblExpander.Enabled = False
        txtExpander.Enabled = False
        lblTerminals.Enabled = False
        lstTerminal.Enabled = False
    End If
            
            '�������� ��������� ����
    txtTerminal.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtDescription.Text = ""
    txtExpander.Text = ""
            '�������� ������ ����������
    lstTerminal.Clear
            
            '������� ������� ������� ������� "������� ����������"
    intVariantOld = 0
    hsbVariant.Value = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableTerminal = False
            '������� ��������� ������� �����
    frmTableTerminal.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '���������
Private Sub cmdCorrection_Click()

            ' "������� ����������" �� �������� ��������������� �����
    If grdTableTerminal.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        MsgBox ("The table is empty")
    
    Else
            '������� ���������� ��������� �������� ����������
            '  ���������� "������� ����������"
        fraColName.Enabled = True
        optTerminal.Enabled = True
        optTerminal.Value = True
        optAddrPort.Enabled = True
        optDescription.Enabled = True
        optExpander.Enabled = True
        lblTerminal.Enabled = True
        txtTerminal.Enabled = True
        lstTerminal.Enabled = True
            '�������� ��������� ����
        txtTerminal.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtDescription.Text = ""
        txtExpander.Text = ""
            '�������� ������ ����
        lstTerminal.Clear
    
            '������� "Terminal"
        grdTableTerminal.Col = 0
                '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNumCorr = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableTerminal.Row = intRowNumCorr
            '���������� ������ "lstName" �������� �� "������� ����������"
            lstTerminal.AddItem grdTableTerminal.Text
        Next
            '�������  ������� ������
        lstTerminal.ListIndex = 0
            '����� �������������� ������ - (1)
        intRowNumCorr = 1
        grdTableTerminal.Row = intRowNumCorr
            '�������� �����
        optTerminal_Click
    End If
    
End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "������� ����������"
Private Sub grdTableTerminal_Click()
            '��������� "��������"
    If lstTerminal.Enabled = True Then
            '����� �������������� ������ "������� ����������"
        intRowNumCorr = grdTableTerminal.RowSel
        grdTableTerminal.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstTerminal.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "������� ����������"
        intColNumCorr = grdTableTerminal.ColSel
        grdTableTerminal.Col = intColNumCorr
            '����� �������������� ������ "������� ����������"
        lstTerminal_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstTerminal.Left, Y:=lstTerminal.Top
            '����� ��������������� ������� "������� ����������"
        Select Case intColNumCorr
            Case 1
            optAddrPort.Value = True
            '���������� ����� �� ��������� ���� ��� ���������
            txtAddress.SetFocus
            Case 2
            optDescription.Value = True
            '���������� ����� �� ��������� ���� ��� ���������
            txtDescription.SetFocus
            Case 3
            optExpander.Value = True
            '���������� ����� �� ��������� ���� ��� ���������
            txtExpander.SetFocus
        End Select
    End If
        
End Sub

            '����� �������������� ������ "������� ����������"
Private Sub lstTerminal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� �������������� ������ "������� ����������"
        intRowNumCorr = lstTerminal.ListIndex + 1
        grdTableTerminal.Row = intRowNumCorr
        grdTableTerminal.Col = 0
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
        txtTerminal.Text = grdTableTerminal.Text
            '����� ��������������� ������� "������� ����������"
        grdTableTerminal.Col = 1
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
        txtAddress.Text = Left(grdTableTerminal.Text, 2)
        txtPort.Text = Mid(grdTableTerminal.Text, 3, 1)
        grdTableTerminal.Col = 2
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
        txtDescription.Text = grdTableTerminal.Text
        grdTableTerminal.Col = 3
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
        txtExpander.Text = grdTableTerminal.Text
            '������������ ����� ��������������� ������� "������� ����������"
        grdTableTerminal.Col = intColNumCorr
    End If

End Sub

            '������� ����� - "Terminal"
Private Sub optTerminal_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 0
    grdTableTerminal.Col = intColNumCorr
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
    txtTerminal.Text = grdTableTerminal.Text
            '������� (��)���������� ��������� ����. ������. ���������� "������� ����������"
    lblTerminal.Enabled = True
    txtTerminal.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtTerminal.SetFocus
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
            '����� ������������� ������� "������� ����������"
    grdTableTerminal.Col = 1
            '����������� ������ "������� ����������" � ��������� ���� ��� �����������
    txtAddress.Text = Left(Trim(grdTableTerminal.Text), 2)
    lblPort.Enabled = False
    txtPort.Enabled = False
            '����������� ������ "������� ����������" � ��������� ���� ��� �����������
    txtPort.Text = Mid(Trim(grdTableTerminal.Text), 3, 1)
    lblDescription.Enabled = False
    txtDescription.Enabled = False
            '����� ������������� ������� "������� ����������"
    grdTableTerminal.Col = 2
            '����������� ������ "������� ����������" � ��������� ���� ��� �����������
    txtDescription.Text = grdTableTerminal.Text
    lblExpander.Enabled = False
    txtExpander.Enabled = False
            '����� ������������� ������� "������� ����������"
    grdTableTerminal.Col = 3
            '����������� ������ "������� ����������" � ��������� ���� ��� �����������
    txtExpander.Text = grdTableTerminal.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableTerminal.Col = intColNumCorr

End Sub
            
            '������� ����� - "AddrPort"
Private Sub optAddrPort_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 1
    grdTableTerminal.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "������� ����������"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = True
    lblAddress.Enabled = True
    txtAddress.Enabled = True
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
    txtAddress.Text = Left(grdTableTerminal.Text, 2)
    txtPort.Text = Mid(grdTableTerminal.Text, 3, 1)
            '���������� ����� �� ��������� ���� ��� ���������
    txtAddress.SetFocus
    lblPort.Enabled = True
    txtPort.Enabled = True
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False

End Sub
            
            '������� ����� - "Description"
Private Sub optDescription_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 2
    grdTableTerminal.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "������� ����������"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = True
    txtDescription.Enabled = True
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
    txtDescription.Text = grdTableTerminal.Text
            '���������� ����� �� ��������� ���� ��� ���������
    txtDescription.SetFocus
    lblExpander.Enabled = False
    txtExpander.Enabled = False

End Sub
            
            '������� ����� - "Expander"
Private Sub optExpander_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 3
    grdTableTerminal.Col = intColNumCorr
            '������� (��)���������� ��������� ����-� ������. ���������� "������� ����������"
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = True
    txtExpander.Enabled = True
            '����������� ������ "������� ����������" � ��������� ���� ��� ���������
    txtExpander.Text = grdTableTerminal.Text
            '���������� ����� �� ��������� ���� ��� ���������
    txtExpander.SetFocus

End Sub
            
            '��������� ����� � ������� ��������������� ����� "Terminal"
Private Sub txtTerminal_KeyPress(KeyAscii As Integer)
            '��� �������
    If KeyAscii = vbKeyReturn Then
            '��� � ���������� ���������
        If Len(Trim(txtTerminal.Text)) < 17 Then
            '��������� ����� "Terminal" � "������� ����������"
            grdTableTerminal.Text = Trim(txtTerminal.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableTerminal = True
            '�������� ����� "optAddrPort"
            optAddrPort.Value = True
            Exit Sub
            '��� � ������������ ���������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Address and Port - Address"
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
            '����� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo AddressError
            '����� � ���������� ��������� ������� (01/15,  00 - ��������� �����)
        If Len(Trim(txtAddress.Text)) = 2 And txtAddress.Text >= 0 _
        And txtAddress.Text < 16 Then
            '��������� ������ "Address and Port" � "������� ����������"
            grdTableTerminal.Text = Trim(txtAddress.Text) + Mid(grdTableTerminal.Text, 3)
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableTerminal = True
            '���������� ����� �� ��������� ���� "Port"
            txtPort.SetFocus
            Exit Sub
            '������ � ������������ ��������� �������
AddressError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If
    
End Sub
            
            '��������� ����� � ������� ��������������� "Address and Port - Port"
Private Sub txtPort_KeyPress(KeyAscii As Integer)
            '����� ����� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo PortError
            '����� ����� � ���������� ��������� (2/8)
        If Len(Trim(txtPort.Text)) = 1 And txtPort.Text > 1 And txtPort.Text < 9 Then
            '��������� ������ "Address and Port" � "������� ����������"
            If Len(Trim(grdTableTerminal.Text)) = 0 Then
                grdTableTerminal.Text = "01" + Trim(txtPort.Text)
            Else
                grdTableTerminal.Text = Left(grdTableTerminal.Text, 2) + Trim(txtPort.Text)
            End If
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableTerminal = True
            '�������� ����� "optDescription"
            optDescription.Value = True
            Exit Sub
            '����� ����� � ������������ ���������
PortError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� ���� "Description"
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtDescription.Text)) < 17 Then
            '��������� ���� "Expander" � "������� ����������"
            grdTableTerminal.Text = Trim(txtDescription.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableTerminal = True
            '�������� ����� "optExpander"
            optExpander.Value = True
            '�������� ������ ������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� ���� "Expander"
Private Sub txtExpander_KeyPress(KeyAscii As Integer)
            '���������� �������
    If KeyAscii = vbKeyReturn Then
        If Len(Trim(txtExpander.Text)) < 9 Then
            '��������� ���� "Expander" � "������� ����������"
            grdTableTerminal.Text = Trim(txtExpander.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableTerminal = True
            '���������� ����� �� ������ "Save"
            cmdSave.SetFocus
            '�������� ������ ������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub

            '��������� ������� "Change" - ��������� ��� �������� "Variant"
Private Sub hsbVariant_Change()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ����������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ����������"
Dim intColNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String

            '���������� ������ ��������
    SetColWidth
            '������� ������������ �������� ���������� ���������� "������� ����������"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            '�������� ��������� ����
    txtTerminal.Text = ""
    txtAddress.Text = ""
    txtPort.Text = ""
    txtDescription.Text = ""
    txtExpander.Text = ""
            '�������� ������ ����������
    lstTerminal.Clear
            
            '���� ������������� ��������� � "������� ����������"
    If gChangesTableTerminal = True Then
            '������ �������� ������
        frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ����������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '��������� "�����" ����� �������� "������� ����������"
            intVariantNew = hsbVariant.Value
            '"������" ����� �������� "������� ����������"
            hsbVariant.Value = intVariantOld
            '���������� "������� ����������" � ����� �� ���������
            cmdSave_Click
            '������������ "�����" ����� �������� "������� ����������"
            hsbVariant.Value = intVariantNew
        End If
    End If
            '"������" ����� �������� "������� ����������"
    intVariantOld = hsbVariant.Value
            '���������� ��������/���������� ����� � "������� ����������"
    gAddDelRowTableTerminal = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableTerminal = False

            '���������� �������� "������� ����������" �� �����
            
            '��������� ����� ������ (������) "������� ����������"
    lngRecordLen = Len(gTerminal)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + _
    Trim(Str(hsbVariant.Value)) + ".dat"
                
                                
            '���� ����������� - ?
    On Error GoTo ErrorTableTerminal
            '���������� ����� "������� ����������" ����� ������� ����� �� ��������� +1
    grdTableTerminal.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� �������  �������� "������� ����������"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableTerminal.Row = intRowNum
            '������ ������ "������� ����������" �� ����� � �����
        Get intFileNum, intRowNum, gTerminal
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            grdTableTerminal.Col = intColNum
            '���������� ������� ������ "������� ����������" �� ������
            Select Case intColNum
                Case 0
                grdTableTerminal.Text = gTerminal.strTerminal
                Case 1
                 grdTableTerminal.Text = gTerminal.strAddrPort
                Case 2
                grdTableTerminal.Text = gTerminal.strDescription
                Case 3
                grdTableTerminal.Text = gTerminal.strExpander
            End Select
        Next
    Next
            '������� ����
    Close intFileNum
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = hsbVariant.Value
            '���������� ����� �� ������ "Correction"
    If frmTableTerminal.Visible = True Then cmdCorrection.SetFocus
    Exit Sub
ErrorTableTerminal:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableTerminal Error !")
    
End Sub
            
            '���������� ������ � "������� ����������"
Private Sub cmdAdd_Click()
    
            '������� ������������ �������� ���������� ���������� "������� ����������"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            '�������� ������ ����������
    lstTerminal.Clear
    
            '������������ ������ ���������
    gTerminal.strTerminal = "Terminal-" + Str(grdTableTerminal.Rows)
            '���������� ������ � ����� "������� ����������"
    grdTableTerminal.AddItem gTerminal.strTerminal
            '���������� ��������/���������� ����� � "������� ����������"
    gAddDelRowTableTerminal = gAddDelRowTableTerminal + 1
            '���������� ������� ��������� ��������� � "������� ����������"
    gChangesTableTerminal = True
            '���������� ����� �� ������ "Add"
    If frmTableTerminal.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            '�������� ������ �� "������� ����������"
Private Sub cmdDelete_Click()
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ����������"
Dim intColNum As Integer
    
            '������� ������������ �������� ���������� ���������� "������� ����������"
    fraColName.Enabled = False
    optTerminal.Enabled = False
    optAddrPort.Enabled = False
    optDescription.Enabled = False
    optExpander.Enabled = False
    lblTerminal.Enabled = False
    txtTerminal.Enabled = False
    lblAddrAndPort.Enabled = False
    lblAddress.Enabled = False
    txtAddress.Enabled = False
    lblPort.Enabled = False
    txtPort.Enabled = False
    lblDescription.Enabled = False
    txtDescription.Enabled = False
    lblExpander.Enabled = False
    txtExpander.Enabled = False
    lblTerminals.Enabled = False
    lstTerminal.Enabled = False
            '�������� ������ ����������
    lstTerminal.Clear
    
            '��������� (�� ���������) ����� "frmSelectRow"
    Load frmSelectRow
            '���������������� �������� "lblColName" ����� "frmSelectRow"
    frmSelectRow.lblColName.Caption = "Terminal"
    
            '������� "Terminal"
    grdTableTerminal.Col = 0
             '�������� ������ ��������
    frmSelectRow.lstSelectRow.Clear
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableTerminal.Row = intRowNum
            '���������� ������ "lstSelectRow" �������� �� "������� ����������"
        frmSelectRow.lstSelectRow.AddItem grdTableTerminal.Text
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
            '�������� ������ �� "������� ����������"
    ElseIf frmSelectRow.lstSelectRow.ListCount > 1 Then
            '����� ��������� ������
        intRowNum = frmSelectRow.lstSelectRow.ListIndex + 1
            '�������� ������
        grdTableTerminal.RemoveItem intRowNum
           '���������� ��������/���������� ����� � "������� ����������"
        gAddDelRowTableTerminal = gAddDelRowTableTerminal - 1
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableTerminal = True
    End If
            '��������� ����� "frmSelectRow"
    UnLoad frmSelectRow
            '���������� ������, ���������� ����������� ������
    Set frmSelectRow = Nothing
            '���������� ����� �� ������ "Delete"
    If frmTableTerminal.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            '���������� "������� ����������" � ����� �� ���������
Public Function SaveTableTerminal()
    Call cmdSave_Click
    SaveTableTerminal = 0
    
End Function
            
            '���������� "������� ����������" � ����� �� ���������
Private Sub cmdSave_Click()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ����������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ����������"
Dim intColNum As Integer
            '��������� ����� ������ (������) "������� ����������"
    lngRecordLen = Len(gTerminal)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + _
    Trim(Str(hsbVariant.Value)) + ".dat"
   
            '�����, ��������� �� "������� ����������" ������ ����������
            '  �����������, ' �.�. ������������ ���� ������ ������
    If gAddDelRowTableTerminal < 0 Then
            '������� "������" ������������ ����
        Kill strPathFileName
    End If
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableTerminal.Row = intRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ����������" � ����
            Select Case intColNum
                Case 0
                gTerminal.strTerminal = grdTableTerminal.Text
                Case 1
                gTerminal.strAddrPort = grdTableTerminal.Text
                Case 2
                gTerminal.strDescription = grdTableTerminal.Text
                Case 3
                gTerminal.strExpander = grdTableTerminal.Text
            End Select
        Next
            '�������� ������ "������� ����������" � ����
        Put intFileNum, intRowNum, gTerminal
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� ����������"
    gAddDelRowTableTerminal = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableTerminal = False
            '���������� ����� �� ������ "Cancel"
    If frmTableTerminal.Visible = True Then cmdCancel.SetFocus
            
End Sub
            
            '���������� "������� ����������" � ���������� �����
Private Sub cmdSaveAs_Click()
            '������ ��� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ����������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ����������"
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
            '������ "������� ����������" � ��������� ����
    Else
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� ����������"
        lngRecordLen = Len(gTerminal)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '�����, ��������� �� "������� ����������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
        If gAddDelRowTableTerminal < 0 Then
            '������� "������" ����, ���� �� ����������
            If Dir(strPathFileName) = strPathFileName Then
                Kill strPathFileName
            End If
        End If

            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableTerminal.Row = intRowNum
            '�� ���� �������� "������� ����������"
            For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
                grdTableTerminal.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ����������" � ����
                Select Case intColNum
                Case 0
                gTerminal.strTerminal = grdTableTerminal.Text
                Case 1
                gTerminal.strAddrPort = grdTableTerminal.Text
                Case 2
                gTerminal.strDescription = grdTableTerminal.Text
                Case 3
                gTerminal.strExpander = grdTableTerminal.Text
                End Select
            Next
            '�������� ������ "������� ����������" � ����
        Put intFileNum, intRowNum, gTerminal
        Next
            '������� ������������ ����
        Close intFileNum
            '���������� ��������/���������� ����� � "������� ����������"
        gAddDelRowTableTerminal = 0
            '�������� ������� ��������� ��������� � "������� ����������"
        gChangesTableTerminal = False
    End If
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '���������� ����� �� ������ "Cancel"
    If frmTableTerminal.Visible = True Then cmdCancel.SetFocus
    
End Sub


            '�������� ����� "������� ����������"
Private Sub Form_Load()
            '������ ��� ������������� ����� (� ��������� "����" � ����)
Dim strPathFileName As String
            '����� ������ "������� ����������"
Dim lngRecordLen As Long
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '������� ����� ������� "������� ����������"
Dim intColNum As Integer
            '���������� �������� � ������� ���������� �������
Dim intTerminalNum As Integer

            '���������� ������ ��������
    SetColWidth
            '���������� ��������� "������� ����������"
    lblVariant99.Caption = "V" + Str(gVarNumTerminal)
    hsbVariant.Max = gVarNumTerminal
            '��������� "������" ����� �������� "������� ����������"
    intVariantOld = hsbVariant.Value
    
            '������� ������ = 0 (��������� ��������)
    grdTableTerminal.Row = 0
    grdTableTerminal.Col = 0
    grdTableTerminal.Text = "Terminal"
            '�������� � ������ (������ 0, ������� 1)
    grdTableTerminal.Col = 1
    grdTableTerminal.Text = "Address and Port"
            '�������� � ������ (������ 0, ������� 2)
    grdTableTerminal.Col = 2
    grdTableTerminal.Text = "Description"
            '�������� � ������ (������ 0, ������� 3)
    grdTableTerminal.Col = 3
    grdTableTerminal.Text = "Expander"
            
            '���������� "������� ����������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� ����������"
    lngRecordLen = Len(gTerminal)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(hsbVariant.Value)) + ".dat"
                
            '���� ����������� - ?
    On Error GoTo ErrorTableTerminal
                '���������� ����� "������� ����������" ����� ������� ����� �� ��������� +1
    grdTableTerminal.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableTerminal.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableTerminal.Row = intRowNum
            '������ ������ "������� ����������" �� ����� � �����
        Get intFileNum, intRowNum, gTerminal
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To grdTableTerminal.Cols - 1 Step 1
            '������� ������� "������� ����������"
            grdTableTerminal.Col = intColNum
            '���������� ������� ������ "������� ����������" �� ������
            Select Case intColNum
                Case 0
                grdTableTerminal.Text = gTerminal.strTerminal
                Case 1
                grdTableTerminal.Text = gTerminal.strAddrPort
                Case 2
                grdTableTerminal.Text = gTerminal.strDescription
                Case 3
                grdTableTerminal.Text = gTerminal.strExpander
            End Select
        Next
    Next
            '������� ������������ ����
    Close intFileNum
            
            '������������ ���������� �������� � ������e ���������� �������
    intTerminalNum = grdTableTerminal.Rows
            '���� �� ���� ��������� ��������� "������� ����������"
    For intVariantNew = 1 To gVarNumTerminal Step 1
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(intVariantNew)) + ".dat"
    
            '������������ ���������� ������� � ������� ���������� �������
        If FileLen(strPathFileName) / lngRecordLen + 1 > intTerminalNum Then
            intTerminalNum = FileLen(strPathFileName) / lngRecordLen + 1
        End If
    Next
            '�������������� ����������� ������� ���������� �������
ReDim gAddrPort(gVarNumTerminal + 1, intTerminalNum) As String * 4
    
            '���� �� ���� ��������� "������� ����������"
    For intVariantNew = 0 To gVarNumTerminal Step 1
            '�������� ��������� ����� �����
        intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
        strPathFileName = App.Path
        If Right(strPathFileName, 1) <> "\" Then
            strPathFileName = strPathFileName + "\"
        End If
        strPathFileName = strPathFileName + "TableTerminal" + Trim(Str(intVariantNew)) + ".dat"
    
            '���������� "��������" �������� � ������� ������ ������� ���������� �������
        intTerminalNum = FileLen(strPathFileName) / lngRecordLen + 1
        gAddrPort(intVariantNew, 0) = intTerminalNum
    
            '������� ���� ��� ������������� �������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� �������  �������� "������� ����������"
        For intRowNum = 1 To intTerminalNum - 1 Step 1
            '������ ������ "������� ����������" �� ����� � �����
            Get intFileNum, intRowNum, gTerminal
                 gAddrPort(intVariantNew, intRowNum) = Trim(gTerminal.strAddrPort) + "0"
            '������������ ��� (�������) ������������� � ������� "Expander"
            '   "������� ����������"
            If gTerminal.strExpander = gPreprocName Then
            '�������� ����� ����� � ������� ������ ������� ���������� ��
            '   ������ ����� ����� �������������
                gAddrPort(intVariantNew, intRowNum) = _
                Left(gAddrPort(intVariantNew, intRowNum), 2) + _
                Left(gPreprocName, 1) + "0"
            End If
        Next
            '������� ����
        Close intFileNum
    Next
            
            '���������� ��������/���������� ����� � "������� ����������"
    gAddDelRowTableTerminal = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableTerminal = False
            '������������ ����� �������� � ��������� ���� "txtVariant"
    txtVariant.Text = 0
    
    Exit Sub
ErrorTableTerminal:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableTerminal Error !")
    
End Sub
            
            '��������� �������������� ������������ ��������
            '' ���������� ���� ��������� "������� ����������"
Private Sub cmdDefaultTerm_Click()
            
            '"������" ����� �������� "������� ����������" �� �������
    If hsbVariant.Value <> 0 Then
            '������� ������� ������� ������� "������� ����������"
        intVariantOld = 0
        hsbVariant.Value = 0
    Else
    
            '������� ������������ �������� ����������
            '  ���������� "������� ����������"
        fraColName.Enabled = False
        optTerminal.Enabled = False
        optAddrPort.Enabled = False
        optDescription.Enabled = False
        optExpander.Enabled = False
        lblTerminal.Enabled = False
        txtTerminal.Enabled = False
        lblAddrAndPort.Enabled = False
        lblAddress.Enabled = False
        txtAddress.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        lblDescription.Enabled = False
        txtDescription.Enabled = False
        lblExpander.Enabled = False
        txtExpander.Enabled = False
        lblTerminals.Enabled = False
        lstTerminal.Enabled = False
            '�������� ��������� ����
        txtTerminal.Text = ""
        txtAddress.Text = ""
        txtPort.Text = ""
        txtDescription.Text = ""
        txtExpander.Text = ""
            '�������� ������ ����������
        lstTerminal.Clear
    End If
            
            '�������� ����� "������� ����������"
    Form_Load
            '���������� ����� �� ������ "Correction"
    If frmTableTerminal.Visible = True Then cmdCorrection.SetFocus
            
End Sub
            
            '��������� ��������� ������ � ������������ �������� "������� ����������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '���� �� ���� ��������
    For intColNumber = 0 To grdTableTerminal.Cols - 1 Step 1
        grdTableTerminal.ColWidth(intColNumber) = 1480
        grdTableTerminal.ColAlignment(intColNumber) = 0
    Next
    
End Sub
