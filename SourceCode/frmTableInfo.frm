VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTableInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "table_information"
   ClientHeight    =   7125
   ClientLeft      =   1185
   ClientTop       =   1080
   ClientWidth     =   9510
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
   ScaleHeight     =   7125
   ScaleWidth      =   9510
   Begin VB.TextBox txtCategory 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   33
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtSite 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   32
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtBrigade 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   31
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtDepartment 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   30
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox txtTNumber 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   29
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtSurName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   28
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   27
      Top             =   1560
      Width           =   5055
   End
   Begin VB.CommandButton cmdFormTableInfoFromTablePerson 
      Cancel          =   -1  'True
      Caption         =   "F&Orm 'TableInfo' From 'TablePerson'"
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
      Left            =   7200
      TabIndex        =   22
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtDeletion 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7320
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtRegistration 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7320
      TabIndex        =   19
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtRemark 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   1212
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
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
      Left            =   5520
      TabIndex        =   12
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete..."
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
      Left            =   6720
      TabIndex        =   11
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   1560
      TabIndex        =   10
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Sa&VeAs..."
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
      Left            =   2760
      TabIndex        =   9
      Top             =   6480
      Width           =   1092
   End
   Begin VB.ListBox lstPersonID 
      Enabled         =   0   'False
      Height          =   1740
      ItemData        =   "frmTableInfo.frx":0000
      Left            =   120
      List            =   "frmTableInfo.frx":0002
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCorrection 
      Caption         =   "Co&Rrection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   1215
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
      Height          =   4095
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton optCategory 
         Caption         =   "Category"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   1455
      End
      Begin VB.OptionButton optSite 
         Caption         =   "Site"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton optBrigade 
         Caption         =   "Brigade"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton optDepartment 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optTNumber 
         Caption         =   "T_Number"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optSurName 
         Caption         =   "SurName"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optRemark 
         Caption         =   "Remark"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1455
      End
      Begin VB.OptionButton optPersonID 
         Caption         =   "PersonID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optCardID 
         Caption         =   "CardID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   186
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.TextBox txtPersonID 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCardID 
      Enabled         =   0   'False
      Height          =   288
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdDefaultPers 
      Caption         =   "D&Efault from HDD"
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
      Left            =   8280
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find..."
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
      Left            =   4320
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdTableInfo 
      Height          =   1815
      Left            =   2160
      TabIndex        =   14
      Top             =   4560
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   9
      Cols            =   13
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
   Begin VB.Label lblDeletion 
      Alignment       =   2  'Center
      Caption         =   "Deletion"
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
      Left            =   7320
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblRegistration 
      Alignment       =   2  'Center
      Caption         =   "Registration"
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
      Left            =   7320
      TabIndex        =   17
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblPersOrTerm 
      Alignment       =   2  'Center
      Caption         =   "PersonID "
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
      Left            =   360
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   6120
      Y2              =   6360
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2040
      X2              =   120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   4440
      Y2              =   6360
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   9360
      X2              =   2040
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   9360
      X2              =   9360
      Y1              =   4440
      Y2              =   120
   End
End
Attribute VB_Name = "frmTableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             '��������� ������
Dim strPassword As String
            '������� ����� �������������� ������ "������� ����������"
Dim intRowNumCorr As Integer
            '������� ����� ��������������� ������� "������� ����������"
Dim intColNumCorr As Integer
            '������� ����� �����
Dim intFileNum As Integer
            '������ "������� ����������"
Dim gInfo As ExtendInfo
            '������ ����������� ���������
Dim strMessage As String

            '������� � ��������� ���������
Private Sub cmdCancel_Click()
            '���������� "������ + ������" � ���� ���������
    Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
    Dim strResponse As String
            
            '���� �� ����������� ��������� � "������� ����������"
    If gChangesTableInfo = True Then
            '������ �������� ������
        frmDemo.BeepSound
            '���� �������� � �������� ���������� "������� ����������" - �� �����
        intButtonsAndIcons = vbYesNo + vbQuestion
        strResponse = MsgBox("Save changes ?", intButtonsAndIcons, "Cancel")
            '������ ������ "��"
        If strResponse = vbYes Then
            '���������� "T������ ������" � ����� �� ���������
            cmdSave_Click
        End If
    End If
    
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ��������� ����
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            '�������� ������ ����
    lstPersonID.Clear
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableInfo = False
            '������� ��������� ������� �����
    frmTableInfo.Visible = False
            '������� ��������� ����� "frmDemo"
    frmDemo.Enabled = True
            '������� ������� ����� "frmDemo"
    frmDemo.Show
    
End Sub
            
            '���������
Private Sub cmdCorrection_Click()
            
            ' "������� ������" �� �������� ��������������� �����
    If grdTableInfo.Rows = 1 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������������� ���������
        If frmDemo.optEnglish = True Then
            MsgBox ("The TableInfo is Empty")
        Else
            MsgBox ("Personas tabula ir neaizpild.")
        End If
    
    Else
            '������� ���������� �������� ���������� ����������
            '  "������� ����������"
        fraColName.Enabled = True
        optPersonID.Value = True
        txtPersonID.Enabled = True
        lstPersonID.Enabled = True
        txtCardID.Enabled = True
        txtTNumber.Enabled = True
        txtName.Enabled = True
        txtSurName.Enabled = True
        txtDepartment.Enabled = True
        txtBrigade.Enabled = True
        txtSite.Enabled = True
        txtCategory.Enabled = True
        txtRemark.Enabled = True
            '�������� ��������� ����
        txtPersonID.Text = ""
        txtCardID.Text = ""
        txtRegistration.Text = ""
        txtDeletion.Text = ""
        txtTNumber.Text = ""
        txtName.Text = ""
        txtSurName.Text = ""
        txtDepartment.Text = ""
        txtBrigade.Text = ""
        txtSite.Text = ""
        txtCategory.Text = ""
        txtRemark.Text = ""
            '�������� ������ ����
        lstPersonID.Clear
    
            '������� "Person or Terminal"
        grdTableInfo.Col = 0
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNumCorr = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNumCorr
            '���������� ������ "lstPersonID" �������� �� "������� ����������"
            lstPersonID.AddItem grdTableInfo.Text
        Next
            '�������  ������� ������
        lstPersonID.ListIndex = 0
            '����� �������������� ������ - (1)
        intRowNumCorr = 1
        grdTableInfo.Row = intRowNumCorr
            '�������� �����
        optPersonID_Click
    
    End If
    
End Sub

            '��������� ������������ "������� ����������"
            '  �� "������� ������"
Private Sub cmdFormTableInfoFromTablePerson_Click()
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ��������� ����
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            '�������� ������ ����
    lstPersonID.Clear
            
            '�����
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            
            '�������� �� "������� ����������" ���� ������������ �����
    grdTableInfo.Rows = 2
    grdTableInfo.Row = 1
            '���� �� ���� ��������������� ������� "������� ������"
    For intRowNum = 1 To gTablePerson.Rows - 1 Step 1
            '������� ������ "������� ������"
        gTablePerson.Row = intRowNum
            '���� ����� ������������� ���� � "������� ������" ����� 16-� ������
            '   - ���������� (����� �� ��������� � �� ���������� ����������)
        gTablePerson.Col = 1
        If Len(Trim(gTablePerson.Text)) = 16 Then
            '���������� ������ � ����� "������� ����������"
            If grdTableInfo.Row = grdTableInfo.Rows - 1 Then
                grdTableInfo.AddItem ""
                grdTableInfo.Row = grdTableInfo.Rows - 1
            End If
            '��������� ������ "PersonID" � "������� ����������"
            gTablePerson.Col = 0
            grdTableInfo.Col = 0
            
            ' ���� � ���� "Info" ��� �������� �����
            If Left(Trim(gTablePerson.Text), 1) <> gVisitor Then
                grdTableInfo.Text = Trim(gTablePerson.Text)
                If Len(Trim(grdTableInfo.Text)) = 16 Then
                    If Mid(Trim(grdTableInfo.Text), 16, 1) = "+" Or _
                    Mid(Trim(grdTableInfo.Text), 16, 1) = "-" Then
                        grdTableInfo.Text = Trim(Left(Trim(grdTableInfo.Text), 15))
                    End If
                End If
            '��������� ������ "CardID" � "������� ����������"
                gTablePerson.Col = 1
                grdTableInfo.Col = 1
                grdTableInfo.Text = Trim(gTablePerson.Text)
            '������� ������� "������� ����������" = 2 (����� � ���� �����������)
                grdTableInfo.Col = 2
                grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
                "  ||  " + Trim(gProtocol.strProtocDate)
            '���������� ��������/���������� ����� � "������� ����������"
                grdTableInfo.Tag = grdTableInfo.Tag + 1
            '���������� ������� ��������� ��������� � "������� ����������"
                gChangesTableInfo = True
            'B ���� "Info" ���� ������� �����
            Else
                grdTableInfo.RemoveItem grdTableInfo.Row
                grdTableInfo.Row = grdTableInfo.Rows - 1
            End If
        
        End If
    Next
            
            '���������� ������� "������� ������ ����" ���
            '  ���������� ������� "Save" ��� "Save As"
    grdTableInfo.Tag = -1
            '���������� ����� �� ������ "Correction"
    If frmTableInfo.Visible = True Then cmdCorrection.SetFocus


End Sub
            
            '��������� �������������� ������������ ��������
            ' ���������� "������� ����������"
Private Sub cmdDefaultPers_Click()
            
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ��������� ����
    txtPersonID.Text = ""
    txtCardID.Text = ""
    txtRegistration.Text = ""
    txtDeletion.Text = ""
    txtTNumber.Text = ""
    txtName.Text = ""
    txtSurName.Text = ""
    txtDepartment.Text = ""
    txtBrigade.Text = ""
    txtSite.Text = ""
    txtCategory.Text = ""
    txtRemark.Text = ""
            '�������� ������ ����
    lstPersonID.Clear
            
            '�������� ����� "������� ����������"
    Form_Load
            '���������� ����� �� ������ "Correction"
    If frmTableInfo.Visible = True Then cmdCorrection.SetFocus

End Sub

            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� �������������� ������ "������� ����������"
Private Sub grdTableInfo_Click()
            '��������� "��������"
    If lstPersonID.Enabled = True Then
            '����� �������������� ������ "������� ����������"
        intRowNumCorr = grdTableInfo.RowSel
        grdTableInfo.Row = intRowNumCorr
            '����� ���������� �������� ������
        lstPersonID.ListIndex = intRowNumCorr - 1
            '����� ��������������� ������� "������� ����������"
        intColNumCorr = grdTableInfo.ColSel
        grdTableInfo.Col = intColNumCorr
            '����� �������������� ������ "������� ����������"
        lstPersonID_MouseDown Button:=vbLeftButton, Shift:=0, X:=lstPersonID.Left, Y:=lstPersonID.Top
            '����� ��������������� ������� "������� ����������"
        Select Case intColNumCorr
            '���������� ����� �� ��������� ���� ��� ���������
            Case 0
            optPersonID.Value = True
            txtPersonID.SetFocus
            Case 1
            optCardID.Value = True
            txtCardID.SetFocus
            Case 2
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 3
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 4
            optTNumber.Value = True
            txtTNumber.SetFocus
            Case 5
            optName.Value = True
            txtName.SetFocus
            Case 6
            optSurName.Value = True
            txtSurName.SetFocus
            Case 7
            optDepartment.Value = True
            txtDepartment.SetFocus
            Case 8
            optBrigade.Value = True
            txtBrigade.SetFocus
            Case 9
            optSite.Value = True
            txtSite.SetFocus
            Case 10
            optCategory.Value = True
            txtCategory.SetFocus
            Case 11
            optRemark.Value = True
            txtRemark.SetFocus
            Case 12
            optRemark.Value = True
            txtRemark.SetFocus
        End Select
    End If
        
End Sub

            '����� �������������� ������ "������� ����������"
Private Sub lstPersonID_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
            '������ ����� ������ "����"
    If Button = vbLeftButton Then
            '����� �������������� ������ "������� ����������"
        intRowNumCorr = lstPersonID.ListIndex + 1
        grdTableInfo.Row = intRowNumCorr
            
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� ��������� ��� �����������
        grdTableInfo.Col = 0
        txtPersonID.Text = grdTableInfo.Text
        grdTableInfo.Col = 1
        txtCardID.Text = grdTableInfo.Text
        grdTableInfo.Col = 2
        txtRegistration.Text = grdTableInfo.Text
        grdTableInfo.Col = 3
        txtDeletion.Text = grdTableInfo.Text
        grdTableInfo.Col = 4
        txtTNumber.Text = grdTableInfo.Text
        grdTableInfo.Col = 5
        txtName.Text = grdTableInfo.Text
        grdTableInfo.Col = 6
        txtSurName.Text = grdTableInfo.Text
        grdTableInfo.Col = 7
        txtDepartment.Text = grdTableInfo.Text
        grdTableInfo.Col = 8
        txtBrigade.Text = grdTableInfo.Text
        grdTableInfo.Col = 9
        txtSite.Text = grdTableInfo.Text
        grdTableInfo.Col = 10
        txtCategory.Text = grdTableInfo.Text
        grdTableInfo.Col = 11
        txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
        grdTableInfo.Col = intColNumCorr
    End If

End Sub

            '������� ����� - "PersonID"
Private Sub optPersonID_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 0
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtPersonID.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "CardID"
Private Sub optCardID_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 1
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtCardID.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtCardID.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "TNumber"
Private Sub optTNumber_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 4
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtTNumber.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtTNumber.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Name"
Private Sub optName_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 5
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtName.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtName.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "SurName"
Private Sub optSurName_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 6
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtSurName.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtSurName.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Department"
Private Sub optDepartment_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 7
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtDepartment.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtDepartment.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Brigade"
Private Sub optBrigade_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 8
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtBrigade.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtBrigade.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Site"
Private Sub optSite_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 9
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtSite.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtSite.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Category"
Private Sub optCategory_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 10
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtCategory.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtCategory.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtRemark.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub
            
            '������� ����� - "Remark"
Private Sub optRemark_Click()
            '����� ��������������� ������� "������� ����������"
    intColNumCorr = 11
    grdTableInfo.Col = intColNumCorr
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtRemark.Enabled = True
            '���������� ����� �� ��������� ���� ��� ���������
    txtRemark.SetFocus
            '������� (��)���������� ��������� �������� ������. ���������� "������� ����������"
    txtPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
            '����������� ����� "������� ����������" � ��������� ����
            '  ��� �����������
    grdTableInfo.Col = 0
    txtPersonID.Text = grdTableInfo.Text
    grdTableInfo.Col = 1
    txtCardID.Text = grdTableInfo.Text
    grdTableInfo.Col = 2
    txtRegistration.Text = grdTableInfo.Text
    grdTableInfo.Col = 3
    txtDeletion.Text = grdTableInfo.Text
    grdTableInfo.Col = 4
    txtTNumber.Text = grdTableInfo.Text
    grdTableInfo.Col = 5
    txtName.Text = grdTableInfo.Text
    grdTableInfo.Col = 6
    txtSurName.Text = grdTableInfo.Text
    grdTableInfo.Col = 7
    txtDepartment.Text = grdTableInfo.Text
    grdTableInfo.Col = 8
    txtBrigade.Text = grdTableInfo.Text
    grdTableInfo.Col = 9
    txtSite.Text = grdTableInfo.Text
    grdTableInfo.Col = 10
    txtCategory.Text = grdTableInfo.Text
    grdTableInfo.Col = 11
    txtRemark.Text = grdTableInfo.Text
            '������������ ����� ��������������� ������� "������� ����������"
    grdTableInfo.Col = intColNumCorr

End Sub

            '��������� ����� � ������� ��������������� ����� "PersonID"
Private Sub txtPersonID_KeyPress(KeyAscii As Integer)
            '��� �������
    If KeyAscii = vbKeyReturn Then
            '��� � ���������� ���������
        If Len(Trim(txtPersonID.Text)) < 17 Then
            '��������� ����� "Person or Terminal" � "������� ����������"
        grdTableInfo.Text = Trim(txtPersonID.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optCardID"
        optCardID.Value = True
            Exit Sub
            '��� � ������������ ���������
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "CardID"
Private Sub txtCardID_KeyPress(KeyAscii As Integer)
            '��� ������
    If KeyAscii = vbKeyReturn Then
            '������� �� ������ �������������� ������
        On Error GoTo CardIDError
            '������������ ��� � ���������� ���������
        If Len(Trim(txtCardID.Text)) = 16 Then
            '��������� ������ "CardID" � "������� ����������"
            grdTableInfo.Text = Trim(txtCardID.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
            gChangesTableInfo = True
            '�������� ����� "optTNumber"
            optTNumber.Value = True
            Exit Sub
            '������������ ��� � ������������ ���������
CardIDError:
            frmDemo.BeepSound
        Else
            frmDemo.BeepSound
        End If
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "TNumber"
Private Sub txtTNumber_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtTNumber.Text)) > 8 Then
            txtTNumber.Text = Left(Trim(txtTNumber.Text), 8)
        End If
            '��������� ������ "TNumber" � "������� ����������"
        grdTableInfo.Text = Trim(txtTNumber.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optName"
        optName.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Name"
Private Sub txtName_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtName.Text)) > 32 Then
            txtName.Text = Left(Trim(txtName.Text), 32)
        End If
            '��������� ������ "Name" � "������� ����������"
        grdTableInfo.Text = Trim(txtName.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optSurName"
        optSurName.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "SurName"
Private Sub txtSurName_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtSurName.Text)) > 32 Then
            txtSurName.Text = Left(Trim(txtSurName.Text), 32)
        End If
            '��������� ������ "SurName" � "������� ����������"
        grdTableInfo.Text = Trim(txtSurName.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optDeparttment"
        optDepartment.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Department"
Private Sub txtDepartment_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtDepartment.Text)) > 32 Then
            txtDepartment.Text = Left(Trim(txtDepartment.Text), 32)
        End If
            '��������� ������ "Department" � "������� ����������"
        grdTableInfo.Text = Trim(txtDepartment.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optRemark"
        optBrigade.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Brigade"
Private Sub txtBrigade_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtBrigade.Text)) > 16 Then
            txtBrigade.Text = Left(Trim(txtBrigade.Text), 16)
        End If
            '��������� ������ "Brigade" � "������� ����������"
        grdTableInfo.Text = Trim(txtBrigade.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optRemark"
        optSite.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Site"
Private Sub txtSite_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtSite.Text)) > 16 Then
            txtSite.Text = Left(Trim(txtSite.Text), 16)
        End If
            '��������� ������ "Site" � "������� ����������"
        grdTableInfo.Text = Trim(txtSite.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optRemark"
        optCategory.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Category"
Private Sub txtCategory_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtCategory.Text)) > 16 Then
            txtCategory.Text = Left(Trim(txtCategory.Text), 16)
        End If
            '��������� ������ "Category" � "������� ����������"
        grdTableInfo.Text = Trim(txtCategory.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optRemark"
        optRemark.Value = True
        Exit Sub
    End If

End Sub
            
            '��������� ����� � ������� ��������������� "Remark"
Private Sub txtRemark_KeyPress(KeyAscii As Integer)
            
            '���������� �������
    If KeyAscii = vbKeyReturn Then
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(txtRemark.Text)) > 64 Then
            txtRemark.Text = Left(Trim(txtRemark.Text), 64)
        End If
            '��������� ������ "Remark" � "������� ����������"
        grdTableInfo.Text = Trim(txtRemark.Text)
            '���������� �������  ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
            '�������� ����� "optPersonID"
        optPersonID.Value = True
        Exit Sub
    End If

End Sub
            
            '���������� ������ � "������� ����������"
Private Sub cmdAdd_Click()
            '��� � ������������ ��� � "������� ����������"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ������ ����
    lstPersonID.Clear
    
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ ��� �������
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Add ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            '�������� �� ������������ ������������ ���
    strCardID = InputBox("CardID: 16 Characters !!!", "Add ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strCardID)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            '��� ��� ������������ ��� �� �������
    If strPersonID = "" Or strCardID = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            '��� � ������������ ��� �������
    Else
            '����������� ������������� ���� � "������� ����������"
        Call Reg(strCardID, strPersonID, "", "", "", "", "", "", "", "")
    End If
            '���������� ����� �� ������ "Add"
    If frmTableInfo.Visible = True Then cmdAdd.SetFocus
    
End Sub
            
            '����� ������ � "������� ����������"
Private Sub cmdFind_Click()
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '��� � ������������ ��� � "������� ����������"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ������ ����
    lstPersonID.Clear
    
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ - ��� ������
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Find ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            '�������� �� ������������ - ��� �����
    strCardID = InputBox("CardID: 16 Characters !!!", "Find ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strCardID)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            '��� ��� ������������ ��� �� �������
    If strPersonID = "" Or strCardID = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            '��� � ������������ ��� �������
    Else
        '������� ������� "������� ����������" = 1 (������������ ���)
        grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = strCardID Then
            '������� ������� "������� ����������" = 0 (���)
                grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '���������� ��������� ���� "������� ����������"
                    grdTableInfo.Col = 0
                    txtPersonID.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 1
                    txtCardID.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 2
                    txtRegistration.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 3
                    txtDeletion.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 4
                    txtTNumber.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 5
                    txtName.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 6
                    txtSurName.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 7
                    txtDepartment.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 8
                    txtBrigade.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 9
                    txtSite.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 10
                    txtCategory.Text = Trim(grdTableInfo.Text)
                    grdTableInfo.Col = 11
                    txtRemark.Text = Trim(grdTableInfo.Text)
            '��������� ����� �� �����
                    Exit For
                End If
            End If
        Next
            '����� ��� ������������� ���� ��� � "������� ����������"
        If intRowNum = grdTableInfo.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
            frmDemo.BeepSound
            MsgBox ("Unexistent PersonID Or CardID")
        End If
            '���������� ����� �� ������ "Find"
        If frmTableInfo.Visible = True Then cmdFind.SetFocus
    End If
            
    
End Sub

            '����������� ������������� ���� � "������� ����������"
            '   ��� ��������: 0 - ����������� ��������� �������;
            '                 1 - ��������� ��������� �������;
            '                 2 - � ����������� ��������.
Public Function Reg(ByVal vntCardID As Variant, ByVal strPersonID As String, _
ByVal strTNumber As String, ByVal strName As String, ByVal strSurName As String, _
ByVal strDepartment As String, ByVal strBrigade As String, _
ByVal strSite As String, ByVal strCategory As String, ByVal strRemark As String)
            '����� ������� ������ � "������� ����������"
Dim intRowNum As Integer
            
            '� ����������� ��������
    Reg = 2
            
            '���� ��� ����� - ��� ���������� � "������� ����������"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Exit Function
    End If

            '�������� �������� �����/������ �� ���� �����
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
    
        '������� ������� "������� ����������" = 1 (������������ ���)
    grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ����������"
        If Trim(grdTableInfo.Text) = vntCardID Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
            grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '��������� ����� �� �����
                Exit For
            End If
        End If
    Next
        

            '��������� ������������ ��� & ��� ��� ���� � "������� ����������"
    If intRowNum < grdTableInfo.Rows Then
            '���������
        Reg = 1
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Correction 'TableInfo'")
        Else
            MsgBox ("Korekcija 'TableInfo'")
        End If
            
            '��������� ����������, ���������� � "������� ����������"
        Call frmTableInfo.Corr(strTNumber, strName, strSurName, _
        strDepartment, strBrigade, strSite, strCategory, strRemark)
            
            '���������� ������������� ���� & ����� ��� � "������� ����������"
    Else
            '�����������
        Reg = 0
            '���������� ������ � ����� "������� ����������"
        grdTableInfo.AddItem strPersonID
        grdTableInfo.Row = grdTableInfo.Rows - 1
            '��������� ������ "Person or Terminal" � "������� ����������"
        grdTableInfo.Col = 0
        grdTableInfo.Text = Trim(strPersonID)
            '��������� ������ "CardID" � "������� ����������"
        grdTableInfo.Col = 1
        grdTableInfo.Text = Trim(vntCardID)
            '��������� ������ "TNumber" � "������� ����������"
        grdTableInfo.Col = 4
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strTNumber)) > 8 Then
            strTNumber = Left(Trim(strTNumber), 8)
        End If
        grdTableInfo.Text = Trim(strTNumber)
            '��������� ������ "Name" � "������� ����������"
        grdTableInfo.Col = 5
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strName)) > 32 Then
            strName = Left(Trim(strName), 32)
        End If
        grdTableInfo.Text = Trim(strName)
            '��������� ������ "SurName" � "������� ����������"
        grdTableInfo.Col = 6
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strSurName)) > 32 Then
            strSurName = Left(Trim(strSurName), 32)
        End If
        grdTableInfo.Text = Trim(strSurName)
            '��������� ������ "Department" � "������� ����������"
        grdTableInfo.Col = 7
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strDepartment)) > 32 Then
            strDepartment = Left(Trim(strDepartment), 32)
        End If
        grdTableInfo.Text = Trim(strDepartment)
            '��������� ������ "Brigade" � "������� ����������"
        grdTableInfo.Col = 8
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strBrigade)) > 16 Then
            strBrigade = Left(Trim(strBrigade), 16)
        End If
        grdTableInfo.Text = Trim(strBrigade)
            '��������� ������ "Site" � "������� ����������"
        grdTableInfo.Col = 9
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strSite)) > 16 Then
            strSite = Left(Trim(strSite), 16)
        End If
        grdTableInfo.Text = Trim(strSite)
            '��������� ������ "Category" � "������� ����������"
        grdTableInfo.Col = 10
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strCategory)) > 16 Then
            strCategory = Left(Trim(strCategory), 16)
        End If
        grdTableInfo.Text = Trim(strCategory)
            '��������� ������ "Remark" � "������� ����������"
        grdTableInfo.Col = 11
            '�������� ������ �������� �� ���������� ����
        If Len(Trim(strRemark)) > 64 Then
            strRemark = Left(Trim(strRemark), 64)
        End If
        grdTableInfo.Text = Trim(strRemark)
            
            '�����
        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '������� ������� "������� ����������" = 2 (����� � ���� �����������)
        grdTableInfo.Col = 2
        grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
        "  ||  " + Trim(gProtocol.strProtocDate)
            
            '������ �������� ���������
        strMessage = "RegInfo " + strPersonID + Chr(7) + _
        Trim(vntCardID) + Chr(7) + grdTableInfo.Text + Chr(7) + " " + Chr(7) + _
        strTNumber + Chr(7) + strName + Chr(7) + strSurName + Chr(7) + _
        strDepartment + Chr(7) + strBrigade + Chr(7) + strSite + Chr(7) + _
        strCategory + Chr(7) + strRemark + Chr(7) + " "
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
        Call frmDemo.SendMessage(strMessage)
            
            '���������� ��������/���������� ����� � "������� ����������"
        grdTableInfo.Tag = grdTableInfo.Tag + 1
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
    
            '��������������� ������ �������
        gProtocol.strProtocName = strPersonID
        gProtocol.strProtocPersonCode = vntCardID
        gProtocol.strProtocStatus = "?? - TableInfo"
            '����������
        gProtocol.strProtocReserve = "Registration"
            '�������� ������ � ���� "������� ���������"
        frmDemo.WriteProtocol
    End If
            
            '������ ����������� ��������� ������� ��� �����������
            '  �������� - �������� ������ ������������ "������� ����������"
    If grdTableInfo.Rows > 32000 Then
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("'TableInfo' > 32000 rows")
        Else
            MsgBox ("'TableInfo' > 32000 rin.")
        End If
    End If
    
End Function
            
            '����� & ������������� ���� & ����� � "������� ����������"
            '   ��� ��������: 0 - ����� �������� �������;
            '                 1 - ����� ����������.
Public Function Find(ByVal vntCardID As Variant, ByVal strPersonID As String, _
ByRef strTNumber As String, ByRef strName As String, ByRef strSurName As String, _
ByRef strDepartment As String, ByRef strBrigade As String, _
ByRef strSite As String, ByRef strCategory As String, ByRef strRemark As String)
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            
            
            '���� ��� ����� - ��� ���������� � "������� ����������"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Find = 1
        Exit Function
    End If
            
            '�������� �������� �����/������ �� ���� �����
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
            
            '����� ����������
    Find = 1
        
        '����� �� �������������� ���� & �����
    If vntCardID <> "" And strPersonID <> "" Then
        '������� ������� "������� ����������" = 1 (������������ ���)
        grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = vntCardID Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
                grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '����� �������� �������
                    Find = 0
            '������� ������� "������� ����������" = 4
                    grdTableInfo.Col = 4
                    strTNumber = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 5
                    grdTableInfo.Col = 5
                    strName = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 6
                    grdTableInfo.Col = 6
                    strSurName = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 7
                    grdTableInfo.Col = 7
                    strDepartment = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 8
                    grdTableInfo.Col = 8
                    strBrigade = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 9
                    grdTableInfo.Col = 9
                    strSite = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 10
                    grdTableInfo.Col = 10
                    strCategory = Trim(grdTableInfo.Text)
            '������� ������� "������� ����������" = 11
                    grdTableInfo.Col = 11
                    strRemark = Trim(grdTableInfo.Text)
            '��������� ����� �� �����
                    Exit For
                End If
        '������� ������� "������� ����������" = 1 (������������ ���)
                grdTableInfo.Col = 1
            End If
        Next
    End If
    
            '����������� ������������� ���� ��� � "������� ����������"
    If Find = 1 Or intRowNum = grdTableInfo.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            '����� & ������������� ���� & ����� � "������� ����������"
            '   ��� ��������: 0 - ����� �������� �������;
            '                 1 - ����� ����������.
Public Function Corr(ByVal strTNumber As String, ByVal strName As String, _
ByVal strSurName As String, ByVal strDepartment As String, _
ByVal strBrigade As String, ByVal strSite As String, ByVal strCategory As String, _
ByVal strRemark As String)
            '������������ ������������� - ������������ ��� ������
Dim strPersonID As String
            '������������ ��� � ������� - ����� ��������
Dim strCardID As String
            '����� � ���� �����������
Dim strTimeDateReg As String
            '����� � ���� ����������
Dim strTimeDateDel As String
            '������
Dim strReserve As String
            
            '���������� ���������� �� ������ "PersonID" � "������� ����������"
    grdTableInfo.Col = 0
    strPersonID = grdTableInfo.Text
            '���������� ���������� �� ������ "CardID" � "������� ����������"
    grdTableInfo.Col = 1
    strCardID = grdTableInfo.Text
            '���������� ���������� �� ������ "TimeDateReg" � "������� ����������"
    grdTableInfo.Col = 2
    strTimeDateReg = grdTableInfo.Text
            '���������� ���������� �� ������ "TimeDateDel" � "������� ����������"
    grdTableInfo.Col = 3
    strTimeDateDel = grdTableInfo.Text
            '���������� ���������� �� ������ "Reserve" � "������� ����������"
    grdTableInfo.Col = 12
    strReserve = grdTableInfo.Text
            
            '��������� ������ "TNumber" � "������� ����������"
    grdTableInfo.Col = 4
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strTNumber)) > 8 Then
        strTNumber = Left(Trim(strTNumber), 8)
    End If
    grdTableInfo.Text = Trim(strTNumber)
            '��������� ������ "Name" � "������� ����������"
    grdTableInfo.Col = 5
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strName)) > 32 Then
        strName = Left(Trim(strName), 32)
    End If
    grdTableInfo.Text = Trim(strName)
            '��������� ������ "SurName" � "������� ����������"
    grdTableInfo.Col = 6
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strSurName)) > 32 Then
        strSurName = Left(Trim(strSurName), 32)
    End If
    grdTableInfo.Text = Trim(strSurName)
            '��������� ������ "Department" � "������� ����������"
    grdTableInfo.Col = 7
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strDepartment)) > 32 Then
        strDepartment = Left(Trim(strDepartment), 32)
    End If
    grdTableInfo.Text = Trim(strDepartment)
            '��������� ������ "Brigade" � "������� ����������"
    grdTableInfo.Col = 8
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strBrigade)) > 16 Then
        strBrigade = Left(Trim(strBrigade), 16)
    End If
    grdTableInfo.Text = Trim(strBrigade)
            '��������� ������ "Site" � "������� ����������"
    grdTableInfo.Col = 9
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strSite)) > 16 Then
        strSite = Left(Trim(strSite), 16)
    End If
    grdTableInfo.Text = Trim(strSite)
            '��������� ������ "Category" � "������� ����������"
    grdTableInfo.Col = 10
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strCategory)) > 16 Then
        strCategory = Left(Trim(strCategory), 16)
    End If
    grdTableInfo.Text = Trim(strCategory)
            '��������� ������ "Remark" � "������� ����������"
    grdTableInfo.Col = 11
            '�������� ������ �������� �� ���������� ����
    If Len(Trim(strRemark)) > 64 Then
        strRemark = Left(Trim(strRemark), 64)
    End If
    grdTableInfo.Text = Trim(strRemark)
            
            '����� � ���� ��������� � "������� ����������"
    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
        
    If Len(Trim(strRemark)) = 64 Then
        grdTableInfo.Text = grdTableInfo.Text + "  " + _
        Trim(gProtocol.strProtocTime) + " || " + _
        Trim(gProtocol.strProtocDate)
    Else
        grdTableInfo.Text = grdTableInfo.Text + _
        Left("                                                                  ", _
        66 - Len(Trim(strRemark))) + _
        Trim(gProtocol.strProtocTime) + " || " + _
        Trim(gProtocol.strProtocDate)
    End If
            
            '������ �������� ���������
    strMessage = "CorInfo " + strPersonID + Chr(7) + strCardID + Chr(7) + _
    strTimeDateReg + Chr(7) + strTimeDateDel + Chr(7) + _
    Trim(strTNumber) + Chr(7) + Trim(strName) + Chr(7) + _
    Trim(strSurName) + Chr(7) + Trim(strDepartment) + Chr(7) + _
    Trim(strBrigade) + Chr(7) + Trim(strSite) + Chr(7) + _
    Trim(strCategory) + Chr(7) + Trim(strRemark) + Chr(7) + _
    Trim(strReserve) + Chr(7)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
    Call frmDemo.SendMessage(strMessage)
            
            '���������� ������� ��������� ��������� � "������� ����������"
    gChangesTableInfo = True
            
            '��������������� ������ �������
    grdTableInfo.Col = 0
    gProtocol.strProtocName = grdTableInfo.Text
    grdTableInfo.Col = 1
    gProtocol.strProtocPersonCode = grdTableInfo.Text
    gProtocol.strProtocStatus = "?? - TableInfo"
            '����������
    gProtocol.strProtocReserve = "Correction"
            '�������� ������ � ���� "������� ���������"
    frmDemo.WriteProtocol
    
End Function
            '�������� (����������) ������ �� "������� ����������"
            '   ��� ��������: 0 - �������� ��������� �������;
            '                 1 - � �������� ��������.
Public Function Del(ByVal vntCardID As Variant, ByVal strPersonID As String)
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '����� � ���� �����������
Dim strTimeDateReg As String
            '����� � ���� ����������
Dim strTimeDateDel As String
            
            '���� ��� ����� - ��� ���������� � "������� ����������"
    If Left(Trim(strPersonID), 1) = gVisitor Then
        Del = 1
        Exit Function
    End If
            
            '�������� �������� �����/������ �� ���� �����
    If Len(Trim(strPersonID)) = 16 Then
        If Mid(Trim(strPersonID), 16, 1) = "+" Or _
        Mid(Trim(strPersonID), 16, 1) = "-" Then
            strPersonID = Trim(Left(Trim(strPersonID), 15))
        End If
    End If
    
            '�������� ����������
    Del = 1
        
        '����� �� �������������� ���� & �����
    If vntCardID <> "" And strPersonID <> "" Then
        '������� ������� "������� ����������" = 1 (������������ ���)
        grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = vntCardID Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
                grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '����� �������� �������
                    Del = 0
            '��������� ������� ������ �� "������� ����������" �
            '  ��������������� ������ �������
                    gProtocol.strProtocName = strPersonID
                    gProtocol.strProtocPersonCode = vntCardID
                    gProtocol.strProtocStatus = "?? - TableInfo"
            '�����
                    gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
                    gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                    gProtocol.strProtocReserve = "Logical Deletion"
            '�������� ������ � ���� "������� ���������"
                    frmDemo.WriteProtocol
            
            '������� ������� "������� ����������" = 3 (����� � ���� ����������)
                    grdTableInfo.Col = 3
                    grdTableInfo.Text = Trim(gProtocol.strProtocTime) + _
                    "  ||  " + Trim(gProtocol.strProtocDate)
                    
            '���������� ���������� �� ������ "TimeDateDel" � "������� ����������"
                    strTimeDateDel = grdTableInfo.Text
            '���������� ���������� �� ������ "TimeDateReg" � "������� ����������"
                    grdTableInfo.Col = 2
                    strTimeDateReg = grdTableInfo.Text

            '������ �������� ���������
                    strMessage = "DelInfo " + Trim(strPersonID) + Chr(7) + _
                    vntCardID + Chr(7) + strTimeDateReg + Chr(7) + _
                    strTimeDateDel + Chr(7)
            '��������� �������� ���������
            '  ���������� ������� "MSMQ"
                    Call frmDemo.SendMessage(strMessage)
            
            '���������� ������� ��������� ��������� � "������� ����������"
                    gChangesTableInfo = True
            '��������� ����� �� �����
                    Exit For
                End If
            End If
        Next
    End If
    
            '����������� ������������� ���� ��� � "������� ����������"
    If Del = 1 Or intRowNum = grdTableInfo.Rows Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            '�������� (����������) ������ �� "������� ����������"
            '   ��� ��������: 0 - �������� ��������� �������;
            '                 1 - � �������� ��������.
Public Function RealDel(ByVal vntCardID As Variant, ByVal strPersonID As String)
            '������� ����� ��������������� ������ "������� ����������"
Dim intRowNum As Integer
            '���������� "������ + ������" � ���� ���������
Dim intButtonsAndIcons  As Integer
            '������ ������ ������������ �� ����� ���� ���������
Dim strResponse As String
            
            '��������� ������ �� ����� ���� ������� - ����� �� ���������
    If grdTableInfo.Rows = 2 Then Exit Function
    
            '�������� ����������
    RealDel = 1
        
        '����� �� �������������� ���� & �����
    If vntCardID <> "" And strPersonID <> "" Then
        '������� ������� "������� ����������" = 1 (������������ ���)
        grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNum
            '���������� ������������ ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = vntCardID Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
                grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
                If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '����� �������� �������
                    RealDel = 0
            '������ ����������� ��������� ������� ��� ����������� ��������
                    frmDemo.BeepSound
            '���� �������� � ��������� �������� ��������
            '   ������������� ���� - �� �����
                    intButtonsAndIcons = vbYesNo + vbQuestion
                    If frmDemo.optEnglish = True Then
                        strResponse = MsgBox("Deletion Information ?", intButtonsAndIcons, "Cancel")
                    Else
                        strResponse = MsgBox("Izslegt Info ?", intButtonsAndIcons, "Cancel")
                    End If
            '������ ������ "��"
                    If strResponse = vbYes Then
            '���������� �������� ������ �� "������� ����������"
                        grdTableInfo.RemoveItem grdTableInfo.Row
            '���������� ��������/���������� ����� � "������� ����������"
                        grdTableInfo.Tag = grdTableInfo.Tag - 1
            
            '��������������� ������ �������
                        gProtocol.strProtocName = strPersonID
                        gProtocol.strProtocPersonCode = vntCardID
                        gProtocol.strProtocStatus = "?? - TableInfo"
            '�����
                        gProtocol.strProtocTime = Format(Now, "hh:mm:ss")
            '����
                        gProtocol.strProtocDate = Format(Now, "dd/mm/yyyy")
            '����������
                        gProtocol.strProtocReserve = "Real Deletion"
            '�������� ������ � ���� "������� ���������"
                        frmDemo.WriteProtocol
            
            '���������� ������� ��������� ��������� � "������� ����������"
                        gChangesTableInfo = True
                    End If
            '��������� ����� �� �����
                    Exit For
                End If
            End If
        Next
    End If
    
            '����������� ������������� ���� ��� � "������� ����������"
    If RealDel = 1 Then
            '������ ����������� ��������� ������� ��� ����������� ��������
        frmDemo.BeepSound
        If frmDemo.optEnglish = True Then
            MsgBox ("Unexistent string 'TableInfo'")
        Else
            MsgBox ("Neeksist. rinda 'TableInfo'")
        End If
    End If
    
End Function
            
            '�������� ������ �� "������� ����������"
Private Sub cmdDelete_Click()
            '��� � ������������ ��� � "������� ����������"
Dim strPersonID As String
Dim strCardID As String

    strPersonID = ""
    strCardID = ""
    
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ������ ����
    lstPersonID.Clear
    
            '������ �������� ������
    frmDemo.BeepSound
            '�������� �� ������������ ��� �������
    strPersonID = InputBox("PersonID: 1 -- 16 Characters !!!", "Delete ...")
    If Len(Trim(strPersonID)) > 16 Then strPersonID = Left(Trim(strPersonID), 16)
    frmDemo.BeepSound
            '�������� �� ������������ ������������ ���
    strCardID = InputBox("CardID: 16 Characters !!!", "Delete ...")
    If Len(Trim(strCardID)) > 16 Then strCardID = _
    Left(Trim(strCardID), 16)
            '����� ������������� ���� ������ 16-� ��������
    If Len(Trim(strCardID)) < 16 Then
            '�������� ����������� ���������� ���������� �����
        strCardID = Left("0000000000000000", _
        16 - Len(Trim(strCardID))) + Trim(strCardID)
    End If
    
            '��� ��� ������������ ��� �� �������
    If strPersonID = "" Or strCardID = "" Then
            '������ �������� ������
       frmDemo.BeepSound
       MsgBox " The PersonID Or CardID isn't selected"
            
            '��� � ������������ ��� �������
    Else
            '�������� �������� ������������� ���� �� "������� ����������"
        Call RealDel(strCardID, strPersonID)
    End If
            '���������� ����� �� ������ "Add"
    If frmTableInfo.Visible = True Then cmdDelete.SetFocus
    
End Sub
            
            '���������� "������� ����������" � ����� �� ���������
Public Function SaveTableInfo()
    Call cmdSave_Click
    SaveTableInfo = 0
    
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
            
            '���� ������ ���� = "�������� ����", �� �����
    If Me.MousePointer = vbHourglass Then Exit Sub
            
            '�������� ����������� ������ ����  �� "�������� ����"
    Me.MousePointer = vbHourglass
            
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ������ ����
    lstPersonID.Clear
            
            '��������� ����� ������ (������) "������� ����������"
    lngRecordLen = Len(gInfo)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
    
    
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableInfo.dat"
    
            '�����, ��������� �� "������� ����������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
    If grdTableInfo.Tag < 0 Then
        On Error Resume Next
            '������� "������" ������������ ����
        Kill strPathFileName
        On Error GoTo 0
    End If
    
            '������� ������������ ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            '������� ������� "������� ����������"
            grdTableInfo.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ����������"
            '  � ����
            Select Case intColNum
                Case 0
                gInfo.strPersonID = grdTableInfo.Text
                Case 1
                gInfo.strCardID = grdTableInfo.Text
                Case 2
                gInfo.strTimeDateReg = grdTableInfo.Text
                Case 3
                gInfo.strTimeDateDel = grdTableInfo.Text
                Case 4
                gInfo.strTNumber = grdTableInfo.Text
                Case 5
                gInfo.strName = grdTableInfo.Text
                Case 6
                gInfo.strSurName = grdTableInfo.Text
                Case 7
                gInfo.strDepartment = grdTableInfo.Text
                Case 8
                gInfo.strBrigade = grdTableInfo.Text
                Case 9
                gInfo.strSite = grdTableInfo.Text
                Case 10
                gInfo.strCategory = grdTableInfo.Text
                Case 11
                gInfo.strRemark = grdTableInfo.Text
                Case 12
                gInfo.strReserve = " "
            End Select
        Next
            '�������� ������ "������� ����������" � ����
        Put intFileNum, intRowNum, gInfo
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� ����������"
    grdTableInfo.Tag = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableInfo = False
            '������������ ����������� ������ ����
    Me.MousePointer = 0
            '���������� ����� �� ������ "Cancel"
    If frmTableInfo.Visible = True Then cmdCancel.SetFocus
            
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

            '���� ������ ���� = "�������� ����", �� �����
    If Me.MousePointer = vbHourglass Then Exit Sub
            
            '������� ������������ �������� ���������� ����������
            '  "������� ����������"
    fraColName.Enabled = False
    txtPersonID.Enabled = False
    lstPersonID.Enabled = False
    txtCardID.Enabled = False
    txtTNumber.Enabled = False
    txtName.Enabled = False
    txtSurName.Enabled = False
    txtDepartment.Enabled = False
    txtBrigade.Enabled = False
    txtSite.Enabled = False
    txtCategory.Enabled = False
    txtRemark.Enabled = False
            '�������� ������ ����
    lstPersonID.Clear
            
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
            '�������� ����������� ������ ����  �� "�������� ����"
        Me.MousePointer = vbHourglass
            
            '������ ��� ����� (� ��������� "����" � ����)
        strPathFileName = frmGetFile.Tag
            '��������� ����� ������ (������) "������� ����������"
        lngRecordLen = Len(gInfo)
            '�������� ��������� ����� �����
        intFileNum = FreeFile
    
            '�����, ��������� �� "������� ����������" ������ ���������� �����������,
            ' �.�. ������������ ���� ������ ������
        If grdTableInfo.Tag < 0 Then
            '������� "������" ����, ���� �� ����������
            If Dir(strPathFileName) = strPathFileName Then
                On Error Resume Next
            '������� "������" ������������ ����
                Kill strPathFileName
                On Error GoTo 0
            End If
        End If

            '������� ��������� ���� ��� ������������� ������� ���
            '   ������� ���, ���� �� �� ����������
        Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
        For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
            grdTableInfo.Row = intRowNum
            '�� ���� �������� "������� ����������"
            For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            '������� ������� "������� ����������"
                grdTableInfo.Col = intColNum
            '���������� ������ ��� ������ ������� ������ "������� ����������"
            '  � ����
                Select Case intColNum
                    Case 0
                    gInfo.strPersonID = grdTableInfo.Text
                    Case 1
                    gInfo.strCardID = grdTableInfo.Text
                    Case 2
                    gInfo.strTimeDateReg = grdTableInfo.Text
                    Case 3
                    gInfo.strTimeDateDel = grdTableInfo.Text
                    Case 4
                    gInfo.strTNumber = grdTableInfo.Text
                    Case 5
                    gInfo.strName = grdTableInfo.Text
                    Case 6
                    gInfo.strSurName = grdTableInfo.Text
                    Case 7
                    gInfo.strDepartment = grdTableInfo.Text
                    Case 8
                    gInfo.strBrigade = grdTableInfo.Text
                    Case 9
                    gInfo.strSite = grdTableInfo.Text
                    Case 10
                    gInfo.strCategory = grdTableInfo.Text
                    Case 11
                    gInfo.strRemark = grdTableInfo.Text
                    Case 12
                    gInfo.strReserve = " "
                End Select
            Next
            '�������� ������ "������� ����������" � ����
            Put intFileNum, intRowNum, gInfo
        Next
            '������� ��������� ����
        Close intFileNum
             '���������� ��������/���������� ����� � "������� ����������"
        grdTableInfo.Tag = 0
               '�������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = False
    End If
    
            '��������� ����� "frmGetFile"
    UnLoad frmGetFile
            '���������� ������, ���������� ����������� ������
    Set frmGetFile = Nothing
            '������������ ����������� ������ ����
    Me.MousePointer = 0
            '���������� ����� �� ������ "Cancel"
    If frmTableInfo.Visible = True Then cmdCancel.SetFocus
    
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

            '���������� ������ ��������
    SetColWidth
            '������� ������ = 0 (��������� ��������)
    grdTableInfo.Row = 0
    grdTableInfo.Col = 0
    grdTableInfo.Text = "PersonID"
            '�������� � ������ (������ 0, ������� 1)
    grdTableInfo.Col = 1
    grdTableInfo.Text = "CardID"
            '�������� � ������ (������ 0, ������� 2)
    grdTableInfo.Col = 2
    grdTableInfo.Text = "Time & Date Registration"
            '�������� � ������ (������ 0, ������� 3)
    grdTableInfo.Col = 3
    grdTableInfo.Text = "Time & Date Deletion"
            '�������� � ������ (������ 0, ������� 4)
    grdTableInfo.Col = 4
    grdTableInfo.Text = "TNumber"
            '�������� � ������ (������ 0, ������� 5)
    grdTableInfo.Col = 5
    grdTableInfo.Text = "Name"
            '�������� � ������ (������ 0, ������� 6)
    grdTableInfo.Col = 6
    grdTableInfo.Text = "SurName"
            '�������� � ������ (������ 0, ������� 7)
    grdTableInfo.Col = 7
    grdTableInfo.Text = "Department"
            '�������� � ������ (������ 0, ������� 8)
    grdTableInfo.Col = 8
    grdTableInfo.Text = "Brigade"
            '�������� � ������ (������ 0, ������� 9)
    grdTableInfo.Col = 9
    grdTableInfo.Text = "Site"
            '�������� � ������ (������ 0, ������� 10)
    grdTableInfo.Col = 10
    grdTableInfo.Text = "Category"
            '�������� � ������ (������ 0, ������� 11)
    grdTableInfo.Col = 11
    grdTableInfo.Text = "Remark"
            '�������� � ������ (������ 0, ������� 12)
    grdTableInfo.Col = 12
    grdTableInfo.Text = "Reserve"
            
            
            '���������� "������� ����������" �� ����� �� ���������
            
            '��������� ����� ������ (������) "������� ����������"
    lngRecordLen = Len(gInfo)
            '�������� ��������� ����� �����
    intFileNum = FreeFile
            '���������� �������������� "����" � �������� ����������� ���������
    strPathFileName = App.Path
    If Right(strPathFileName, 1) <> "\" Then
        strPathFileName = strPathFileName + "\"
    End If
    strPathFileName = strPathFileName + "TableInfo.dat"
                
            '���� ����������� - ?
    On Error GoTo ErrorTableInfo
                '���������� ����� "������� ����������"
    grdTableInfo.Rows = FileLen(strPathFileName) / lngRecordLen + 1
    
            '������� ������������ ���� ��� ������������� �������
    Open strPathFileName For Random As intFileNum Len = lngRecordLen
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '������ ������ "������� ����������" �� ����� � �����
        Get intFileNum, intRowNum, gInfo
            '�� ���� �������� "������� ����������"
        For intColNum = 0 To grdTableInfo.Cols - 1 Step 1
            '������� ������� "������� ����������"
            grdTableInfo.Col = intColNum
            '���������� ������� ������ "������� ����������" �� ������
            Select Case intColNum
                Case 0
                grdTableInfo.Text = gInfo.strPersonID
                Case 1
                grdTableInfo.Text = gInfo.strCardID
                Case 2
                grdTableInfo.Text = gInfo.strTimeDateReg
                Case 3
                grdTableInfo.Text = gInfo.strTimeDateDel
                Case 4
                grdTableInfo.Text = gInfo.strTNumber
                Case 5
                grdTableInfo.Text = gInfo.strName
                Case 6
                grdTableInfo.Text = gInfo.strSurName
                Case 7
                grdTableInfo.Text = gInfo.strDepartment
                Case 8
                grdTableInfo.Text = gInfo.strBrigade
                Case 9
                grdTableInfo.Text = gInfo.strSite
                Case 10
                grdTableInfo.Text = gInfo.strCategory
                Case 11
                grdTableInfo.Text = gInfo.strRemark
                Case 12
                grdTableInfo.Text = gInfo.strReserve
            End Select
        Next
    Next
            '������� ������������ ����
    Close intFileNum
            '���������� ��������/���������� ����� � "������� ����������"
    grdTableInfo.Tag = 0
            '�������� ������� ��������� ��������� � "������� ����������"
    gChangesTableInfo = False
    
    Exit Sub
    
ErrorTableInfo:
            '������ �������� ������
    frmDemo.BeepSound
    MsgBox ("TableInfo Error !")
            '���������� ��������/���������� ����� � "������� ����������"
    grdTableInfo.Tag = 0
    
End Sub
            
            '���������� ������ � �������� ������ � "������� ����������"
            '  �� ��������� MSMQ, ���������� �� ����
Public Function MSMQReg(ByVal strMessage As String)

            '����� ������� ������ � "������� ����������"
Dim intRowNum As Integer
            '����� ������������� ������ � ������ ���������
Dim intNumber As Integer
            '������ ������ "������� ����������"
Dim strInfo As String
            
            '������������ ������������� - ������������ ��� ������
Dim strPersonID As String
            '������������ ��� � ������� - ����� ��������
Dim strCardID As String
            '����� � ���� �����������
Dim strTimeDateReg As String
            '����� � ���� ����������
Dim strTimeDateDel As String
            '��������� �����
Dim strTNumber As String
            '���
Dim strName As String
            '�������
Dim strSurName As String
            '�����/���
Dim strDepartment As String
            '�������/�������������
Dim strBrigade As String
            '�������/������� �����
Dim strSite As String
            '���������/���������/���������
Dim strCategory As String
            '�������������� ����������
Dim strRemark As String
            '������
Dim strReserve As String
            
            '����� ������������� ������ � ������ ���������
    intNumber = 1
            '������ � ������ ��������� ������� "07H" - ����������� �������
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            '���������� ������ "PersonID" ��� "������� ����������"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            '���������� ������ "CardID" ��� "������� ����������"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            '���������� ������ "TimeDateReg" ��� "������� ����������"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            '���������� ������ "TimeDateDel" ��� "������� ����������"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            '���������� ������ "TNumber" ��� "������� ����������"
            strTNumber = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 6 Then
            '���������� ������ "Name" ��� "������� ����������"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 7 Then
            '���������� ������ "SurName" ��� "������� ����������"
            strSurName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 8 Then
            '���������� ������ "Department" ��� "������� ����������"
            strDepartment = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 9 Then
            '���������� ������ "Brigade" ��� "������� ����������"
            strBrigade = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 10 Then
            '���������� ������ "Site" ��� "������� ����������"
            strSite = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 11 Then
            '���������� ������ "Category" ��� "������� ����������"
            strCategory = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 12 Then
            '���������� ������ "Remark" ��� "������� ����������"
            strRemark = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 13 Then
            '���������� ������ "Reserve" ��� "������� ����������"
            strReserve = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            '�������������� �����
            Exit Do
        End If
    Loop
        
        '������� ������� "������� ����������" = 1 (������������ ���)
    grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ����������"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
            grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '��������� ����� �� �����
                Exit For
            End If
        End If
    Next

            '���������� ������������ ��� & ��� ��� ���� � "������� ����������" -
            '   ���������
    If intRowNum < grdTableInfo.Rows Then
            '��������� ������ "PersonID" � "������� ����������"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            '��������� ������ "CardID" � "������� ����������"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            '��������� ������ "TimeDateReg" � "������� ����������"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            '��������� ������ "TimeDateDel" � "������� ����������"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            '��������� ������ "TNumber" � "������� ����������"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            '��������� ������ "Name" � "������� ����������"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            '��������� ������ "SurName" � "������� ����������"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            '��������� ������ "Department" � "������� ����������"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            '��������� ������ "Brigade" � "������� ����������"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            '��������� ������ "Site" � "������� ����������"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            '��������� ������ "Category" � "������� ����������"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            '��������� ������ "Remark" � "������� ����������"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            '��������� ������ "Reserve" � "������� ����������"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
            
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
    
            '����������� ������������� ���� & ����� ��� � "������� ����������" -
            '   ����������� (����������)
    Else
            '���������� ������ � ����� "������� ����������"
        grdTableInfo.AddItem strInfo
        grdTableInfo.Row = grdTableInfo.Rows - 1
            '��������� ������ "PersonID" � "������� ����������"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            '��������� ������ "CardID" � "������� ����������"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            '��������� ������ "TimeDateReg" � "������� ����������"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            '��������� ������ "TimeDateDel" � "������� ����������"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            '��������� ������ "TNumber" � "������� ����������"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            '��������� ������ "Name" � "������� ����������"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            '��������� ������ "SurName" � "������� ����������"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            '��������� ������ "Department" � "������� ����������"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            '��������� ������ "Brigade" � "������� ����������"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            '��������� ������ "Site" � "������� ����������"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            '��������� ������ "Category" � "������� ����������"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            '��������� ������ "Remark" � "������� ����������"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            '��������� ������ "Reserve" � "������� ����������"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
    
            '���������� ��������/���������� ����� � "������� ����������"
        grdTableInfo.Tag = grdTableInfo.Tag + 1
    
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
    
    End If
    
End Function
            
            '�������� (����������) ������ � �������� ������������ �����
            '  �� "������� ����������" �� ��������� MSMQ, ���������� �� ����
Public Function MSMQDel(ByVal strMessage As String)

            '����� ������� ������ � "������� ����������"
Dim intRowNum As Integer
            '����� ������������� ������ � ������ ���������
Dim intNumber As Integer
            
            '������������ ������������� - ������������ ��� ������
Dim strPersonID As String
            '������������ ��� � ������� - ����� ��������
Dim strCardID As String
            '����� � ���� �����������
Dim strTimeDateReg As String
            '����� � ���� ����������
Dim strTimeDateDel As String
            
            '����� ������������� ������ � ������ ���������
    intNumber = 1
            '������ � ������ ��������� ������� "07H" - ����������� �������
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            '���������� ������ "PersonID" ��� "������� ����������"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            '���������� ������ "CardID" ��� "������� ����������"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            '���������� ������ "TimeDateReg" ��� "������� ����������"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            '���������� ������ "TimeDateDel" ��� "������� ����������"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            '�������������� �����
            Exit Do
        End If
    Loop
        
        '������� ������� "������� ����������" = 1 (������������ ���)
    grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ����������"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
            grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '��������� ����� �� �����
                Exit For
            End If
        End If
    Next

            '���������� ������������ ��� & ��� ��� ���� � "������� ����������" -
            '   ���������
    If intRowNum < grdTableInfo.Rows Then
            '��������� ������ "TimeDateDel" � "������� ����������"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
    
    End If

End Function
            
            '��������� �������� ����� ������ "������� ������"
            '  �� ��������� MSMQ, ���������� �� ����
Public Function MSMQCor(ByVal strMessage As String)

            '����� ������� ������ � "������� ����������"
Dim intRowNum As Integer
            '����� ������������� ������ � ������ ���������
Dim intNumber As Integer
            
            '������������ ������������� - ������������ ��� ������
Dim strPersonID As String
            '������������ ��� � ������� - ����� ��������
Dim strCardID As String
            '����� � ���� �����������
Dim strTimeDateReg As String
            '����� � ���� ����������
Dim strTimeDateDel As String
            '��������� �����
Dim strTNumber As String
            '���
Dim strName As String
            '�������
Dim strSurName As String
            '�����/���
Dim strDepartment As String
            '�������/�������������
Dim strBrigade As String
            '�������/������� �����
Dim strSite As String
            '���������/���������/���������
Dim strCategory As String
            '�������������� ����������
Dim strRemark As String
            '������
Dim strReserve As String
            
            '����� ������������� ������ � ������ ���������
    intNumber = 1
            '������ � ������ ��������� ������� "07H" - ����������� �������
    Do While InStr(1, strMessage, Chr(7)) <> 0
        If intNumber = 1 Then
            '���������� ������ "PersonID" ��� "������� ����������"
            strPersonID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 2 Then
            '���������� ������ "CardID" ��� "������� ����������"
            strCardID = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 3 Then
            '���������� ������ "TimeDateReg" ��� "������� ����������"
            strTimeDateReg = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 4 Then
            '���������� ������ "TimeDateDel" ��� "������� ����������"
            strTimeDateDel = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 5 Then
            '���������� ������ "TNumber" ��� "������� ����������"
            strTNumber = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 6 Then
            '���������� ������ "Name" ��� "������� ����������"
            strName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 7 Then
            '���������� ������ "SurName" ��� "������� ����������"
            strSurName = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 8 Then
            '���������� ������ "Department" ��� "������� ����������"
            strDepartment = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 9 Then
            '���������� ������ "Brigade" ��� "������� ����������"
            strBrigade = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 10 Then
            '���������� ������ "Site" ��� "������� ����������"
            strSite = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 11 Then
            '���������� ������ "Category" ��� "������� ����������"
            strCategory = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 12 Then
            '���������� ������ "Remark" ��� "������� ����������"
            strRemark = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            intNumber = intNumber + 1
        ElseIf intNumber = 13 Then
            '���������� ������ "Reserve" ��� "������� ����������"
            strReserve = Left(strMessage, InStr(1, strMessage, Chr(7)) - 1)
            strMessage = Mid(strMessage, InStr(1, strMessage, Chr(7)) + 1)
            '�������������� �����
            Exit Do
        End If
    Loop
        
        '������� ������� "������� ����������" = 1 (������������ ���)
    grdTableInfo.Col = 1
            '���� �� ���� ��������������� ������� "������� ����������"
    For intRowNum = 1 To grdTableInfo.Rows - 1 Step 1
            '������� ������ "������� ����������"
        grdTableInfo.Row = intRowNum
            '��������� ������������ ��� ���� � "������� ����������"
        If Trim(grdTableInfo.Text) = Trim(strCardID) Then
            '������� ������� "������� ����������" = 0 (������� ��� ��������)
            grdTableInfo.Col = 0
            '���������� ��� ���� � "������� ����������"
            If Trim(grdTableInfo.Text) = Trim(strPersonID) Then
            '��������� ����� �� �����
                Exit For
            End If
        End If
    Next

            '���������� ������������ ��� & ��� ��� ���� � "������� ����������" -
            '   ���������
    If intRowNum < grdTableInfo.Rows Then
            '��������� ������ "PersonID" � "������� ����������"
        grdTableInfo.Col = 0
        grdTableInfo.Text = strPersonID
            '��������� ������ "CardID" � "������� ����������"
        grdTableInfo.Col = 1
        grdTableInfo.Text = strCardID
            '��������� ������ "TimeDateReg" � "������� ����������"
        grdTableInfo.Col = 2
        grdTableInfo.Text = strTimeDateReg
            '��������� ������ "TimeDateDel" � "������� ����������"
        grdTableInfo.Col = 3
        grdTableInfo.Text = strTimeDateDel
            '��������� ������ "TNumber" � "������� ����������"
        grdTableInfo.Col = 4
        grdTableInfo.Text = strTNumber
            '��������� ������ "Name" � "������� ����������"
        grdTableInfo.Col = 5
        grdTableInfo.Text = strName
            '��������� ������ "SurName" � "������� ����������"
        grdTableInfo.Col = 6
        grdTableInfo.Text = strSurName
            '��������� ������ "Department" � "������� ����������"
        grdTableInfo.Col = 7
        grdTableInfo.Text = strDepartment
            '��������� ������ "Brigade" � "������� ����������"
        grdTableInfo.Col = 8
        grdTableInfo.Text = strBrigade
            '��������� ������ "Site" � "������� ����������"
        grdTableInfo.Col = 9
        grdTableInfo.Text = strSite
            '��������� ������ "Category" � "������� ����������"
        grdTableInfo.Col = 10
        grdTableInfo.Text = strCategory
            '��������� ������ "Remark" � "������� ����������"
        grdTableInfo.Col = 11
        grdTableInfo.Text = strRemark
            '��������� ������ "Reserve" � "������� ����������"
        grdTableInfo.Col = 12
        grdTableInfo.Text = strReserve
            
            '���������� ������� ��������� ��������� � "������� ����������"
        gChangesTableInfo = True
    
    End If
    
End Function

            '��������� ��������� ������ � ������������ �������� "������� ����������"
Public Sub SetColWidth()
            '���������� ���������� - ������� ����� �������
Dim intColNumber As Integer
            '���� �� ���� ��������
    For intColNumber = 0 To grdTableInfo.Cols - 1 Step 1
        grdTableInfo.ColWidth(intColNumber) = 2500
        grdTableInfo.ColAlignment(intColNumber) = 0
    Next
            '���������� ������� 2-�� � 3-�� �������� (����� � ���� ���/����)
    intColNumber = 2
    grdTableInfo.ColWidth(intColNumber) = 1850
    intColNumber = 3
    grdTableInfo.ColWidth(intColNumber) = 1850
            '���������� ������� 4-�� ������a (��������� �����)
    intColNumber = 4
    grdTableInfo.ColWidth(intColNumber) = 1300
            '���������� ������� 5-�� � 6-�� �������� (���, ������� � ���)
    intColNumber = 5
    grdTableInfo.ColWidth(intColNumber) = 4300
    intColNumber = 6
    grdTableInfo.ColWidth(intColNumber) = 4300
    intColNumber = 7
    grdTableInfo.ColWidth(intColNumber) = 4300
            '���������� ������� 11-�� ������� (����������)
    intColNumber = 11
    grdTableInfo.ColWidth(intColNumber) = 4950

End Sub
