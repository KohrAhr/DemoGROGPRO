VERSION 5.00
Begin VB.Form frmLease 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lease"
   ClientHeight    =   2535
   ClientLeft      =   600
   ClientTop       =   1260
   ClientWidth     =   4335
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
   ScaleHeight     =   126.75
   ScaleMode       =   0  'User
   ScaleWidth      =   216.75
   Visible         =   0   'False
   Begin VB.CheckBox chkStock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkStock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkStock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkStock 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   12
      Top             =   600
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Image imgStockNo 
      Height          =   495
      Index           =   3
      Left            =   3480
      Picture         =   "frmLease.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockNo 
      Height          =   495
      Index           =   2
      Left            =   2640
      Picture         =   "frmLease.frx":0A4A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockNo 
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "frmLease.frx":1354
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockNo 
      Height          =   495
      Index           =   0
      Left            =   960
      Picture         =   "frmLease.frx":1BBE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblAlt4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Alt+4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblAlt3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Alt+3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblAlt2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Alt+2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblAlt1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Alt+1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblMinus 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label lblPlus 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Image imgStockYes 
      Height          =   495
      Index           =   3
      Left            =   3480
      Picture         =   "frmLease.frx":2590
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockYes 
      Height          =   495
      Index           =   2
      Left            =   2640
      Picture         =   "frmLease.frx":2DDA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockYes 
      Height          =   495
      Index           =   1
      Left            =   1800
      Picture         =   "frmLease.frx":35E4
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
   Begin VB.Image imgStockYes 
      Height          =   495
      Index           =   0
      Left            =   960
      Picture         =   "frmLease.frx":3A36
      Stretch         =   -1  'True
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmLease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '������ ���������� ���������
Dim strLease As String

            '����������� ������� �����
Private Sub Form_Activate()

            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '������������� ������ ���������� ���������
        strLease = "0000"
            '������� �������� ��� ������ ������������� ���������
        imgStockNo(0).Visible = True
        imgStockNo(1).Visible = True
        imgStockNo(2).Visible = True
        imgStockNo(3).Visible = True
            '���������� ����� ������������� ���������
        chkStock(0).Value = 0
        chkStock(1).Value = 0
        chkStock(2).Value = 0
        chkStock(3).Value = 0
            '���������������� ��������� �����
    If frmDataAccessOut.Visible = True Then
            '������������� ������ ���������� ���������
        strLease = Left(frmDataAccessOut.txtInfo, 4)
            '������� �������� ������ ��������������� ���������
            '  � ���������� ����� ��������������� ���������
        If Mid(strLease, 1, 1) = "+" Then
            imgStockNo(0).Visible = False
            chkStock(0).Value = 1
        End If
        If Mid(strLease, 2, 1) = "+" Then
            imgStockNo(1).Visible = False
            chkStock(1).Value = 1
        End If
        If Mid(strLease, 3, 1) = "+" Then
            imgStockNo(2).Visible = False
            chkStock(2).Value = 1
        End If
        If Mid(strLease, 4, 1) = "+" Then
            imgStockNo(3).Visible = False
            chkStock(3).Value = 1
        End If
    End If
            '���������� ����� �� ������ "��"
    cmdOK.SetFocus

End Sub

            '����������� ����� ���������� ���������
Private Sub chkStock_Click(Index As Integer)
            
            '����������� ����� ��������������� ���������
    If chkStock(Index).Value = 1 Then
            '������� ��������� ������ ������������� ���������
        imgStockNo(Index).Visible = False
            '�������������� ������ ���������� ���������
        If Index = 0 Then
            strLease = "+" + Mid(strLease, 2)
        ElseIf Index < 3 Then
            strLease = Left(strLease, Index) + "+" + Mid(strLease, Index + 2)
        ElseIf Index = 3 Then
            strLease = Left(strLease, 3) + "+"
        End If
            '����������� ����� ������������� ���������
    ElseIf chkStock(Index).Value = 0 Then
            '������� ������� ������ ������������� ���������
        imgStockNo(Index).Visible = True
            '�������������� ������ ���������� ���������
        If Index = 0 Then
            strLease = "0" + Mid(strLease, 2)
        ElseIf Index < 3 Then
            strLease = Left(strLease, Index) + "0" + Mid(strLease, Index + 2)
        ElseIf Index = 3 Then
            strLease = Left(strLease, 3) + "0"
        End If
    End If
            
            '���������� ����� �� ������ "��"
    cmdOK.SetFocus

End Sub
            
            '�������� ������� ������ "Alt"+ {"1", "2", "3" � "4"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������ ������� "1"
    If KeyCode = 49 And Shift = 0 Then
            '������� ��������� ������ ������������� ���������
        imgStockNo(0).Visible = False
            '���������� ����� ��������������� ���������
        chkStock(0).Value = 1
            '�������������� ������ ���������� ���������
        strLease = "+" + Mid(strLease, 2)
            '������ ������� "2"
    ElseIf KeyCode = 50 And Shift = 0 Then
            '������� ��������� ������ ������������� ���������
        imgStockNo(1).Visible = False
            '���������� ����� ��������������� ���������
        chkStock(1).Value = 1
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 1) + "+" + Mid(strLease, 3)
            '������ ������� "3"
    ElseIf KeyCode = 51 And Shift = 0 Then
            '������� ��������� ������ ������������� ���������
        imgStockNo(2).Visible = False
            '���������� ����� ��������������� ���������
        chkStock(2).Value = 1
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 2) + "+" + Mid(strLease, 4)
            '������ ������� "4"
    ElseIf KeyCode = 52 And Shift = 0 Then
            '������� ��������� ������ ������������� ���������
        imgStockNo(3).Visible = False
            '���������� ����� ��������������� ���������
        chkStock(3).Value = 1
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 3) + "+"
            '������ ���������� ������ "Alt+1"
    ElseIf KeyCode = 49 And Shift = 4 Then
            '������� ������� ������ ������������� ���������
        imgStockNo(0).Visible = True
            '���������� ����� ������������� ���������
        chkStock(0).Value = 0
            '�������������� ������ ���������� ���������
        strLease = "0" + Mid(strLease, 2)
            '������ ���������� ������ "Alt+2"
    ElseIf KeyCode = 50 And Shift = 4 Then
            '������� ������� ������ ������������� ���������
        imgStockNo(1).Visible = True
            '���������� ����� ������������� ���������
        chkStock(1).Value = 0
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 1) + "0" + Mid(strLease, 3)
            '������ ���������� ������ "Alt+3"
    ElseIf KeyCode = 51 And Shift = 4 Then
            '������� ������� ������ ������������� ���������
        imgStockNo(2).Visible = True
            '���������� ����� ������������� ���������
        chkStock(2).Value = 0
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 2) + "0" + Mid(strLease, 4)
            '������ ���������� ������ "Alt+4"
    ElseIf KeyCode = 52 And Shift = 4 Then
            '������� ������� ������ ������������� ���������
        imgStockNo(3).Visible = True
            '���������� ����� ������������� ���������
        chkStock(3).Value = 0
            '�������������� ������ ���������� ���������
        strLease = Left(strLease, 3) + "0"
    End If
            
            '���������� ����� �� ������ "��"
    cmdOK.SetFocus

End Sub

            '������� � ��������� ��������� (������ "OK _ +")

Private Sub cmdOK_Click()
            '�������� ������ ���������� ��������� ��������� �����
    frmLease.Tag = strLease
            '������ � ������ ����� "frmLease"
    frmLease.Hide
    
End Sub
            
            '������� � ��������� ��������� (������ "Cancel _ Exit")
Private Sub cmdCancel_Click()
            '����� �� ��������� ������ ���������� ���������
    frmLease.Tag = "Exit"
            '������ � ������ ����� "frmLease"
    frmLease.Hide

End Sub
