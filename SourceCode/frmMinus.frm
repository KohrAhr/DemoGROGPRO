VERSION 5.00
Begin VB.Form frmMinus 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minus"
   ClientHeight    =   3120
   ClientLeft      =   615
   ClientTop       =   900
   ClientWidth     =   3600
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
   ScaleHeight     =   3120
   ScaleMode       =   0  'User
   ScaleWidth      =   3600
   Visible         =   0   'False
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   240
      Picture         =   "frmMinus.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   1560
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "frmMinus.frx":030A
      ScaleHeight     =   615
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtMoneyMinus 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtMoneyReal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Tag             =   "0"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtMoney 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   186
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Tag             =   "0"
      Top             =   120
      Width           =   2415
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
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1212
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
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "   Max =      320,99        Ls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmMinus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

            '����������� ������� �����
Private Sub Form_Activate()

            '������� ����� ������� � ���������� ���� ���������� ��
            '  ����������� - ����� �� ��������� (��� ������������ ���������
            '  ��������� �����������)
    If Me.Visible = True And Me.Tag <> 0 Then Exit Sub
            
            '�������� ��������� ���� ��������� ������
    txtMoneyReal.Text = ""
            '����� ��� ���������� ����
    txtMoneyReal.BackColor = vbWhite
            '���������������� ��������� ����� �
            ' ������������� ������ ��������� ������
    If frmDataAccessIn.Visible = True Then
        txtMoney.Text = Left(frmDataAccessIn.txtMoneyDate.Text, 9)
    ElseIf frmDataAccessOut.Visible = True Then
        txtMoney.Text = Left(frmDataAccessOut.txtMoneyDate.Text, 9)
    ElseIf frmDataAccessServ.Visible = True Then
        txtMoney.Text = Left(frmDataAccessServ.txtMoneyDate.Text, 9)
    ElseIf frmDataParkingIn.Visible = True Then
        txtMoney.Text = Left(frmDataParkingIn.txtMoneyDate.Text, 9)
    ElseIf frmDataParkingOut.Visible = True Then
        txtMoney.Text = Left(frmDataParkingOut.txtMoneyDate.Text, 9)
    ElseIf frmDataParkingServ.Visible = True Then
        txtMoney.Text = Left(frmDataParkingServ.txtMoneyDate.Text, 9)
    End If
            '������������ ������ ��������� ������
    If Mid(txtMoney.Text, 2, 1) = "," Then
        txtMoney.Text = "00" + Left(txtMoney.Text, 7)
    ElseIf Mid(txtMoney.Text, 3, 1) = "," Then
        txtMoney.Text = "0" + Left(txtMoney.Text, 8)
    End If
            '������������� ������ ��������� ������
    txtMoneyReal.Text = Left(txtMoney.Text, 6)
            '������������� ������ ������������ �����
    txtMoneyMinus.Text = "000,00 Ls"
            
            '������� ��� ���������� ����
    txtMoneyReal.BackColor = vbCyan
            '���������� ����� �� ��������� ���� "��������� ������"
    txtMoneyReal.SetFocus

End Sub

            '��������� ������� "������" ������ ���� ��
            ' ��������� ���� ��������� ������
Private Sub txtMoneyReal_Click()
            '����� ��� ���������� ����
    txtMoneyReal.BackColor = vbWhite
            '������� ����������� ������ "OK _ +"
    cmdOK.Enabled = False

End Sub

            '�������� ����� ���������� ���������� �
            ' ��������� ���� ��������� ������
Private Sub txtMoneyReal_KeyPress(KeyAscii As Integer)
            
            '������ ���������� ������
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            '�������� ������ �� "������� ������� (ENTER)"
        If KeyAscii <> vbKeyReturn Then
            '������� ����������� ������ "OK _ +"
            cmdOK.Enabled = False
            '�������� ������
            frmDemo.BeepSound
            Exit Sub
        End If
    End If
    
            '�������� ������ ������
    If KeyAscii <> vbKeyReturn Or Len(Trim(txtMoneyReal.Text)) < 4 Then
            '������� ����������� ������ "OK _ +"
        cmdOK.Enabled = False
            '�������� ������
        frmDemo.BeepSound
        Exit Sub
    End If
            '�������� ������ ������
    If KeyAscii = vbKeyReturn And _
    (Mid(Trim(txtMoneyReal.Text), Len(Trim(txtMoneyReal.Text)) - 2, 1) <> _
    "," Or _
    Len(Trim(txtMoneyReal.Text)) < 4 Or _
    Len(Trim(txtMoneyReal.Text)) > 6) Then
            '�������������� ������ ��������� ������
        txtMoneyReal.Text = Left(txtMoney.Text, _
        Len(Trim(txtMoney.Text)) - 3)
            '������� ����������� ������ "OK _ +"
        cmdOK.Enabled = False
            '�������� ������
        frmDemo.BeepSound
        Exit Sub
    End If
            '������ ������ "������� ������� (ENTER)"
    If KeyAscii = vbKeyReturn Then
            '�������� ������ ������
        If CLng(Left(txtMoneyReal.Text, Len(Trim(txtMoneyReal.Text)) - 3)) > 320 Then
            '�������� ������
            frmDemo.BeepSound
            Exit Sub
        ElseIf (CInt(Left(txtMoneyReal.Text, Len(Trim(txtMoneyReal.Text)) - 3)) * _
        100 + CInt(Right(txtMoneyReal.Text, 2))) < _
        (CInt(Left(txtMoney.Text, 3)) * 100 + _
        CInt(Mid(txtMoney.Text, 5, 2))) Then
            '�������� ������
            frmDemo.BeepSound
            Exit Sub
        End If
            '������� ��� ���������� ����
        txtMoneyReal.BackColor = vbCyan
            '��������� � ���������� ������ ����� � ��������� ������
        txtMoneyMinus.Text = Trim(Str(Int(Int((CInt(Left(txtMoneyReal.Text, _
        Len(Trim(txtMoneyReal.Text)) - 3)) * 100 + CInt(Right(txtMoneyReal.Text, 2))) - _
        (CInt(Left(txtMoney.Text, 3)) * 100 + _
        CInt(Mid(txtMoney.Text, 5, 2)))) / 100)))
        txtMoneyMinus.Text = txtMoneyMinus.Text + "," + _
        Trim(Str((CInt(Left(txtMoneyReal.Text, Len(Trim(txtMoneyReal.Text)) - 3)) * _
        100 + CInt(Right(txtMoneyReal.Text, 2))) - CInt(txtMoneyMinus.Text) * 100 - _
        (CInt(Left(txtMoney.Text, 3)) * 100 + CInt(Mid(txtMoney.Text, 5, 2))))) + " Ls"
            '������������ ������ ������������ �����
        If Trim(txtMoneyMinus.Text) = "0,0 Ls" Then
            txtMoneyMinus.Text = "000,00 Ls"
        ElseIf Len(Trim(txtMoneyMinus.Text)) = 7 Then
            txtMoneyMinus.Text = "00" + txtMoneyMinus.Text
        ElseIf Len(Trim(txtMoneyMinus.Text)) = 8 Then
            txtMoneyMinus.Text = "0" + txtMoneyMinus.Text
        End If
            '������� ��������� ������ "OK _ +"
        cmdOK.Enabled = True
            '���������� ����� �� ������ "��"
        cmdOK.SetFocus
    End If
    
End Sub

            '������� � ��������� ��������� (������ "OK _ +")

Private Sub cmdOK_Click()
            '�������� ������ ����� � ��������� ������
            '  ��������� �����
    frmMinus.Tag = Left(Trim(txtMoneyMinus.Text), 6)
            '������ � ������ ����� "frmMinus"
    frmMinus.Hide
    
End Sub
            
            '������� � ��������� ��������� (������ "Cancel _ Exit")
Private Sub cmdCancel_Click()
            '����� �� ��������� ����� � ��������� ������
    frmMinus.Tag = "Exit"
            '������ � ������ ����� "frmMinus"
    frmMinus.Hide

End Sub
