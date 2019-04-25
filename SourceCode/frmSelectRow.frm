VERSION 5.00
Begin VB.Form frmSelectRow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmSelectRow"
   ClientHeight    =   2310
   ClientLeft      =   3555
   ClientTop       =   4785
   ClientWidth     =   3825
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
   ScaleHeight     =   2310
   ScaleWidth      =   3825
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
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
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1212
   End
   Begin VB.ListBox lstSelectRow 
      Height          =   1320
      ItemData        =   "frmSelectRow.frx":0000
      Left            =   120
      List            =   "frmSelectRow.frx":0002
      TabIndex        =   1
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label lblColName 
      Caption         =   "Col name"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "frmSelectRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            
            '��������� ������� "OK"
Private Sub cmdOK_Click()
            '������ � ������ �����
    frmSelectRow.Hide

End Sub

            '�������� �����
Private Sub Form_Load()
   
End Sub

            '��������� ������� "Cancel"
Private Sub cmdCancel_Click()
            '�������� ������� "Tag" �����
    frmSelectRow.Tag = ""
            '������ � ������ �����
    frmSelectRow.Hide

End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '����� ������
Private Sub lstSelectRow_Click()
            '��������� ���������� ������ � �������� "Tag" �����
    frmSelectRow.Tag = lstSelectRow.Text
    If frmSelectRow.Visible = True Then cmdOK.SetFocus

End Sub
            
            '�������� ������� ���������� ������ "Alt"+ {"^" � "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            '������ ������
    If lstSelectRow.ListCount = 0 Then
            '������ �������� ������
        frmDemo.BeepSound
            '����� ��������� � ������ ������
        If frmDemo.optEnglish = True Then
            MsgBox ("The List is Empty")
        Else
            MsgBox ("Saraksts ir neaizpild.")
        End If
            '������ �� ������
    Else
            '������������ "������" ���� �� ���������� �������� ������
        If KeyCode = 38 And Shift = 4 And lstSelectRow.ListIndex <> 0 Then
            '�������  ������� �����a
            lstSelectRow.ListIndex = lstSelectRow.ListIndex - 1
            '������������ "������" ���� �� ��������� �������� ������
        ElseIf KeyCode = 40 And Shift = 4 And _
        lstSelectRow.ListIndex <> lstSelectRow.ListCount - 1 Then
            '�������  ������� �����a
            lstSelectRow.ListIndex = lstSelectRow.ListIndex + 1
            '������������ "������" ���� �� ������ �������� ������
        ElseIf KeyCode = 33 And Shift = 4 And lstSelectRow.ListIndex <> 0 Then
            '�������  �������� �������
            lstSelectRow.ListIndex = 0
            '������������ "������" ���� �� ��������� �������� ������
        ElseIf KeyCode = 34 And Shift = 4 And _
        lstSelectRow.ListIndex <> lstSelectRow.ListCount - 1 Then
            '�������  �������� �������
            lstSelectRow.ListIndex = lstSelectRow.ListCount - 1
            '������������ "������" ���� �� ������� �������� ������
        ElseIf (KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or _
        KeyCode = 34) And Shift = 4 Then
        
        End If
    End If

End Sub
