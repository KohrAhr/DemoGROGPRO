VERSION 5.00
Begin VB.Form frmGetFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmFileSelection"
   ClientHeight    =   4665
   ClientLeft      =   3375
   ClientTop       =   3495
   ClientWidth     =   4815
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
   ScaleHeight     =   4665
   ScaleWidth      =   4815
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
      Left            =   3000
      TabIndex        =   11
      Top             =   4080
      Width           =   1212
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
      Left            =   600
      TabIndex        =   10
      Top             =   4080
      Width           =   1212
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   330
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   2172
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   1290
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   2172
   End
   Begin VB.ComboBox cboFileType 
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3600
      Width           =   2172
   End
   Begin VB.FileListBox filFiles 
      Height          =   1350
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2172
   End
   Begin VB.TextBox txtFileName 
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3240
      Width           =   1572
   End
   Begin VB.Label lblDirName 
      BorderStyle     =   1  'Fixed Single
      Height          =   612
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   2172
   End
   Begin VB.Label lblDitectories 
      Caption         =   "Directories"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label lblFileType 
      Caption         =   "File type"
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
      TabIndex        =   3
      Top             =   3240
      Width           =   1452
   End
   Begin VB.Label lblFileName 
      Caption         =   "File name"
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
Attribute VB_Name = "frmGetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            
            '��������� ���� �����
Private Sub cboFileType_Click()
Dim intPatternPos1 As Integer
Dim intPatternPos2 As Integer
Dim intPatternLen As Integer
Dim strPattern As String
            '����� � ������ ���������������� ���� "cboFileType" ������ �����
    intPatternPos1 = InStr(1, cboFileType.Text, "(") + 1
            '����� � ������ ���������������� ���� "cboFileType" ����� �����
    intPatternPos2 = InStr(1, cboFileType.Text, ")") - 1
            '��������� ����� �����
    intPatternLen = intPatternPos2 - intPatternPos1 + 1
            '������� ������ ����� �� ������ ���������������� ���� "cboFileType"
    strPattern = Mid(cboFileType.Text, intPatternPos1, intPatternLen)
            '�������� ����� � ������ ������
    filFiles.Pattern = strPattern

End Sub

            '��������� ������� "Cancel"
Private Sub cmdCancel_Click()
            '�������� ������� "Tag" �����
    frmGetFile.Tag = ""
            '������ � ������ �����
    frmGetFile.Hide

End Sub
            
            '��������� ������� "OK"
Private Sub cmdOK_Click()
Dim strPathAndName As String
Dim strPath As String
            '���� ���� �� ������, �� �����
    If txtFileName = "" Then
            '������ �������� ������
        frmDemo.BeepSound
        MsgBox "The file isn't selected !", vbExclamation, "Error"
        Exit Sub
    End If
            '������ ���� � ����� ������ ������������� �������� "\"
    If Right(filFiles.Path, 1) <> "\" Then
        strPath = filFiles.Path + "\"
    Else
        strPath = filFiles.Path
    End If
            '������� ��� ���������� ����� � ���� � ����
    If txtFileName.Text = filFiles.FileName Then
        strPathAndName = strPath + filFiles.FileName
    Else
        strPathAndName = strPath + txtFileName.Text
    End If
            '��������� ������ ��� ����� � �������� "Tag" �����
    frmGetFile.Tag = strPathAndName
            '������ � ������ �����
    frmGetFile.Hide

End Sub

            '��������� ��������
Private Sub dirDirectory_Change()
            '�������� ���� � ������ ������
    filFiles.Path = dirDirectory.Path
            '�������� �������� "lblDirName"
    lblDirName.Caption = dirDirectory.Path
    
End Sub

            '��������� ������ ����������
Private Sub drvDrive_Change()
            '������������� ����� �������� �� ������
    On Error GoTo DriveError
            '�������� ���� � ������ ��������� �� ����� ����������
    dirDirectory.Path = drvDrive.Drive
            '������ ���, �����
    Exit Sub
            '��������� ������
DriveError:
            '������ �������� ������
    frmDemo.BeepSound
            '��������� ������, �������� �� ���� ������������
            ' � ������������ ��������� ������ ���������
    MsgBox "The drive selection Error !", vbExclamation, "Error"
    drvDrive.Drive = dirDirectory.Path
    Exit Sub
    
End Sub
            
            '������������ �������� ����� ������� ����� "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            '��������� ����� � ������� ����� �����
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
            '��� ����� �������
    If KeyAscii = vbKeyReturn Then
            '��������� ��������� "cmdOK_Click()"
        cmdOK_Click
    End If
    
End Sub
            
            '������ ����
Private Sub filFiles_Click()
            '�������� ��� ����� � ���� "txtFileName
    txtFileName.Text = filFiles.FileName
    
End Sub
            '������ ����
Private Sub filFiles_DblClick()
            '�������� ��� ����� � ���� "txtFileName
    txtFileName.Text = filFiles.FileName
            '��������� ��������� "cmdOK_Click()"
    cmdOK_Click

End Sub

            '�������� �����
Private Sub Form_Load()
            '���������������� �������� "lblDirName"
    lblDirName.Caption = dirDirectory.Path
   
End Sub
