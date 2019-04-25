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
            
            'Обработка события "OK"
Private Sub cmdOK_Click()
            'Убрать с экрана форму
    frmSelectRow.Hide

End Sub

            'Загрузка формы
Private Sub Form_Load()
   
End Sub

            'Обработка события "Cancel"
Private Sub cmdCancel_Click()
            'Очистить своство "Tag" формы
    frmSelectRow.Tag = ""
            'Убрать с экрана форму
    frmSelectRow.Hide

End Sub
            
            'Блокирование Выгрузки формы кнопкой формы "x"
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub

            'Выбор строки
Private Sub lstSelectRow_Click()
            'Сохранить выбираемую строку в свойстве "Tag" формы
    frmSelectRow.Tag = lstSelectRow.Text
    If frmSelectRow.Visible = True Then cmdOK.SetFocus

End Sub
            
            'Перехват нажатия комбинаций клавиш "Alt"+ {"^" и "v"}
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            
            'Список пустой
    If lstSelectRow.ListCount = 0 Then
            'Издать звуковой сигнал
        frmDemo.BeepSound
            'Вывод сообщения о пустом списке
        If frmDemo.optEnglish = True Then
            MsgBox ("The List is Empty")
        Else
            MsgBox ("Saraksts ir neaizpild.")
        End If
            'Список не пустой
    Else
            'Альтернатива "щелчку" мыши на предыдущем элементе списка
        If KeyCode = 38 And Shift = 4 And lstSelectRow.ListIndex <> 0 Then
            'Выбрать  элемент спискa
            lstSelectRow.ListIndex = lstSelectRow.ListIndex - 1
            'Альтернатива "щелчку" мыши на следующем элементе списка
        ElseIf KeyCode = 40 And Shift = 4 And _
        lstSelectRow.ListIndex <> lstSelectRow.ListCount - 1 Then
            'Выбрать  элемент спискa
            lstSelectRow.ListIndex = lstSelectRow.ListIndex + 1
            'Альтернатива "щелчку" мыши на первом элементе списка
        ElseIf KeyCode = 33 And Shift = 4 And lstSelectRow.ListIndex <> 0 Then
            'Выбрать  элементы списков
            lstSelectRow.ListIndex = 0
            'Альтернатива "щелчку" мыши на последнем элементе списка
        ElseIf KeyCode = 34 And Shift = 4 And _
        lstSelectRow.ListIndex <> lstSelectRow.ListCount - 1 Then
            'Выбрать  элементы списков
            lstSelectRow.ListIndex = lstSelectRow.ListCount - 1
            'Альтернатива "щелчку" мыши на текущем элементе списка
        ElseIf (KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or _
        KeyCode = 34) And Shift = 4 Then
        
        End If
    End If

End Sub
