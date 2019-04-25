VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITablePerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
            
            'The decrement of the accessable attribute
Public Sub AccessMinus()

End Sub

            'The increment of the accessable attribute
Public Sub AccessPlus()

End Sub

            'Adds a new row to a FlexGrid control at run time
Public Sub AddItem(ByVal Item As String)

End Sub

            'Removes a row from a FlexGrid control at run time
Public Sub RemoveItem(ByVal Index As Long)

End Sub

            'Returns attribute the accessable your object
Public Property Get Access() As String

End Property

            'Returns the active cell in a FlexGrid
Public Property Get Col() As Long

End Property

            'Sets the active cell in a FlexGrid
Public Property Let Col(ByVal vNewValue As Long)

End Property

            'Sets the alignment of data in a column
Public Property Let ColAlignment(ByVal vNewValue As Long)

End Property

            'Determines the total number of columns in a FlexGrid
Public Property Get Cols() As Long

End Property

            'Determines the total number of columns in a FlexGrid
Public Property Let Cols(ByVal vNewValue As Long)

End Property

            'Determines the starting or ending column for a range of cells
Public Property Get ColSel() As Long

End Property

            'Determines the starting or ending column for a range of cells
Public Property Let ColSel(ByVal vNewValue As Long)

End Property

            'Determines the width of the specified column in Twips
Public Property Let ColWidth(ByVal vNewValue As Long)

End Property

            'Returns the active cell in a FlexGrid
Public Property Get Row() As Long

End Property

            'Sets the active cell in a FlexGrid
Public Property Let Row(ByVal vNewValue As Long)

End Property

            'Determines the total number of rows in a FlexGrid
Public Property Get Rows() As Long

End Property

            'Determines the total number of rows in a FlexGrid
Public Property Let Rows(ByVal vNewValue As Long)

End Property

            'Determines the starting or ending row for a range of cells
Public Property Get RowSel() As Long

End Property

            'Determines the starting or ending row for a range of cells
Public Property Let RowSel(ByVal vNewValue As Long)

End Property

            'Returns any extra data needed for your program
Public Property Get Tag() As String

End Property

            'Stores any extra data needed for your program
Public Property Let Tag(ByVal vNewValue As String)

End Property

            'Returns the text contents of a cell
Public Property Get Text() As String

End Property

            'Sets the text contents of a cell
Public Property Let Text(ByVal vNewValue As String)

End Property
