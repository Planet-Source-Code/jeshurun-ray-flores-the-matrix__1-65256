VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private theMatrix() As Double
Private matRow As Integer
Private matCol As Integer

Private Sub Initialize_Class()
    Initialize 1, 1

End Sub

Public Sub Initialize(ByVal row As Integer, ByVal col As Integer)
    ReDim theMatrix(1 To row, 1 To col) As Double
    matRow = row
    matCol = col
    Dim a As Integer, b As Integer
    
End Sub

Property Get getRow()
    getRow = matRow
    
End Property

Property Get getColumn()
    getColumn = matCol
    
End Property

Sub setElementAt(ByVal val As Double, ByVal rowpos As Integer, ByVal colpos As Integer)
    theMatrix(rowpos, colpos) = val
    
End Sub
Function getElementAt(ByVal rowpos As Integer, ByVal colpos As Integer) As Double
    getElementAt = theMatrix(rowpos, colpos)
    
End Function
Sub clone(ByVal from As Matrix)
    matRow = from.getRow
    matCol = from.getColumn
    
    ReDim theMatrix(1 To matRow, 1 To matCol) As Double
    Dim a As Integer, b As Integer
    
    For a = 1 To matRow Step 1
        For b = 1 To matCol Step 1
            theMatrix(a, b) = from.getElementAt(a, b)
        Next
    Next
End Sub
Function toString() As String
    Dim tostr As String
    Dim a As Integer, b As Integer
    
    tostr = ""
    
    For a = 1 To matRow Step 1
        For b = 1 To matCol Step 1
          tostr = tostr & theMatrix(a, b) & " "
        Next
        tostr = tostr & vbCrLf
    Next
    
    toString = tostr
    
End Function
