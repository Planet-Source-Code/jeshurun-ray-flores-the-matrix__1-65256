VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wd As Long
Private ln As Long

Private Sub Initialize_Class()
    Initialize

End Sub

Public Sub Initialize(ByVal newln As Integer, ByVal newwd As Integer)
    wd = newwd
    ln = newln
    
End Sub

Property Get getWidth() As Long
    getWidth = wd
    
End Property

Property Get getLength()
    getLength = ln
    
End Property


Property Let wd(ByVal newwd As Long)
    wd = newwd
    
End Property

