Attribute VB_Name = "modMisc"
Public mymatrix(1 To 2) As New Matrix
Public max As Integer
Public elem() As TextBox

Public Sub initMatrix(ByRef newMa3x As Matrix, ByVal where As Integer)
    
End Sub

Public Sub openMatrix(ByRef theMatrix As Matrix)
    Dim col As Integer, row As Integer
    col = theMatrix.getColumn
    row = theMatrix.getRow
    
    Dim a As Integer, b As Integer
    
    For a = 1 To col Step 1
        For b = 1 To row Step 1
            elem(a, b).Text = theMatrix.getElementAt(b, a)
        Next
    Next
End Sub

Public Sub initializeThis(ByRef mat As Matrix)
    Dim a As Integer, b As Integer
    For a = 1 To mat.getRow Step 1
        For b = 1 To mat.getColumn Step 1
            mat.setElementAt elem(b, a).Text, a, b
        Next
    Next
    
    
End Sub

Public Sub hideAll()
    Dim a As Integer, b As Integer
    
    For a = 1 To max Step 1
        For b = 1 To max Step 1
            elem(a, b).Visible = False
        Next
    Next
endAll:
    Exit Sub
End Sub

Public Sub showMatrixFields(ByVal r As Integer, ByVal c As Integer)
    Dim a As Integer, b As Integer
    For a = 1 To c Step 1
        For b = 1 To r Step 1
            elem(a, b).Visible = True
        Next
    Next
    
    If a * ((120 + 400) + 120) > 8250 Then
        dlgProp.Width = c * ((120 + 400) + 120)
    End If
    dlgProp.Height = r * (120 + 400) + 1120
    
End Sub

Public Sub addTillMax()
    ReDim elem(1 To max, 1 To max) As TextBox

    Dim a As Integer, b As Integer
    
    For b = 1 To max Step 1
        For a = 1 To max Step 1
            On Error Resume Next
            Set elem(a, b) = dlgProp.Controls.Add("VB.TextBox", "elem" & a & "_" & b)
            elem(a, b).Move (a - 1) * (120 + 400) + 120, (b - 1) * (120 + 400) + 720, 400, 400
            elem(a, b).Text = 0
        Next
    Next
End Sub
