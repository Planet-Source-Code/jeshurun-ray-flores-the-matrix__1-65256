Attribute VB_Name = "modOperations"
Public Function Multiplication(ByVal mat1 As Matrix, ByVal mat2 As Matrix) As Matrix
    Dim temp As New Matrix
    Dim a As Integer, b As Integer, c As Integer
    
    temp.Initialize mat1.getRow, mat2.getColumn
    
    For a = 1 To temp.getRow Step 1
        For b = 1 To temp.getColumn Step 1
            Dim sum As Double
            sum = 0
            For c = 1 To mat1.getColumn Step 1
                Dim prod As Integer
                prod = mat1.getElementAt(a, c) * mat2.getElementAt(c, b)
                sum = sum + prod
            Next
            temp.setElementAt sum, a, b
        Next
    Next
    
    Set Multiplication = temp
    
End Function
Public Function Subtraction(ByVal mat1 As Matrix, ByVal mat2 As Matrix) As Matrix
    Dim temp As New Matrix, diffMat As New Matrix
    
    Set temp = scalarMultiplication(mat2, -1)
    Set diffMat = Addition(mat1, temp)
    
    Set Subtraction = diffMat
    
End Function

Public Function scalarMultiplication(ByVal mat As Matrix, ByVal finalConst As Double) As Matrix
    Dim a As Integer, b As Integer
    Dim result As New Matrix
    result.Initialize mat.getRow, mat.getColumn
    
    For a = 1 To result.getRow Step 1
        For b = 1 To result.getColumn Step 1
            Dim prod As Double
            prod = mat.getElementAt(a, b) * finalConst
            result.setElementAt prod, a, b
            
        Next
    Next
    
    Set scalarMultiplication = result

End Function

Public Function Addition(ByVal matA As Matrix, ByVal matB As Matrix) As Matrix
    Dim result As New Matrix
    result.Initialize matA.getRow, matA.getColumn
    Dim a As Integer, b As Integer
    
    For a = 1 To result.getRow Step 1
        For b = 1 To result.getColumn Step 1
            Dim sum As Double
            sum = matA.getElementAt(a, b) + matB.getElementAt(a, b)
            result.setElementAt sum, a, b
            
        Next
    Next
    Set Addition = result
    
End Function

Public Function isValidAdd(ByVal matA As Matrix, ByVal matB As Matrix) As Boolean
    If matA.getColumn = matB.getColumn And matA.getRow = matB.getRow Then
        isValidAdd = True
    Else
        isValidAdd = False
    End If

End Function

Public Function isValidMult(ByVal matA As Matrix, ByVal matB As Matrix) As Boolean
    If matA.getColumn = matB.getRow Then
        isValidMult = True
    Else
        isValidMult = False
    End If
    
End Function

Public Function check(ByVal opIndex As Integer) As String
    Dim res As String
    Select Case opIndex
        Case 0, 1, 2
            If isValidAdd(mymatrix(1), mymatrix(2)) Then
                check = ""
            Else
                check = "Unable to complete operation! " & vbCrLf & _
                                    "Matrices have unequal dimensions."
            End If
            
        Case 3, 4
            If isValidMult(mymatrix(1), mymatrix(2)) And isValidMult(mymatrix(2), mymatrix(1)) Then
                check = ""
            Else
                Select Case opIndex
                    Case 3
                        If isValidMult(mymatrix(1), mymatrix(2)) = False Then
                            check = "Unable to complete process. " & vbCrLf & _
                                                "Row of Matrix A is not equal to Column of Matrix B."
                        End If
                    Case 4
                        If isValidMult(mymatrix(2), mymatrix(1)) = False Then
                            check = "Unable to complete process. " & vbCrLf & _
                                                "Row of Matrix B is not equal to Column of Matrix A."
                        End If
                End Select
            End If
    End Select
    
End Function
