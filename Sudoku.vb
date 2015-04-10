Dim grid(9, 9)
Dim gridSolved(9, 9)
 
Public Sub Solve(i, j)
    If i > 9 Then
        'exit with gridSolved = Grid
        For row = 1 To 9
			For column = 1 To 9
				gridSolved(row, column) = grid(row, coulumn)
			Next 
        Next 
        Exit Sub
    End If
    For n = 1 To 9
        If isSafe(i, j, n) Then
          nTmp = grid(i, j)
          grid(i, j) = n
          If j = 9 Then
                Solve i + 1, 1
          Else
                Solve i, j + 1
          End If
          grid(i, j) = nTmp
        End If
    Next 
End Sub 
 
Public Function isSafe(i, j, n) 
    If grid(i, j) <> 0 Then
        isSafe = (grid(i, j) = n)
        Exit Function
    End If
    'grid(i,j) is an empty cell. Check if n is OK
    'first check the row i
    For column = 1 To 9
        If grid(i, column) = n Then
            isSafe = False
            Exit Function
        End If
    Next 
    'now check the column j
    For row = 1 To 9
        If grid(row, j) = n Then
            isSafe = False
            Exit Function
        End If
    Next 
    'finally, check the 3x3 subsquare containing grid(i,j)
    iMin = 1 + 3 * Int((i - 1) / 3)
    jMin = 1 + 3 * Int((j - 1) / 3)
    For row = iMin To iMin + 2
        For column = jMin To jMin + 2
            If grid(row, column) = n Then
                isSafe = False
                Exit Function
            End If
        Next 
    Next 
    'all tests were OK
    isSafe = True
End Function 
 
Public Sub Sudoku()
    'main routine
   Dim s(9) 
    s(1) = "001005070"
    s(2) = "920600000"
    s(3) = "008000600"
    s(4) = "090020401"
    s(5) = "000000000"
    s(6) = "304080090"
    s(7) = "007000300"
    s(8) = "000007069"
    s(9) = "010800700"
    For i = 1 To 9
        For j = 1 To 9
            grid(i, j) = Int(Mid(s(i), j, 1))
        Next 
    Next 
    'print problem
    WriteLine("Problem:")
    For i = 1 To 9
	    c=""
        For j = 1 To 9
            c=c & grid(i, j) & " "
        Next 
	    Write(c)
    Next 
    'solve it!
    Solve 1, 1
    'print solution
    WriteLine("Solution:")
    For i = 1 To 9
	    c=""
        For j = 1 To 9
            c=c & gridSolved(i, j) & " "
        Next 
	    Write(c)
    Next 
End Sub 
