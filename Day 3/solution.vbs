' ====================================================
'
' Advent of Code 2023 - Day 3
' https://adventofcode.com/2023/day/3
' Day 3 - Part 1 and 2
'
' VBScript Only Challenge, by 89mpxf | https://github.com/89mpxf/vbscript-2023_aoc
'
' ====================================================

' Create system objects
Set fso = CreateObject("Scripting.FileSystemObject")

' Define adjacent flag matrix
adjacentMatrix = Array()

' Define output value
partTotal = 0
gearTotal = 0

' Define hacky work-around for appending to arrays
Function AppendArray(array, value)
    ReDim Preserve array(UBound(array) + 1)
    array(UBound(array)) = value
    AppendArray = array
End Function

' Define symbol validation function
Function IsSymbol(symbol)
    Dim symbols
    symbols = Array(".", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    IsSymbol = True
    For Each s In symbols
        If symbol = s Then
            IsSymbol = False
            Exit Function
        End If
    Next
End Function

' Define number validation function
Function IsNumber(number)
    Dim numbers
    numbers = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    IsNumber = False
    For Each n In numbers
        If number = n Then
            IsNumber = True
            Exit Function
        End If
    Next
End Function

' Define position validation function
Function IsValidPosition(x, y, rows, cols)
    IsValidPosition = False
    If x >= 0 And x < rows And y >= 0 And y < cols Then
        IsValidPosition = True
    End If
End Function

' Define adjacent validation function for part 1
Function HasAdjacent(matrix, x, y, rows, cols)
    HasAdjacent = False

    If IsValidPosition(x-1, y, rows, cols) Then
        If matrix(x-1)(y) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x+1, y, rows, cols) Then
        If matrix(x+1)(y) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x, y-1, rows, cols) Then 
        If matrix(x)(y-1) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x, y+1, rows, cols) Then
        If matrix(x)(y+1) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x-1, y-1, rows, cols) Then
        If matrix(x-1)(y-1) <= -2 Then
            HasAdjacent = True
        End If
    End If
    
    If IsValidPosition(x-1, y+1, rows, cols) Then 
        If matrix(x-1)(y+1) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x+1, y-1, rows, cols) Then 
        If matrix(x+1)(y-1) <= -2 Then
            HasAdjacent = True
        End If
    End If

    If IsValidPosition(x+1, y+1, rows, cols) Then 
        If matrix(x+1)(y+1) <= -2 Then
            HasAdjacent = True
        End If
    End If

End Function

' Define adjacent validation function for part 2
Function GetAdjacentNumbers(matrix, x, y, rows, cols)
    GetAdjacentNumbers = ""

    ' ...X...
    ' ...*...
    ' .......
    If matrix(x-1)(y) >= 0 And matrix(x-1)(y-1) < 0 And matrix(x-1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y) & ", "
    End If

    ' ....X..
    ' ...*...
    ' .......
    If matrix(x-1)(y+1) >= 0 And matrix(x-1)(y) < 0 And matrix(x-1)(y+2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y+1) & ", "
    End If

    ' ..X....
    ' ...*...
    ' .......
    If matrix(x-1)(y-1) >= 0 And matrix(x-1)(y) < 0 And matrix(x-1)(y-2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-1) & ", "
    End If

    ' ...XX..
    ' ...*...
    ' .......
    If matrix(x-1)(y) >= 0 And matrix(x-1)(y+1) >= 0 And matrix(x-1)(y+2) < 0 And matrix(x-1)(y-1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y) & matrix(x-1)(y+1) & ", "
    End If

    ' ..XX...
    ' ...*...
    ' .......
    If matrix(x-1)(y) >= 0 And matrix(x-1)(y-1) >= 0 And matrix(x-1)(y-2) < 0 And matrix(x-1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-1) & matrix(x-1)(y) & ", "
    End If

    ' ....XX.
    ' ...*...
    ' .......
    If matrix(x-1)(y+1) >= 0 And matrix(x-1)(y+2) >= 0 And matrix(x-1)(y+3) < 0 And matrix(x-1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y+1) & matrix(x-1)(y+2) & ", "
    End If

    ' .XX....
    ' ...*...
    ' .......
    If matrix(x-1)(y-2) >= 0 And matrix(x-1)(y-1) >= 0 And matrix(x-1)(y-3) < 0 And matrix(x-1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-2) & matrix(x-1)(y-1) & ", "
    End If

    ' ..XXX..
    ' ...*...
    ' .......
    If matrix(x-1)(y-1) >= 0 And matrix(x-1)(y) >= 0 And matrix(x-1)(y+1) >= 0 And matrix(x-1)(y-2) < 0 And matrix(x-1)(y+2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-1) & matrix(x-1)(y) & matrix(x-1)(y+1) & ", "
    End If

    ' ...XXX.
    ' ...*...
    ' .......
    If matrix(x-1)(y) >= 0 And matrix(x-1)(y+1) >= 0 And matrix(x-1)(y+2) >= 0 And matrix(x-1)(y-1) < 0 And matrix(x-1)(y+3) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y) & matrix(x-1)(y+1) & matrix(x-1)(y+2) & ", "
    End If

    ' .XXX...
    ' ...*...
    ' .......
    If matrix(x-1)(y-2) >= 0 And matrix(x-1)(y-1) >= 0 And matrix(x-1)(y) >= 0 And matrix(x-1)(y-3) < 0 And matrix(x-1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-2) & matrix(x-1)(y-1) & matrix(x-1)(y) & ", "
    End If

    ' ....XXX
    ' ...*...
    ' .......
    If matrix(x-1)(y+1) >= 0 And matrix(x-1)(y+2) >= 0 And matrix(x-1)(y+3) >= 0 And matrix(x-1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y+1) & matrix(x-1)(y+2) & matrix(x-1)(y+3) & ", "
    End If

    ' XXX....
    ' ...*...
    ' .......
    If matrix(x-1)(y-3) >= 0 And matrix(x-1)(y-2) >= 0 And matrix(x-1)(y-1) >= 0 And matrix(x-1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x-1)(y-3) & matrix(x-1)(y-2) & matrix(x-1)(y-1) & ", "
    End If

    ' .......
    ' ...*X..
    ' .......
    If matrix(x)(y+1) >= 0 And matrix(x)(y+2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y+1) & ", "
    End If

    ' .......
    ' ..X*...
    ' .......
    If matrix(x)(y-1) >= 0 And matrix(x)(y-2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y-1) & ", "
    End If

    ' .......
    ' ...*XX.
    ' .......
    If matrix(x)(y+1) >= 0 And matrix(x)(y+2) >= 0 And matrix(x)(y+3) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y+1) & matrix(x)(y+2) & ", "
    End If

    ' .......
    ' .XX*...
    ' .......
    If matrix(x)(y-2) >= 0 And matrix(x)(y-1) >= 0 And matrix(x)(y-3) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y-2) & matrix(x)(y-1) & ", "
    End If

    ' .......
    ' ...*XXX
    ' .......
    If matrix(x)(y+1) >= 0 And matrix(x)(y+2) >= 0 And matrix(x)(y+3) >= 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y+1) & matrix(x)(y+2) & matrix(x)(y+3) & ", "
    End If

    ' .......
    ' XXX*...
    ' .......
    If matrix(x)(y-3) >= 0 And matrix(x)(y-2) >= 0 And matrix(x)(y-1) >= 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x)(y-3) & matrix(x)(y-2) & matrix(x)(y-1) & ", "
    End If

    ' .......
    ' ...*...
    ' ...X...
    If matrix(x+1)(y) >= 0 And matrix(x+1)(y-1) < 0 And matrix(x+1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y) & ", "
    End If

    ' .......
    ' ...*...
    ' ....X..
    If matrix(x+1)(y+1) >= 0 And matrix(x+1)(y) < 0 And matrix(x+1)(y+2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y+1) & ", "
    End If

    ' .......
    ' ...*...
    ' ..X....
    If matrix(x+1)(y-1) >= 0 And matrix(x+1)(y) < 0 And matrix(x+1)(y-2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-1) & ", "
    End If

    ' .......
    ' ...*...
    ' ...XX..
    If matrix(x+1)(y) >= 0 And matrix(x+1)(y+1) >= 0 And matrix(x+1)(y+2) < 0 And matrix(x+1)(y-1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y) & matrix(x+1)(y+1) & ", "
    End If

    ' .......
    ' ...*...
    ' ..XX...
    If matrix(x+1)(y) >= 0 And matrix(x+1)(y-1) >= 0 And matrix(x+1)(y-2) < 0 And matrix(x+1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-1) & matrix(x+1)(y) & ", "
    End If

    ' .......
    ' ...*...
    ' ....XX.
    If matrix(x+1)(y+1) >= 0 And matrix(x+1)(y+2) >= 0 And matrix(x+1)(y+3) < 0 And matrix(x+1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y+1) & matrix(x+1)(y+2) & ", "
    End If

    ' .......
    ' ...*...
    ' .XX....
    If matrix(x+1)(y-2) >= 0 And matrix(x+1)(y-1) >= 0 And matrix(x+1)(y-3) < 0 And matrix(x+1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-2) & matrix(x+1)(y-1) & ", "
    End If

    ' .......
    ' ...*...
    ' ..XXX..
    If matrix(x+1)(y-1) >= 0 And matrix(x+1)(y) >= 0 And matrix(x+1)(y+1) >= 0 And matrix(x+1)(y-2) < 0 And matrix(x+1)(y+2) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-1) & matrix(x+1)(y) & matrix(x+1)(y+1) & ", "
    End If

    ' .......
    ' ...*...
    ' ...XXX.
    If matrix(x+1)(y) >= 0 And matrix(x+1)(y+1) >= 0 And matrix(x+1)(y+2) >= 0 And matrix(x+1)(y-1) < 0 And matrix(x+1)(y+3) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y) & matrix(x+1)(y+1) & matrix(x+1)(y+2) & ", "
    End If

    ' .......
    ' ...*...
    ' .XXX...
    If matrix(x+1)(y-2) >= 0 And matrix(x+1)(y-1) >= 0 And matrix(x+1)(y) >= 0 And matrix(x+1)(y-3) < 0 And matrix(x+1)(y+1) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-2) & matrix(x+1)(y-1) & matrix(x+1)(y) & ", "
    End If

    ' .......
    ' ...*...
    ' ....XXX
    If matrix(x+1)(y+1) >= 0 And matrix(x+1)(y+2) >= 0 And matrix(x+1)(y+3) >= 0 And matrix(x+1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y+1) & matrix(x+1)(y+2) & matrix(x+1)(y+3) & ", "
    End If

    ' .......
    ' ...*...
    ' XXX....
    If matrix(x+1)(y-3) >= 0 And matrix(x+1)(y-2) >= 0 And matrix(x+1)(y-1) >= 0 And matrix(x+1)(y) < 0 Then
        GetAdjacentNumbers = GetAdjacentNumbers & matrix(x+1)(y-3) & matrix(x+1)(y-2) & matrix(x+1)(y-1) & ", "
    End If

    ' Trim output buffer
    If GetAdjacentNumbers <> "" Then
        GetAdjacentNumbers = Left(GetAdjacentNumbers, Len(GetAdjacentNumbers) - 2)
    End If
End Function

' Read input file
Set inputFile = fso.OpenTextFile("input.txt", 1)

' Read input file line by line
Do Until inputFile.AtEndOfStream
    ' Read line
    line = inputFile.ReadLine

    ' Define matrix line buffer
    buf = Array()

    For i = 1 To Len(line)
        char = Mid(line, i, 1)

        ' Check if symbol
        If char = "*" Then
            ' Append star flag to buffer
            buf = AppendArray(buf, -3)
        ElseIf IsSymbol(char) Then
            ' Append symbol flag to buffer
            buf = AppendArray(buf, -2)
        ElseIf IsNumber(char) Then
            ' Append number flag to buffer
            buf = AppendArray(buf, char)
        Else
            ' Append space flag to buffer
            buf = AppendArray(buf, -1)
        End If
    Next

    ' Append line to matrix
    adjacentMatrix = AppendArray(adjacentMatrix, buf)
Loop

' Close input file
inputFile.Close

' Iterate over matrix rows
For i = LBound(adjacentMatrix) To UBound(adjacentMatrix)

    ' Define column skip flag
    skipCols = 0

    ' Iterate over matrix columns
    For j = LBound(adjacentMatrix(i)) To UBound(adjacentMatrix(i))
        ' Skip columns if necessary, decrease flag
        If skipCols > 0 Then
            skipCols = skipCols - 1
        Else
            ' Part 1 Logic
            If adjacentMatrix(i)(j) >= 0 Then
                If IsValidPosition(i, j+1, UBound(adjacentMatrix) + 1, Len(line)) And adjacentMatrix(i)(j+1) >= 0 Then
                    If IsValidPosition(i, j+2, UBound(adjacentMatrix) + 1, Len(line)) And adjacentMatrix(i)(j+2) >= 0 Then
                        skipCols = 2
                        If HasAdjacent(adjacentMatrix, i, j, UBound(adjacentMatrix) + 1, Len(line)) Or HasAdjacent(adjacentMatrix, i, j+1, UBound(adjacentMatrix) + 1, Len(line)) Or HasAdjacent(adjacentMatrix, i, j+2, UBound(adjacentMatrix) + 1, Len(line)) Then
                            partTotal = partTotal + CInt(adjacentMatrix(i)(j) & adjacentMatrix(i)(j+1) & adjacentMatrix(i)(j+2))
                        End If
                    Else
                        skipCols = 1
                        If HasAdjacent(adjacentMatrix, i, j, UBound(adjacentMatrix) + 1, Len(line)) Or HasAdjacent(adjacentMatrix, i, j+1, UBound(adjacentMatrix) + 1, Len(line)) Then
                            partTotal = partTotal + CInt(adjacentMatrix(i)(j) & adjacentMatrix(i)(j+1))
                        End If
                    End If
                Else
                    If HasAdjacent(adjacentMatrix, i, j, UBound(adjacentMatrix) + 1, Len(line)) Then
                        partTotal = partTotal + CInt(adjacentMatrix(i)(j))
                    End If
                End If

            ' Part 2 Logic
            ElseIf adjacentMatrix(i)(j) = -3 Then
                out = GetAdjacentNumbers(adjacentMatrix, i, j, UBound(adjacentMatrix) + 1, Len(line))
                list = Split(out, ", ")
                If UBound(list) = 1 Then
                    gearTotal = gearTotal + (CInt(list(0)) * CInt(list(1)))
                End If
            End If
        End If
    Next
Next

' Output result
MsgBox "The solution to Part 1 of Day 3's problem is: " & partTotal, vbInformation, "Solution"
MsgBox "The solution to Part 2 of Day 3's problem is: " & gearTotal, vbInformation, "Solution"