' ====================================================
'
' Advent of Code 2023 - Day 2
' https://adventofcode.com/2023/day/2
' Day 2 - Part 1 and 2
'
' VBScript Only Challenge, by 89mpxf | https://github.com/89mpxf/vbscript-2023_aoc
'
' ====================================================

' Create system objects
Set fso = CreateObject("Scripting.FileSystemObject")

' Define maximum possible colors in bag
Const maxRed = 12
Const maxGreen = 13
Const maxBlue = 14

' Read input file
Set inputFile = fso.OpenTextFile("input.txt", 1)

' Define output flag
validIDSum = 0
totalPower = 0

' Define regex logic
Set gameIDRegex = New RegExp
gameIDRegex.Pattern = "Game (\d+):"

Set turnRedRegex = New RegExp
turnRedRegex.Pattern = "(\d+) red"

Set turnGreenRegex = New RegExp
turnGreenRegex.Pattern = "(\d+) green"

Set turnBlueRegex = New RegExp
turnBlueRegex.Pattern = "(\d+) blue"

' Read file line by line
Do Until inputFile.AtEndOfStream
    ' Read line
    line = inputFile.ReadLine

    ' Get game ID from line
    Set gameIDMatch = gameIDRegex.Execute(line)
    gameID = gameIDMatch(0).SubMatches(0)

    ' Remove game ID from line
    line = Mid(line, Len(gameID) + 8)
    turns = Split(line, ";")

    ' Define game minimum flags
    minRed = 0
    minGreen = 0
    minBlue = 0

    ' Iterate over each turn in game
    validGame = True
    For Each turn In turns

        turn = Trim(turn)

        ' Get red count from turn
        Set turnRedMatch = turnRedRegex.Execute(turn)
        If turnRedMatch.Count > 0 Then
            turnRed = turnRedMatch(0).SubMatches(0)

            ' Determine game validity
            If CInt(turnRed) > maxRed Then
                validGame = False
            End If

            ' Determine game minimums
            If CInt(turnRed) > minRed Then
                minRed = CInt(turnRed)
            End If
        End If

        ' Get green count from turn
        Set turnGreenMatch = turnGreenRegex.Execute(turn)
        If turnGreenMatch.Count > 0 Then
            turnGreen = turnGreenMatch(0).SubMatches(0)

            ' Determine game validity
            If CInt(turnGreen) > maxGreen Then
                validGame = False
            End If

            ' Determine game minimums
            If CInt(turnGreen) > minGreen Then
                minGreen = CInt(turnGreen)
            End If
        End If

        ' Get blue count from turn
        Set turnBlueMatch = turnBlueRegex.Execute(turn)
        If turnBlueMatch.Count > 0 Then
            turnBlue = turnBlueMatch(0).SubMatches(0)

            ' Determine game validity
            If CInt(turnBlue) > maxBlue Then
                validGame = False
            End If

            ' Determine game minimums
            If CInt(turnBlue) > minBlue Then
                minBlue = CInt(turnBlue)
            End If
        End If

    Next

    totalPower = totalPower + (minRed * minGreen * minBlue)

    ' If game is valid, add to total
    If validGame Then
        validIDSum = validIDSum + CInt(gameID)
    End If
Loop

' Close input file
inputFile.Close

' Output result
MsgBox "The solution to Part 1 of Day 2's problem is: " & validIDSum, vbInformation, "Solution"
MsgBox "The solution to Part 2 of Day 2's problem is: " & totalPower, vbInformation, "Solution"