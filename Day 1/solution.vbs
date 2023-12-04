' ====================================================
'
' Advent of Code 2023 - Day 1
' https://adventofcode.com/2023/day/1
' Day 1 - Part 1 and 2
'
' VBScript Only Challenge, by 89mpxf | https://github.com/89mpxf/vbscript-2023_aoc
'
' NOTE: The challenge explanation/example does not explain that combined words, i.e.
'       "oneight" should be treated as "1" and "8" and not "1ight". This had me pulling
'       my hair out and I had to jank my way around it. Credit to this Reddit post:
'       https://www.reddit.com/r/adventofcode/comments/1884fpl/2023_day_1for_those_who_stuck_on_part_2/
'
' ====================================================

' Create system objects
Set fso = CreateObject("Scripting.FileSystemObject")

' Create output buffer
Dim digitsum, totalsum
digitsum = 0
totalsum = 0

' Define number translation table
Dim numbersDict
Set numbersDict = CreateObject("Scripting.Dictionary")
numbersDict.Add "one", 1
numbersDict.Add "two", 2
numbersDict.Add "three", 3
numbersDict.Add "four", 4
numbersDict.Add "five", 5
numbersDict.Add "six", 6
numbersDict.Add "seven", 7
numbersDict.Add "eight", 8
numbersDict.Add "nine", 9

' Define combined words translation table
Dim fixesDict
Set fixesDict = CreateObject("Scripting.Dictionary")
fixesDict.Add "oneight", "18"
fixesDict.Add "twone", "21"
fixesDict.Add "threeight", "38"
fixesDict.Add "fiveight", "58"
fixesDict.Add "sevenine", "79"
fixesDict.Add "eightwo", "82"
fixesDict.Add "eightree", "83"
fixesDict.Add "nineight", "98"

' Create necessary regex objects
Set digitRegex = New RegExp
digitRegex.Global = True
digitRegex.Pattern = "[0-9]"

Set fixRegex = New RegExp
fixRegex.Pattern = "(oneight|twone|threeight|fiveight|sevenine|eightwo|eightree|nineight)"
fixRegex.Global = True
fixRegex.IgnoreCase = True

Set spellRegex = New RegExp
spellRegex.Pattern = "(one|two|three|four|five|six|seven|eight|nine)"
spellRegex.Global = True
spellRegex.IgnoreCase = True

' Open challenge input file for reading
Set file = fso.OpenTextFile("input.txt", 1)

' Read file line by line
Do Until file.AtEndOfStream
    ' Read line
    line = file.ReadLine

    ' Find all digits in line
    Set digitMatches = digitRegex.Execute(line)

    ' Get first and last digits
    first = digitMatches(0).Value
    last = digitMatches(digitMatches.Count - 1).Value

    ' Multiply first digit by 10 and add second digit
    digitsum = digitsum + ((first * 10) + last)

    ' Copy line for replacement
    replacedLine = line

    ' Apply fixes to combined spelled out words
    Set fixMatches = fixRegex.Execute(replacedLine)
    For Each match In fixMatches
        replacedLine = Replace(replacedLine, match.Value, CStr(fixesDict(match.Value)))
    Next

    ' Replace all spelled out numbers with digits
    Set wordMatches = spellRegex.Execute(replacedLine)
    For Each match In wordMatches
        replacedLine = Replace(replacedLine, match.Value, CStr(numbersDict(match.Value)))
    Next

    ' Find all digits in line
    Set allMatches = digitRegex.Execute(replacedLine)

    ' Get combined first and last digits
    first = allMatches(0).Value
    last = allMatches(allMatches.Count - 1).Value

    ' Multiply first digit by 10 and add second digit
    totalsum = totalsum + ((first * 10) + last)
Loop

' Close file
file.Close

' Output results
MsgBox "The solution to Part 1 of Day 1's problem is: " & digitsum, vbInformation, "Solution"
MsgBox "The solution to Part 2 of Day 1's problem is: " & totalsum, vbInformation, "Solution"