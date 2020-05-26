Public Const Gray As Long = &H404040
Public Const Green As Long = &H8000&
Public Const DotSymbol As String = "n"
Public HoldButtons As Collection
Public CurrentRoll As Integer
Public CurrentPlayer As Integer
Public Sub LoadHoldButtons()
    If Not HoldButtons Is Nothing Then Exit Sub
    Set HoldButtons = New Collection
    HoldButtons.Add Sheet1.btnHold1
    HoldButtons.Add Sheet1.btnHold2
    HoldButtons.Add Sheet1.btnHold3
    HoldButtons.Add Sheet1.btnHold4
    HoldButtons.Add Sheet1.btnHold5
End Sub
Public Sub ResetHoldButtons()
    LoadHoldButtons
    Dim V As Variant
    Dim B As CommandButton
    For Each V In HoldButtons
        Set B = V
        B.Caption = "Hold"
        B.BackColor = Gray
    Next V
End Sub
Public Function RandomNumber() As Integer
    Randomize
    RandomNumber = Int((6 - 1 + 1) * Rnd + 1)
End Function
Public Function DiceNumber(i As Integer) As Integer
    Dim R As Range: Set R = Range("Dice" & i)
    DiceNumber = WorksheetFunction.CountA(R)
End Function
Public Sub SetHeading(i As Integer)
    Dim ii As Integer
    Range("Headers").Interior.Pattern = xlNone
    For ii = 1 To 2
        Dim R As Range
        If ii = 1 Then
            Set R = Range("Player" & i & "Header")
        Else
            Set R = Range("Player" & i & "Header2")
        End If
        With R.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0.399975585192419
            .PatternTintAndShade = 0
        End With
    Next ii

End Sub
Public Function CountOfOneDie(i As Integer) As Integer
    Dim ii As Integer
    Dim Total As Integer
    For ii = 1 To 5
        If DiceNumber(ii) = DiceNumber(i) Then Total = Total + 1
    Next ii
    CountOfOneDie = Total
End Function
Public Function DiceSum() As Integer
    Dim i As Integer
    Dim Total As Integer
    For i = 1 To 5
        Total = Total + DiceNumber(i)
    Next i
    DiceSum = Total
End Function
Public Function LargestStraight() As Integer
    Dim Unsorted As New Collection
    Dim i As Integer
    For i = 1 To 5
        Unsorted.Add DiceNumber(i)
    Next i
    
    Dim Sorted As New Collection
    Dim Lowest As Integer
    Dim LowestIndex As Integer
    Do While Unsorted.Count > 0
        Lowest = 7
        For i = 1 To Unsorted.Count
            If Unsorted(i) < Lowest Then
                Lowest = Unsorted(i)
                LowestIndex = i
            End If
        Next i
        Sorted.Add Lowest
        Unsorted.Remove LowestIndex
    Loop
    
    Dim ii As Integer
    Dim Consecutive As Integer
    Dim Previous As Integer
    Dim Largest As Integer
    For i = 1 To 4
        Previous = Sorted(i)
        Consecutive = 0
        For ii = i + 1 To 5
            If Sorted(ii) - Previous = 1 Then
                Consecutive = Consecutive + 1
            Else
                Exit For
            End If
            Previous = Sorted(ii)
        Next ii
        If Consecutive > Largest Then Largest = Consecutive
    Next i
    If Largest > 0 Then
        LargestStraight = Largest + 1
    Else
        LargestStraight = 0
    End If
End Function
