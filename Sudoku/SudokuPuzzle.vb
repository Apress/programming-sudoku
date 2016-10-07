Public Class SudokuPuzzle
    '---used to represent the values in the grid---
    Private actual(9, 9) As Integer

    '---used to represent the possible values of cells in the grid---
    Private possible(9, 9) As String

    '---indicate if the brute-force subroutine should stop---
    Private BruteForceStop As Boolean = False

    '---used to store the state of the grid---
    Private ActualStack As New Stack(Of Integer(,))()
    Private PossibleStack As New Stack(Of String(,))()

    '---store the total score accumulated---
    Private totalscore As Integer

    '---backup a copy of the Actual array---
    Dim actual_backup(9, 9) As Integer

    '==================================================
    ' Steps to solve the puzzle
    '==================================================
    Private Function SolvePuzzle() As Boolean
        Dim changes As Boolean
        Dim ExitLoop As Boolean = False
        Try
            Do '---Look for Triplets in Columns---
                Do '---Look for Triplets in Rows---
                    Do '---Look for Triplets in Minigrids---
                        Do '---Look for Twins in Columns---
                            Do '---Look for Twins in Rows---
                                Do '---Look for Twins in Minigrids---
                                    Do '---Look for Lone Rangers in Columns---
                                        Do '---Look for Lone Rangers in Rows---
                                            Do  '---Look for Lone Rangers in 
                                                ' Minigrids---
                                                Do '---Perform Col/Row and 
                                                    ' Minigrid Elimination---
                                                    changes = CheckColumnsAndRows()
                                                    If IsPuzzleSolved() Then
                                                        ExitLoop = True
                                                        Exit Do
                                                    End If
                                                Loop Until Not changes

                                                If ExitLoop Then Exit Do
                                                '---Look for Lone Rangers in 
                                                ' Minigrids---
                                                changes = _
                                                   LookForLoneRangersinMinigrids()
                                                If IsPuzzleSolved() Then
                                                    ExitLoop = True
                                                    Exit Do
                                                End If
                                            Loop Until Not changes

                                            If ExitLoop Then Exit Do
                                            '---Look for Lone Rangers in Rows---
                                            changes = LookForLoneRangersinRows()
                                            If IsPuzzleSolved() Then
                                                ExitLoop = True
                                                Exit Do
                                            End If
                                        Loop Until Not changes

                                        If ExitLoop Then Exit Do
                                        '---Look for Lone Rangers in Columns---
                                        changes = LookForLoneRangersinColumns()
                                        If IsPuzzleSolved() Then
                                            ExitLoop = True
                                            Exit Do
                                        End If
                                    Loop Until Not changes

                                    If ExitLoop Then Exit Do
                                    '---Look for Twins in Minigrids---
                                    changes = LookForTwinsinMinigrids()
                                    If IsPuzzleSolved() Then
                                        ExitLoop = True
                                        Exit Do
                                    End If
                                Loop Until Not changes

                                If ExitLoop Then Exit Do
                                '---Look for Twins in Rows---
                                changes = LookForTwinsinRows()
                                If IsPuzzleSolved() Then
                                    ExitLoop = True
                                    Exit Do
                                End If
                            Loop Until Not changes

                            If ExitLoop Then Exit Do
                            '---Look for Twins in Columns---
                            changes = LookForTwinsinColumns()
                            If IsPuzzleSolved() Then
                                ExitLoop = True
                                Exit Do
                            End If
                        Loop Until Not changes

                        If ExitLoop Then Exit Do
                        '---Look for Triplets in Minigrids---
                        changes = LookForTripletsinMinigrids()
                        If IsPuzzleSolved() Then
                            ExitLoop = True
                            Exit Do
                        End If
                    Loop Until Not changes

                    If ExitLoop Then Exit Do
                    '---Look for Triplets in Rows---
                    changes = LookForTripletsinRows()
                    If IsPuzzleSolved() Then
                        ExitLoop = True
                        Exit Do
                    End If
                Loop Until Not changes

                If ExitLoop Then Exit Do
                '---Look for Triplets in Columns---
                changes = LookForTripletsinColumns()
                If IsPuzzleSolved() Then
                    ExitLoop = True
                    Exit Do
                End If
            Loop Until Not changes

        Catch ex As Exception
            Throw New Exception("Invalid Move")
        End Try

        If IsPuzzleSolved() Then
            Return True
        Else
            Return False
        End If
    End Function

    '==================================================
    ' Calculates the possible values for all the cell
    '==================================================
    Private Function CheckColumnsAndRows() As Boolean
        Dim changes As Boolean = False
        '---check all cells---
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                If actual(col, row) = 0 Then
                    Try
                        possible(col, row) = CalculatePossibleValues(col, row)
                    Catch ex As Exception
                        Throw New Exception("Invalid Move")
                    End Try

                    If possible(col, row).Length = 1 Then

                        '---number is confirmed---
                        actual(col, row) = CInt(possible(col, row))
                        changes = True

                        '---accumulate the total score---
                        totalscore += 1
                    End If
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Calculates the possible values for a cell
    '==================================================
    Private Function CalculatePossibleValues( _
                    ByVal col As Integer, _
                    ByVal row As Integer) _
                    As String
        Dim str As String
        If possible(col, row) = String.Empty Then
            str = "123456789"
        Else
            str = possible(col, row)
        End If

        Dim r, c As Integer

        '---Step (1) check by column---
        For r = 1 To 9
            If actual(col, r) <> 0 Then
                '---that means there is a actual value in it---
                str = str.Replace(actual(col, r).ToString(), String.Empty)
            End If
        Next
        '---Step (2) check by row---
        For c = 1 To 9
            If actual(c, row) <> 0 Then
                '---that means there is a actual value in it---
                str = str.Replace(actual(c, row).ToString(), String.Empty)
            End If
        Next

        '---Step (3) check within the minigrid---
        Dim startC, startR As Integer
        startC = col - ((col - 1) Mod 3)
        startR = row - ((row - 1) Mod 3)
        For rr As Integer = startR To startR + 2
            For cc As Integer = startC To startC + 2
                If actual(cc, rr) <> 0 Then
                    str = str.Replace(actual(cc, rr).ToString(), String.Empty)
                End If
            Next
        Next

        '---if possible value is string.Empty, then error---
        If str = String.Empty Then
            Throw New Exception("Invalid Move")
        End If
        Return str
    End Function

    '==================================================
    ' Look for lone rangers in Minigrids
    '==================================================
    Private Function LookForLoneRangersinMinigrids() As Boolean
        Dim changes As Boolean = False
        Dim NextMiniGrid As Boolean
        Dim occurrence As Integer
        Dim cPos, rPos As Integer

        '---check for each number from 1 to 9---
        For n As Integer = 1 To 9

            '---check the 9 mini-grids---
            For r As Integer = 1 To 9 Step 3
                For c As Integer = 1 To 9 Step 3
                    NextMiniGrid = False

                    '---check within the mini-grid---
                    occurrence = 0
                    For rr As Integer = 0 To 2
                        For cc As Integer = 0 To 2
                            If actual(c + cc, r + rr) = 0 AndAlso _
                               possible(c + cc, r + rr).Contains( _
                               n.ToString()) Then
                                occurrence += 1
                                cPos = c + cc
                                rPos = r + rr
                                If occurrence > 1 Then
                                    NextMiniGrid = True
                                    Exit For
                                End If
                            End If
                        Next
                        If NextMiniGrid Then Exit For
                    Next

                    If (Not NextMiniGrid) AndAlso occurrence = 1 Then
                        '---that means number is confirmed---
                        actual(cPos, rPos) = n
                        changes = True
                        '---accumulate the total score---
                        totalscore += 2
                    End If
                Next
            Next
        Next
        Return changes
    End Function

    '=========================================================
    'Look for Lone Rangers in Rows
    '=========================================================
    Private Function LookForLoneRangersinRows() As Boolean
        Dim changes As Boolean = False
        Dim occurrence As Integer
        Dim cPos, rPos As Integer

        '---check by row----
        For r As Integer = 1 To 9
            For n As Integer = 1 To 9
                occurrence = 0
                For c As Integer = 1 To 9
                    If actual(c, r) = 0 AndAlso _
                       possible(c, r).Contains(n.ToString()) Then
                        occurrence += 1

                        '---if multiple occurrence, not a lone ranger anymore
                        If occurrence > 1 Then Exit For
                        cPos = c
                        rPos = r
                    End If
                Next
                If occurrence = 1 Then
                    '--number is confirmed---
                    actual(cPos, rPos) = n
                    changes = True

                    '---accumulate the total score---
                    totalscore += 2
                End If
            Next
        Next
        Return changes
    End Function

    '=========================================================
    'Look for Lone Rangers in Columns
    '=========================================================
    Private Function LookForLoneRangersinColumns() As Boolean
        Dim changes As Boolean = False
        Dim occurrence As Integer
        Dim cPos, rPos As Integer

        '----check by column----
        For c As Integer = 1 To 9
            For n As Integer = 1 To 9
                occurrence = 0
                For r As Integer = 1 To 9
                    If actual(c, r) = 0 AndAlso _
                       possible(c, r).Contains(n.ToString()) Then
                        occurrence += 1

                        '---if multiple occurrence, not a lone ranger anymore
                        If occurrence > 1 Then Exit For
                        cPos = c
                        rPos = r
                    End If
                Next
                If occurrence = 1 Then
                    '--number is confirmed---
                    actual(cPos, rPos) = n
                    changes = True

                    '---accumulate the total score---
                    totalscore += 2
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Twins in Minigrids
    '==================================================
    Private Function LookForTwinsinMinigrids() As Boolean
        Dim changes As Boolean = False

        '---look for twins in each cell---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '---scan by the mini-grid that the current cell is in---
                    Dim startC, startR As Integer
                    startC = c - ((c - 1) Mod 3)
                    startR = r - ((r - 1) Mod 3)
                    For rr As Integer = startR To startR + 2
                        For cc As Integer = startC To startC + 2

                            '---for cells other than the pair of twins---
                            If (Not ((cc = c) AndAlso (rr = r))) AndAlso _
                               possible(cc, rr) = possible(c, r) Then

                                '---remove the twins from all the other possible 
                                ' values in the minigrid---
                                For rrr As Integer = startR To startR + 2
                                    For ccc As Integer = startC To startC + 2
                                        If actual(ccc, rrr) = 0 AndAlso _
                                           possible(ccc, rrr) <> _
                                           possible(c, r) Then

                                            '---save a copy of the original 
                                            ' possible values (twins)---
                                            Dim original_possible As String = _
                                               possible(ccc, rrr)

                                            '---remove first twin number from 
                                            ' possible values---
                                            possible(ccc, rrr) = _
                                               possible(ccc, rrr).Replace( _
                                               possible(c, r)(0), String.Empty)

                                            '---remove second twin number from 
                                            ' possible values---
                                            possible(ccc, rrr) = _
                                               possible(ccc, rrr).Replace( _
                                               possible(c, r)(1), String.Empty)

                                            '---if the possible values are 
                                            ' modified, then set the changes 
                                            ' variable to true to indicate 
                                            ' that the possible values of cells 
                                            ' in the minigrid have been modified---
                                            If original_possible <> _
                                               possible(ccc, rrr) Then
                                                changes = True
                                            End If

                                            '---if possible value reduces to 
                                            ' empty string, then the user has 
                                            ' placed a move that results in 
                                            ' the puzzle not solvable---
                                            If possible(ccc, rrr) = _
                                               String.Empty Then
                                                Throw New Exception("Invalid Move")
                                            End If

                                            '---if left with 1 possible value 
                                            ' for the current cell, cell is 
                                            ' confirmed---
                                            If possible(ccc, rrr).Length = 1 Then
                                                actual(ccc, rrr) = _
                                                   CInt(possible(ccc, rrr))

                                                '---accumulate the total score---
                                                totalscore += 3
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                        Next
                    Next
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Twins in Rows
    '==================================================
    Private Function LookForTwinsinRows() As Boolean
        Dim changes As Boolean = False

        '---for each row, check each column in the row---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '--scan columns in this row---
                    For cc As Integer = c + 1 To 9
                        If (possible(cc, r) = possible(c, r)) Then

                            '---remove the twins from all the other possible 
                            ' values in the row---
                            For ccc As Integer = 1 To 9
                                If (actual(ccc, r) = 0) AndAlso _
                                   (ccc <> c) AndAlso (ccc <> cc) Then

                                    '---save a copy of the original possible 
                                    ' values (twins)---
                                    Dim original_possible As String = _
                                       possible(ccc, r)

                                    '---remove first twin number from possible 
                                    ' values---
                                    possible(ccc, r) = possible(ccc, r).Replace( _
                                       possible(c, r)(0), String.Empty)

                                    '---remove second twin number from possible  
                                    ' values---
                                    possible(ccc, r) = possible(ccc, r).Replace( _
                                       possible(c, r)(1), String.Empty)

                                    '---if the possible values are modified, then 
                                    ' set the changes variable to true to indicate 
                                    ' that the possible values of cells in the  
                                    ' minigrid have been modified---
                                    If original_possible <> possible(ccc, r) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string, 
                                    ' then the user has placed a move that results 
                                    ' in the puzzle not solvable---
                                    If possible(ccc, r) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(ccc, r).Length = 1 Then
                                        actual(ccc, r) = CInt(possible(ccc, r))

                                        '---accumulate the total score---
                                        totalscore += 3
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Twins in Columns
    '==================================================
    Private Function LookForTwinsinColumns() As Boolean
        Dim changes As Boolean = False

        '---for each column, check each row in the column---
        For c As Integer = 1 To 9
            For r As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '--scan rows in this column---
                    For rr As Integer = r + 1 To 9
                        If (possible(c, rr) = possible(c, r)) Then

                            '---remove the twins from all the other possible 
                            ' values in the row---
                            For rrr As Integer = 1 To 9
                                If (actual(c, rrr) = 0) AndAlso _
                                   (rrr <> r) AndAlso (rrr <> rr) Then

                                    '---save a copy of the original possible 
                                    ' values (twins)---
                                    Dim original_possible As String = _
                                       possible(c, rrr)

                                    '---remove first twin number from possible  
                                    ' values---
                                    possible(c, rrr) = possible(c, rrr).Replace( _
                                       possible(c, r)(0), String.Empty)

                                    '---remove second twin number from possible 
                                    ' values---
                                    possible(c, rrr) = possible(c, rrr).Replace( _
                                       possible(c, r)(1), String.Empty)

                                    '---if the possible values are modified, then 
                                    'set the changes variable to true to indicate 
                                    ' that the possible values of cells in the 
                                    ' minigrid have been modified---
                                    If original_possible <> possible(c, rrr) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string, 
                                    ' then the user has placed a move that results 
                                    ' in the puzzle not solvable---
                                    If possible(c, rrr) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(c, rrr).Length = 1 Then
                                        actual(c, rrr) = CInt(possible(c, rrr))

                                        '---accumulate the total score---
                                        totalscore += 3
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Triplets in Minigrids
    '==================================================
    Private Function LookForTripletsinMinigrids() As Boolean
        Dim changes As Boolean = False

        '---check each cell---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '--- three possible values; check for triplets---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 3 Then

                    '---first potential triplet found---
                    Dim tripletsLocation As String = c.ToString() & r.ToString()

                    '---scan by mini-grid---
                    Dim startC, startR As Integer
                    startC = c - ((c - 1) Mod 3)
                    startR = r - ((r - 1) Mod 3)
                    For rr As Integer = startR To startR + 2
                        For cc As Integer = startC To startC + 2
                            If (Not ((cc = c) AndAlso (rr = r))) AndAlso _
                               ((possible(cc, rr) = possible(c, r)) OrElse _
                                (possible(cc, rr).Length = 2 AndAlso _
                                possible(c, r).Contains( _
                                   possible(cc, rr)(0).ToString()) AndAlso _
                                possible(c, r).Contains( _
                                   possible(cc, rr)(1).ToString()))) Then

                                '---save the coorindates of the triplets
                                tripletsLocation &= cc.ToString() & rr.ToString()
                            End If
                        Next
                    Next

                    '--found 3 cells as triplets; remove all from the other 
                    ' cells---
                    If tripletsLocation.Length = 6 Then

                        '---remove each cell's possible values containing the 
                        ' triplet---
                        For rrr As Integer = startR To startR + 2
                            For ccc As Integer = startC To startC + 2

                                '---look for the cell that is not part of the 
                                ' 3 cells found---
                                If actual(ccc, rrr) = 0 AndAlso _
                                   ccc <> CInt(tripletsLocation(0).ToString()) _
                                   AndAlso _
                                   rrr <> CInt(tripletsLocation(1).ToString()) _
                                   AndAlso _
                                   ccc <> CInt(tripletsLocation(2).ToString()) _
                                   AndAlso _
                                   rrr <> CInt(tripletsLocation(3).ToString()) _
                                   AndAlso _
                                   ccc <> CInt(tripletsLocation(4).ToString()) _
                                   AndAlso _
                                   rrr <> CInt(tripletsLocation(5).ToString()) Then

                                    '---save the original possible values---
                                    Dim original_possible As String = _
                                       possible(ccc, rrr)

                                    '---remove first triplet number from possible 
                                    ' values---
                                    possible(ccc, rrr) = _
                                       possible(ccc, rrr).Replace( _
                                          possible(c, r)(0), String.Empty)

                                    '---remove second triplet number from possible 
                                    ' values---
                                    possible(ccc, rrr) = _
                                       possible(ccc, rrr).Replace( _
                                          possible(c, r)(1), String.Empty)

                                    '---remove third triplet number from possible 
                                    ' values---
                                    possible(ccc, rrr) = _
                                       possible(ccc, rrr).Replace( _
                                          possible(c, r)(2), String.Empty)

                                    '---if the possible values are modified, then 
                                    ' set the changes variable to true to indicate 
                                    ' that the possible values of cells in the 
                                    ' minigrid have been modified---
                                    If original_possible <> possible(ccc, rrr) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string, 
                                    ' then the user has placed a move that results 
                                    ' in the puzzle not solvable---
                                    If possible(ccc, rrr) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(ccc, rrr).Length = 1 Then
                                        actual(ccc, rrr) = CInt(possible(ccc, rrr))

                                        '---accumulate the total score---
                                        totalscore += 4
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Triplets in Rows
    '==================================================
    Private Function LookForTripletsinRows() As Boolean
        Dim changes As Boolean = False

        '---for each row, check each column in the row---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '--- three possible values; check for triplets---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 3 Then

                    '---first potential triplet found---
                    Dim tripletsLocation As String = c.ToString() & r.ToString()

                    '---scans columns in this row---
                    For cc As Integer = 1 To 9

                        '---look for other triplets---
                        If (cc <> c) AndAlso _
                           ((possible(cc, r) = possible(c, r)) OrElse _
                            (possible(cc, r).Length = 2 AndAlso _
                             possible(c, r).Contains( _
                                possible(cc, r)(0).ToString()) AndAlso _
                             possible(c, r).Contains( _
                                possible(cc, r)(1).ToString()))) Then

                            '---save the coorindates of the triplet---
                            tripletsLocation &= cc.ToString() & r.ToString()
                        End If
                    Next

                    '--found 3 cells as triplets; remove all from the other 
                    ' cells---
                    If tripletsLocation.Length = 6 Then

                        '---remove each cell's possible values containing the 
                        ' triplet---
                        For ccc As Integer = 1 To 9
                            If actual(ccc, r) = 0 AndAlso _
                               ccc <> CInt(tripletsLocation(0).ToString()) _
                               AndAlso _
                               ccc <> CInt(tripletsLocation(2).ToString()) _
                               AndAlso _
                               ccc <> CInt(tripletsLocation(4).ToString()) Then

                                '---save the original possible values---
                                Dim original_possible As String = possible(ccc, r)

                                '---remove first triplet number from possible 
                                ' values---
                                possible(ccc, r) = _
                                   possible(ccc, r).Replace( _
                                      possible(c, r)(0), String.Empty)

                                '---remove second triplet number from possible 
                                ' values---
                                possible(ccc, r) = _
                                   possible(ccc, r).Replace( _
                                      possible(c, r)(1), String.Empty)

                                '---remove third triplet number from possible 
                                ' values---
                                possible(ccc, r) = _
                                   possible(ccc, r).Replace( _
                                      possible(c, r)(2), String.Empty)

                                '---if the possible values are modified, then set  
                                ' the changes variable to true to indicate that 
                                ' the possible values of cells in the minigrid 
                                ' have been modified---
                                If original_possible <> possible(ccc, r) Then
                                    changes = True
                                End If

                                '---if possible value reduces to empty string, 
                                ' then the user has placed a move that results 
                                ' in the puzzle not solvable---
                                If possible(ccc, r) = String.Empty Then
                                    Throw New Exception("Invalid Move")
                                End If

                                '---if left with 1 possible value for the current 
                                ' cell, cell is confirmed---
                                If possible(ccc, r).Length = 1 Then
                                    actual(ccc, r) = CInt(possible(ccc, r))

                                    '---accumulate the total score---
                                    totalscore += 4
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Look for Triplets in Columns
    '==================================================
    Private Function LookForTripletsinColumns() As Boolean
        Dim changes As Boolean = False

        '---for each column, check each row in the column---
        For c As Integer = 1 To 9
            For r As Integer = 1 To 9

                '--- three possible values; check for triplets---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 3 Then

                    '---first potential triplet found---
                    Dim tripletsLocation As String = c.ToString() & r.ToString()

                    '---scans rows in this column---
                    For rr As Integer = 1 To 9
                        If (rr <> r) AndAlso _
                           ((possible(c, rr) = possible(c, r)) OrElse _
                           (possible(c, rr).Length = 2 AndAlso _
                            possible(c, r).Contains( _
                               possible(c, rr)(0).ToString()) AndAlso _
                            possible(c, r).Contains( _
                               possible(c, rr)(1).ToString()))) Then

                            '---save the coorindates of the triplet---
                            tripletsLocation += c.ToString() & rr.ToString()
                        End If
                    Next

                    '--found 3 cells as triplets; remove all from the other cells---
                    If tripletsLocation.Length = 6 Then

                        '---remove each cell's possible values containing the 
                        ' triplet---
                        For rrr As Integer = 1 To 9
                            If actual(c, rrr) = 0 AndAlso _
                               rrr <> CInt(tripletsLocation(1).ToString()) _
                               AndAlso _
                               rrr <> CInt(tripletsLocation(3).ToString()) _
                               AndAlso _
                               rrr <> CInt(tripletsLocation(5).ToString()) Then

                                '---save the original possible values---
                                Dim original_possible As String = possible(c, rrr)

                                '---remove first triplet number from possible 
                                ' values---
                                possible(c, rrr) = _
                                   possible(c, rrr).Replace( _
                                      possible(c, r)(0), String.Empty)

                                '---remove second triplet number from possible 
                                ' values---
                                possible(c, rrr) = _
                                   possible(c, rrr).Replace( _
                                      possible(c, r)(1), String.Empty)

                                '---remove third triplet number from possible 
                                ' values---
                                possible(c, rrr) = _
                                   possible(c, rrr).Replace( _
                                      possible(c, r)(2), String.Empty)

                                '---if the possible values are modified, then set 
                                ' the changes variable to true to indicate that 
                                ' the possible values of cells in the minigrid 
                                ' have been modified---
                                If original_possible <> possible(c, rrr) Then
                                    changes = True
                                End If

                                '---if possible value reduces to empty string, 
                                ' then the user has placed a move that results 
                                ' in the puzzle not solvable---
                                If possible(c, rrr) = String.Empty Then
                                    Throw New Exception("Invalid Move")
                                End If

                                '---if left with 1 possible value for the current 
                                ' cell, cell is confirmed---
                                If possible(c, rrr).Length = 1 Then
                                    actual(c, rrr) = CInt(possible(c, rrr))

                                    '---accumulate the total score---
                                    totalscore += 4
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Next
        Return changes
    End Function

    '=========================================================
    ' Find the cell with the small number of possible values
    '=========================================================
    Private Sub FindCellWithFewestPossibleValues( _
       ByRef col As Integer, ByRef row As Integer)
        Dim min As Integer = 10
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9
                If actual(c, r) = 0 AndAlso possible(c, r).Length < min Then
                    min = possible(c, r).Length
                    col = c
                    row = r
                End If
            Next
        Next
    End Sub

    '==================================================
    ' Solve puzzle by brute force
    '==================================================
    Private Sub SolvePuzzleByBruteForce()
        Dim c, r As Integer

        '---accumulate the total score---
        totalscore += 5

        '---find out which cell has the smallest number of possible values---
        FindCellWithFewestPossibleValues(c, r)

        '---get the possible values for the chosen cell---
        Dim possibleValues As String = possible(c, r)

        '---randomize the possible values----
        RandomizeThePossibleValues(possibleValues)
        '-------------------

        '---push the actual and possible stacks into the stack---
        ActualStack.Push(CType(actual.Clone(), Integer(,)))
        PossibleStack.Push(CType(possible.Clone(), String(,)))

        '---select one value and try---
        For i As Integer = 0 To possibleValues.Length - 1
            actual(c, r) = CInt(possibleValues(i).ToString())
            Try
                If SolvePuzzle() Then
                    '---if the puzzle is solved, the recursion can stop now---
                    BruteForceStop = True
                    Return
                Else
                    '---no problem with current selection, proceed with next 
                    ' cell---
                    SolvePuzzleByBruteForce()
                    If BruteForceStop Then Return
                End If
            Catch ex As Exception
                '---accumulate the total score---
                totalscore += 5
                actual = ActualStack.Pop()
                possible = PossibleStack.Pop()
            End Try
        Next
    End Sub

    '==================================================
    ' Check if the puzzle is solved
    '==================================================
    Private Function IsPuzzleSolved() As Boolean
        Dim pattern As String
        Dim r, c As Integer

        '---check row by row---
        For r = 1 To 9
            pattern = "123456789"
            For c = 1 To 9
                pattern = pattern.Replace(actual(c, r).ToString(), String.Empty)
            Next
            If pattern.Length > 0 Then
                Return False
            End If
        Next

        '---check col by col---
        For c = 1 To 9
            pattern = "123456789"
            For r = 1 To 9
                pattern = pattern.Replace(actual(c, r).ToString(), String.Empty)
            Next
            If pattern.Length > 0 Then
                Return False
            End If
        Next

        '---check by minigrid---
        For c = 1 To 9 Step 3
            pattern = "123456789"
            For r = 1 To 9 Step 3
                For cc As Integer = 0 To 2
                    For rr As Integer = 0 To 2
                        pattern = pattern.Replace( _
                           actual(c + cc, r + rr).ToString(), String.Empty)
                    Next
                Next
            Next
            If pattern.Length > 0 Then
                Return False
            End If
        Next
        Return True
    End Function

    '=========================================================
    ' Randomly swap the list of possible values
    '=========================================================
    Private Sub RandomizeThePossibleValues(ByRef str As String)
        Dim s(str.Length - 1) As Char
        Dim i, j As Integer
        Dim temp As Char
        Randomize()
        s = str.ToCharArray
        For i = 0 To str.Length - 1
            j = CInt((str.Length - i + 1) * Rnd() + i) Mod str.Length
            '---swap the chars---
            temp = s(i)
            s(i) = s(j)
            s(j) = temp
        Next i
        str = s
    End Sub

    '============================================================
    ' Generate a random number between the specified range
    '============================================================
    Private Function RandomNumber(ByVal min As Integer, _
                         ByVal max As Integer) As Integer
        Return Int((max - min + 1) * Rnd()) + min
    End Function

    '============================================================
    ' Get Puzzle
    '============================================================
    Public Function GetPuzzle(ByVal level As Integer) As String
        Dim score As Integer
        Dim result As String
        Do
            result = GenerateNewPuzzle(level, score)
            If result <> String.Empty Then
                '---check if puzzle matches the level of difficult---
                Select Case level
                    '---average for this level is 44---
                    Case 1
                        If score >= 42 And score <= 46 Then
                            Exit Do
                        End If
                        '---average for this level is 51---
                    Case 2
                        If score >= 49 And score <= 53 Then
                            Exit Do
                        End If
                        '---average for this level is 58---
                    Case 3
                        If score >= 56 And score <= 60 Then
                            Exit Do
                        End If
                        '---average for this level is 114---
                    Case 4
                        If score >= 112 And score <= 116 Then
                            Exit Do
                        End If
                End Select
            End If
        Loop Until False
        Return result
    End Function

    '============================================================
    '  Create empty cells in the grid
    '============================================================
    Private Sub CreateEmptyCells(ByVal empty As Integer)
        Dim c, r As Integer
        '----choose random locations for empty cells----
        Dim emptyCells(empty - 1) As String
        For i As Integer = 0 To (empty \ 2)
            Dim duplicate As Boolean
            Do
                duplicate = False
                '---get a cell in the first half of the grid
                Do
                    c = RandomNumber(1, 9)
                    r = RandomNumber(1, 5)
                Loop While (r = 5 And c > 5)

                For j As Integer = 0 To i
                    '---if cell is already selected to be empty
                    If emptyCells(j) = c.ToString() & r.ToString() Then
                        duplicate = True
                        Exit For
                    End If
                Next

                If Not duplicate Then
                    '---set the empty cell---
                    emptyCells(i) = c.ToString() & r.ToString()
                    actual(c, r) = 0
                    possible(c, r) = String.Empty
                    '---reflect the top half of the grid and make it symmetrical---
                    emptyCells(empty - 1 - i) = _
                       (10 - c).ToString() & (10 - r).ToString()
                    actual(10 - c, 10 - r) = 0
                    possible(10 - c, 10 - r) = String.Empty
                End If
            Loop While duplicate
        Next
    End Sub

    '============================================================
    ' Generate a new Sudoku puzzle
    '============================================================
    Private Function GenerateNewPuzzle( _
       ByVal level As Integer, _
       ByRef score As Integer) As String

        Dim c, r As Integer
        Dim str As String
        Dim numberofemptycells As Integer

        '---initialize the entire board---
        For r = 1 To 9
            For c = 1 To 9
                actual(c, r) = 0
                possible(c, r) = String.Empty
            Next
        Next

        '---clear the stacks---
        ActualStack.Clear()
        PossibleStack.Clear()

        '---populate the board with numbers by solving an empty grid---
        Try
            '---use logical methods to setup the grid first---
            If Not SolvePuzzle() Then
                '---then use brute-force---
                SolvePuzzleByBruteForce()
            End If
        Catch ex As Exception
            '---just in case an error occured, return an empty string---
            Return String.Empty
        End Try

        '---make a backup copy of the Actual array---
        actual_backup = actual.Clone()

        '---set the number of empty cells based on the level of difficulty---
        Select Case level
            Case 1 : numberofemptycells = RandomNumber(40, 45)
            Case 2 : numberofemptycells = RandomNumber(46, 49)
            Case 3 : numberofemptycells = RandomNumber(50, 53)
            Case 4 : numberofemptycells = RandomNumber(54, 58)
        End Select

        '---clear the stacks that are used in brute-force elimination ---
        ActualStack.Clear()
        PossibleStack.Clear()
        BruteForceStop = False

        '----create empty cells----
        CreateEmptyCells(numberofemptycells)

        '---convert the values in the actual array to a string---
        str = String.Empty
        For r = 1 To 9
            For c = 1 To 9
                str &= actual(c, r).ToString()
            Next
        Next

        '---verify the puzzle has only one solution---
        Dim tries As Integer = 0
        Do
            totalscore = 0
            Try
                If Not SolvePuzzle() Then
                    '---if puzzle is not solved and 
                    ' this is a level 1 to 3 puzzle---
                    If level < 4 Then
                        '---choose another pair of cells to empty---
                        VacateAnotherPairOfCells(str)
                        tries += 1
                    Else
                        '---level 4 puzzles does not guarantee single
                        ' solution and potentially need guessing---
                        SolvePuzzleByBruteForce()
                        Exit Do
                    End If
                Else
                    '---puzzle does indeed have 1 solution---
                    Exit Do
                End If
            Catch ex As Exception
                Return String.Empty
            End Try

            '---if too many tries, exit the loop---
            If tries > 50 Then
                Return String.Empty
            End If
        Loop While True
        '==================================================

        '---return the score as well as the puzzle as a string---
        score = totalscore
        Return str
    End Function

    '============================================================
    ' Vacate another pair of cells
    '============================================================
    Private Sub VacateAnotherPairOfCells(ByRef str As String)
        Dim c, r As Integer

        '---look for a pair of cells to restore---
        Do
            c = RandomNumber(1, 9)
            r = RandomNumber(1, 9)
        Loop Until str((c - 1) + (r - 1) * 9).ToString() = 0

        '---restore the value of the cell from the actual_backup array---
        str = str.Remove((c - 1) + (r - 1) * 9, 1)
        str = str.Insert((c - 1) + (r - 1) * 9, _
              actual_backup(c, r).ToString())

        '---restore the value of the symmetrical cell from 
        ' the actual_backup array---
        str = str.Remove((10 - c - 1) + (10 - r - 1) * 9, 1)
        str = str.Insert((10 - c - 1) + (10 - r - 1) * 9, _
              actual_backup(10 - c, 10 - r).ToString())

        '---look for another pair of cells to vacate---
        Do
            c = RandomNumber(1, 9)
            r = RandomNumber(1, 9)
        Loop Until str((c - 1) + (r - 1) * 9).ToString() <> 0

        '---remove the cell from the str---
        str = str.Remove((c - 1) + (r - 1) * 9, 1)
        str = str.Insert((c - 1) + (r - 1) * 9, "0")

        '---remove the symmetrical cell from the str---
        str = str.Remove((10 - c - 1) + (10 - r - 1) * 9, 1)
        str = str.Insert((10 - c - 1) + (10 - r - 1) * 9, "0")

        '---reinitialize the board---
        Dim counter As Short = 0
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                If CInt(str(counter).ToString()) <> 0 Then
                    actual(col, row) = CInt(str(counter).ToString())
                    possible(col, row) = str(counter).ToString()
                Else
                    actual(col, row) = 0
                    possible(col, row) = String.Empty
                End If
                counter += 1
            Next
        Next
    End Sub
End Class