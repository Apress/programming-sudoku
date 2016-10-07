Public Class Form1
    '---dimension of each cell in the grid---
    Const CellWidth As Integer = 32
    Const cellHeight As Integer = 32

    '---offset from the top left corner of the window---
    Const xOffset As Integer = -20
    Const yOffset As Integer = 25

    '---color for empty cell---
    Private DEFAULT_BACKCOLOR As Color = Color.White

    '---color for original puzzle values---  
    Private FIXED_FORECOLOR As Color = Color.Blue
    Private FIXED_BACKCOLOR As Color = Color.LightSteelBlue

    '---color for user inserted values---
    Private USER_FORECOLOR As Color = Color.Black
    Private USER_BACKCOLOR As Color = Color.LightYellow

    '---the number currently selected for insertion---
    Private SelectedNumber As Integer

    '---stacks to keep track of all the moves---
    Private Moves As Stack(Of String)
    Private RedoMoves As Stack(Of String)

    '---keep track of file name to save to---
    Private saveFileName As String = String.Empty

    '---used to represent the values in the grid---
    Private actual(9, 9) As Integer

    '---used to keep track of elapsed time---
    Private seconds As Integer = 0

    '---has the game started?---
    Private GameStarted As Boolean = False

    Private possible(9, 9) As String
    Private HintMode As Boolean

    '==================================================
    ' Form Load event - invoked when the form loads up
    '==================================================
    Private Sub Form1_Load( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) Handles MyBase.Load
        '---initialize the status bar---
        ToolStripStatusLabel1.Text = String.Empty
        ToolStripStatusLabel2.Text = String.Empty
        '---draw the board---
        DrawBoard()
    End Sub

    '==================================================
    ' Draws the board for the puzzle
    '==================================================
    Public Sub ClearBoard()
        '---initialize the stacks---
        Moves = New Stack(Of String)
        RedoMoves = New Stack(Of String)

        '---initialize the cells in the board---
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                SetCell(col, row, 0, 1)
            Next
        Next
    End Sub

    '==================================================
    ' Draws the cells and initialize the grid
    '==================================================
    Public Sub DrawBoard()
        '---default selected number is 1---
        ToolStripButton1.Checked = True
        SelectedNumber = 1

        '---used to store the location of the cell---
        Dim location As New Point
        '---draws the cells
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                location.X = col * (CellWidth + 1) + xOffset
                location.Y = row * (cellHeight + 1) + yOffset
                Dim lbl As New Label
                With lbl
                    .Name = col.ToString() & row.ToString()
                    .BorderStyle = BorderStyle.Fixed3D
                    .Location = location
                    .Width = CellWidth
                    .Height = cellHeight
                    .TextAlign = ContentAlignment.MiddleCenter
                    .BackColor = DEFAULT_BACKCOLOR
                    .Font = New Font(.Font, .Font.Style Or FontStyle.Bold)
                    .Tag = "1"
                    AddHandler lbl.Click, AddressOf Cell_Click
                End With
                Me.Controls.Add(lbl)
            Next
        Next
    End Sub

    '==================================================
    ' Displays a message in the Activities text box
    '==================================================
    Public Sub DisplayActivity( _
       ByVal str As String, _
       ByVal soundBeep As Boolean)
        If soundBeep Then Beep()
        txtActivities.Text &= str & Environment.NewLine
    End Sub

    '==================================================
    ' Click event for the label (cell) controls
    '==================================================
    Private Sub Cell_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs)

        '---check to see if game has even started or not---
        If Not GameStarted Then
            DisplayActivity("Click File->New to start a new" & _
            " game or File->Open to load an existing game", True)
            Return
        End If

        Dim cellLabel As Label = CType(sender, Label)

        '---if cell is not erasable then exit---
        If cellLabel.Tag.ToString() = "0" Then
            DisplayActivity("Selected cell is not empty", False)
            Return
        End If

        '---determine the col and row of the selected cell---
        Dim col As Integer = cellLabel.Name.Substring(0, 1)
        Dim row As Integer = cellLabel.Name.ToString().Substring(1, 1)

        '---If erasing a cell---
        If SelectedNumber = 0 Then
            '---if cell is empty then no need to erase---
            If actual(col, row) = 0 Then Return

            '---save the value in the array---
            SetCell(col, row, SelectedNumber, 1)
            DisplayActivity("Number erased at (" & _
            col & "," & row & ")", False)

        ElseIf cellLabel.Text = String.Empty Then
            '---else setting a value; check if move is valid---
            If Not IsMoveValid(col, row, SelectedNumber) Then
                DisplayActivity("Invalid move at (" & col & _
                "," & row & ")", False)
                Return
            End If
            '---save the value in the array---
            SetCell(col, row, SelectedNumber, 1)
            DisplayActivity("Number placed at (" & col & _
            "," & row & ")", False)

            '---saves the move into the stack---
            Moves.Push(cellLabel.Name.ToString() _
            & SelectedNumber)

            If IsPuzzleSolved() Then
                Timer1.Enabled = False
                Beep()
                ToolStripStatusLabel1.Text = "*****Puzzle Solved*****"
            End If
        End If
    End Sub

    '==================================================
    ' Event handler for the ToolStripButton controls
    '==================================================
    Private Sub ToolStripButton_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles _
       ToolStripButton1.Click, _
       ToolStripButton2.Click, _
       ToolStripButton3.Click, _
       ToolStripButton4.Click, _
       ToolStripButton5.Click, _
       ToolStripButton6.Click, _
       ToolStripButton7.Click, _
       ToolStripButton8.Click, _
       ToolStripButton9.Click, _
       ToolStripButton10.Click

        Dim selectedButton As ToolStripButton = _
           CType(sender, ToolStripButton)

        '---uncheck all the button controls in the ToolStrip---
        For i As Integer = 1 To 10
            CType(ToolStrip1.Items.Item(i), ToolStripButton).Checked = False
        Next

        '---set the selected button to "checked"---
        selectedButton.Checked = True

        '---set the appropriate number selected---
        If selectedButton.Text = "Erase" Then
            SelectedNumber = 0
        Else
            SelectedNumber = CInt(selectedButton.Text)
        End If
    End Sub

    '==================================================
    ' Set a cell to a given value
    '==================================================
    Public Sub SetCell( _
       ByVal col As Integer, ByVal row As Integer, _
       ByVal value As Integer, ByVal erasable As Short)

        '---Locate the particular Label control---
        Dim lbl() As Control = _
        Me.Controls.Find(col.ToString() & row.ToString(), True)
        Dim cellLabel As Label = CType(lbl(0), Label)

        '---save the value in the array
        actual(col, row) = value

        '---if erasing a cell, you need to reset the possible values for all cells--- 
        If value = 0 Then
            For r As Integer = 1 To 9
                For c As Integer = 1 To 9
                    If actual(c, r) = 0 Then possible(c, r) = String.Empty
                Next
            Next
        Else
            possible(col, row) = value.ToString()
        End If

        '---set the appearance for the Label control---
        If value = 0 Then '---erasing the cell---
            cellLabel.Text = String.Empty
            cellLabel.Tag = erasable
            cellLabel.BackColor = DEFAULT_BACKCOLOR
        Else
            If erasable = 0 Then '---means default puzzle values---
                cellLabel.BackColor = FIXED_BACKCOLOR
                cellLabel.ForeColor = FIXED_FORECOLOR
            Else '---means user-set value---
                cellLabel.BackColor = USER_BACKCOLOR
                cellLabel.ForeColor = USER_FORECOLOR
            End If
            cellLabel.Text = value
            cellLabel.Tag = erasable
        End If
    End Sub

    '==================================================
    ' Undo a move
    '==================================================
    Private Sub UndoToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles UndoToolStripMenuItem.Click

        '---if no previous moves, then exit---
        If Moves.Count = 0 Then Return

        '---remove from one stack and push into the redo stack---
        Dim str As String = Moves.Pop
        RedoMoves.Push(str)

        '---save the value in the array
        SetCell(Integer.Parse(str(0)), Integer.Parse(str(1)), 0, 1)
        DisplayActivity("Value removed at (" & Integer.Parse(str(0)) & _
           "," & Integer.Parse(str(1)) & ")", False)
    End Sub

    '==================================================
    ' Redo the move
    '==================================================
    Private Sub RedoToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles RedoToolStripMenuItem.Click

        '---if no more next move, then exit---
        If RedoMoves.Count = 0 Then Return

        '---remove from one stack and push into the moves stack---
        Dim str As String = RedoMoves.Pop
        Moves.Push(str)

        '---save the value in the array
        SetCell(Integer.Parse(str(0)), Integer.Parse(str(1)), Integer.Parse(str(2)), 1)
        DisplayActivity("Value reinserted at (" & _
           Integer.Parse(str(0)) & "," & Integer.Parse(str(1)) & ")", False)
    End Sub

    '==================================================
    ' Open a saved game
    '==================================================
    Private Sub OpenToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles OpenToolStripMenuItem.Click

        If GameStarted Then
            Dim response As MsgBoxResult = _
            MessageBox.Show("Do you want to save current game?", _
                            "Save current game", _
                            MessageBoxButtons.YesNoCancel, _
                            MessageBoxIcon.Question)

            If response = MsgBoxResult.Yes Then
                SaveGameToDisk(False)
            ElseIf response = MsgBoxResult.Cancel Then
                Return
            End If
        End If

        '---load the game from disk---
        Dim fileContents As String
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "SDO files (*.sdo)|*.sdo|All files (*.*)|*.*"
        openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = False

        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            fileContents = My.Computer.FileSystem.ReadAllText(openFileDialog1.FileName)
            ToolStripStatusLabel1.Text = openFileDialog1.FileName
            saveFileName = openFileDialog1.FileName
        Else
            Return
        End If

        StartNewGame()

        '---initialize the board---
        Dim counter As Short = 0
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                Try
                    If CInt(fileContents(counter).ToString()) <> 0 Then
                        SetCell(col, row, CInt(fileContents(counter).ToString()), 0)
                    End If
                Catch ex As Exception
                    MsgBox("File does not contain a valid Sudoku puzzle")
                    Exit Sub
                End Try
                counter += 1
            Next
        Next
    End Sub

    '==================================================
    ' Start a new game by resetting some of the values
    '==================================================
    Public Sub StartNewGame()
        saveFileName = String.Empty
        txtActivities.Text = String.Empty
        seconds = 0
        ClearBoard()
        GameStarted = True
        Timer1.Enabled = True
        ToolStripStatusLabel1.Text = "New game started"

        '---chap3---
        ToolTip1.RemoveAll()
    End Sub

    '==================================================
    ' Start a new game
    '==================================================
    Private Sub NewToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles NewToolStripMenuItem.Click

        If GameStarted Then
            Dim response As MsgBoxResult = _
               MessageBox.Show("Do you want to save current game?", _
                               "Save current game", _
                               MessageBoxButtons.YesNoCancel, _
                               MessageBoxIcon.Question)

            If response = MsgBoxResult.Yes Then
                SaveGameToDisk(False)
            ElseIf response = MsgBoxResult.Cancel Then
                Return
            End If
        End If

        '---Change to the hourglass cursor---
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        ToolStripStatusLabel1.Text = "Generating new puzzle..."

        '---create an instance of the SudokuPuzzle class---
        Dim sp As New SudokuPuzzle
        Dim puzzle As String = String.Empty

        '---determine the correct level---
        If EasyToolStripMenuItem.Checked Then
            puzzle = sp.GetPuzzle(1)
        ElseIf MediumToolStripMenuItem.Checked Then
            puzzle = sp.GetPuzzle(2)
        ElseIf DifficultToolStripMenuItem.Checked Then
            puzzle = sp.GetPuzzle(3)
        ElseIf ExtremelyDifficultToolStripMenuItem.Checked Then
            puzzle = sp.GetPuzzle(4)
        End If

        '---Change back to the default cursor
        Windows.Forms.Cursor.Current = Cursors.Default

        StartNewGame()

        '---initialize the board---
        Dim counter As Integer = 0
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                If puzzle(counter).ToString() <> "0" Then
                    SetCell(col, row, CInt(puzzle(counter).ToString()), 0)
                End If
                counter += 1
            Next
        Next
    End Sub

    '==================================================
    ' Draw the lines outlining the mini-grids
    '==================================================
    Private Sub Form1_Paint( _
       ByVal sender As Object, _
       ByVal e As System.Windows.Forms.PaintEventArgs) _
       Handles Me.Paint

        Dim x1, y1, x2, y2 As Integer
        '---draw the horizontal lines---
        x1 = 1 * (CellWidth + 1) + xOffset - 1
        x2 = 9 * (CellWidth + 1) + xOffset + CellWidth
        For r As Integer = 1 To 10 Step 3
            y1 = r * (cellHeight + 1) + yOffset - 1
            y2 = y1
            e.Graphics.DrawLine(Pens.Black, x1, y1, x2, y2)
        Next
        '---draw the vertical lines---
        y1 = 1 * (cellHeight + 1) + yOffset - 1
        y2 = 9 * (cellHeight + 1) + yOffset + cellHeight
        For c As Integer = 1 To 10 Step 3
            x1 = c * (CellWidth + 1) + xOffset - 1
            x2 = x1
            e.Graphics.DrawLine(Pens.Black, x1, y1, x2, y2)
        Next
    End Sub

    '==================================================
    ' Save the game to disk
    '==================================================
    Public Sub SaveGameToDisk(ByVal saveAs As Boolean)
        '---if saveFileName is empty, means game has not been saved before
        If saveFileName = String.Empty OrElse saveAs Then
            Dim saveFileDialog1 As New SaveFileDialog
            saveFileDialog1.Filter = "SDO files (*.sdo)|*.sdo|All files (*.*)|*.*"
            saveFileDialog1.FilterIndex = 1
            saveFileDialog1.RestoreDirectory = False
            If saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                '---store the filename first---
                saveFileName = saveFileDialog1.FileName
            Else
                Return
            End If
        End If

        '---formulate the string representing the values to store---
        Dim str As New System.Text.StringBuilder
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                str.Append(actual(col, row).ToString())
            Next
        Next

        '---save the values to file---
        Try
            Dim fileExists As Boolean
            fileExists = My.Computer.FileSystem.FileExists(saveFileName)
            If fileExists Then My.Computer.FileSystem.DeleteFile(saveFileName)
            My.Computer.FileSystem.WriteAllText(saveFileName, str.ToString(), True)
            ToolStripStatusLabel1.Text = "Puzzle saved in " & saveFileName
        Catch ex As Exception
            MsgBox("Error saving game. Please try again.")
        End Try
    End Sub

    '==================================================
    ' Save as...
    '==================================================
    Private Sub SaveAsToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles SaveAsToolStripMenuItem.Click

        If Not GameStarted Then
            DisplayActivity("Game not started yet.", True)
            Return
        End If

        SaveGameToDisk(True)
    End Sub

    '==================================================
    ' Debugging Use. For printing out the possible values in empty grids. NOT covered in book.
    '==================================================
    Public Sub checkvalues()
        '----print results
        Console.WriteLine("==============================================================")
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                Dim length As Integer
                If possible(col, row) IsNot Nothing Then
                    length = 8 - possible(col, row).Length
                Else
                    length = 1
                End If
                Dim space As String = String.Empty
                For j As Integer = 1 To length
                    space += " "
                Next
                If possible(col, row) IsNot Nothing Then
                    Console.Write(possible(col, row) & vbTab)
                Else
                    Console.Write(actual(col, row) & vbTab)
                End If
                'Console.Write(actual(col, row) & space)
            Next
            Console.WriteLine()
            Console.WriteLine("---------------------------------------------------------------------------")
            Console.WriteLine()
        Next
        Console.WriteLine("==============================================================")
        '-------------------
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        checkvalues()
    End Sub

    '==================================================
    ' Check if move is valid
    '==================================================
    Public Function IsMoveValid( _
       ByVal col As Integer, _
       ByVal row As Integer, _
       ByVal value As Integer) As Boolean

        Dim puzzleSolved As Boolean = True

        '---scan through column
        For r As Integer = 1 To 9
            If actual(col, r) = value Then 'duplicate
                Return False
            End If
        Next

        '---scan through row
        For c As Integer = 1 To 9
            If actual(c, row) = value Then
                Return False
            End If
        Next

        '---scan through mini-grid
        Dim startC, startR As Integer
        startC = col - ((col - 1) Mod 3)
        startR = row - ((row - 1) Mod 3)

        For rr As Integer = 0 To 2
            For cc As Integer = 0 To 2
                If actual(startC + cc, startR + rr) = value Then 'duplicate
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    '==================================================
    ' Increment the time counter
    '==================================================
    Private Sub Timer1_Tick( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel2.Text = "Elapsed time: " & seconds & " second(s)"
        seconds += 1
    End Sub

    '==================================================
    ' Save the game
    '==================================================
    Private Sub SaveToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles SaveToolStripMenuItem.Click
        If Not GameStarted Then
            DisplayActivity("Game not started yet.", True)
            Return
        End If
        SaveGameToDisk(False)
    End Sub

    '==================================================
    ' Exit the application
    '==================================================
    Private Sub ExitToolStripMenuItem_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles ExitToolStripMenuItem.Click
        If GameStarted Then
            Dim response As MsgBoxResult = _
            MsgBox("Do you want to save current game?", _
            MsgBoxStyle.YesNoCancel, "Save current game")

            If response = MsgBoxResult.Yes Then
                SaveGameToDisk(False)
            ElseIf response = MsgBoxResult.Cancel Then
                Return
            End If
        End If
        '---exit the application---
        End
    End Sub

    '==================================================
    ' Calculates the possible values for all the cell
    '==================================================
    Public Function CheckColumnsAndRows() As Boolean
        Dim changes As Boolean = False
        '---check all cells
        For row As Integer = 1 To 9
            For col As Integer = 1 To 9
                If actual(col, row) = 0 Then
                    Try
                        possible(col, row) = CalculatePossibleValues(col, row)
                    Catch ex As Exception
                        DisplayActivity("Invalid placement, please undo move", False)
                        Throw New Exception("Invalid Move")
                    End Try

                    '---display the possible values in the Tooltip
                    SetToolTip(col, row, possible(col, row))

                    If possible(col, row).Length = 1 Then
                        '---that means a number is found---
                        SetCell(col, row, CInt(possible(col, row)), 1)

                        '----Number is confirmed
                        actual(col, row) = CInt(possible(col, row))
                        DisplayActivity("Col/Row and Minigrid Elimination", False)
                        DisplayActivity("==========================", False)
                        DisplayActivity("Inserted value " & actual(col, row) & _
                                        " in " & "(" & col & "," & row & ")", False)
                        Application.DoEvents()
                        '---saves the move into the stack
                        Moves.Push(col & row & possible(col, row))
                        '---if user only ask for a hint, stop at this point---
                        changes = True
                        If HintMode Then Return True
                    End If
                End If
            Next
        Next
        Return changes
    End Function

    '==================================================
    ' Calculates the possible values for a cell
    '==================================================
    Public Function CalculatePossibleValues( _
                    ByVal col As Integer, _
                    ByVal row As Integer) _
                    As String
        '---get the current possible values for the cell
        Dim str As String
        If possible(col, row) = String.Empty Then
            str = "123456789"
        Else
            str = possible(col, row)
        End If

        '---Step (1) check by column
        Dim r, c As Integer
        For r = 1 To 9
            If actual(col, r) <> 0 Then
                'that means there is a actual value in it
                str = str.Replace(actual(col, r).ToString(), String.Empty)
            End If
        Next

        '---Step (2) check by row
        For c = 1 To 9
            If actual(c, row) <> 0 Then
                'that means there is a actual value in it
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
                    'that means there is a actual value in it
                    str = str.Replace(actual(cc, rr).ToString(), String.Empty)
                End If
            Next
        Next

        '---if possible value is string.Empty, then error
        If str = String.Empty Then
            Throw New Exception("Invalid Move")
        End If
        Return str
    End Function

    '==================================================
    ' Set the Tooltip for a Label control
    '==================================================
    Public Sub SetToolTip( _
           ByVal col As Integer, ByVal row As Integer, _
           ByVal possiblevalues As String)

        '---Locate the particular Label control---
        Dim lbl() As Control = _
        Me.Controls.Find(col.ToString() & row.ToString(), True)
        ToolTip1.SetToolTip(CType(lbl(0), Label), possiblevalues)
    End Sub

    '==================================================
    ' Solve button
    '==================================================
    Private Sub btnSolvePuzzle_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnSolvePuzzle.Click

        ActualStack.Clear()
        PossibleStack.Clear()
        BruteForceStop = False
        '---solve the puzzle; no need to stop
        HintMode = False
        Try
            If Not SolvePuzzle() Then
                SolvePuzzleByBruteForce()
            End If
        Catch ex As Exception
            MsgBox("Puzzle not solvable.")
        End Try

        If Not IsPuzzleSolved() Then
            MsgBox("Puzzle Not solved.")
        End If

    End Sub

    '==================================================
    ' Hint button
    '==================================================
    Private Sub btnHint_Click( _
       ByVal sender As System.Object, _
       ByVal e As System.EventArgs) _
       Handles btnHint.Click

        '---show hints one cell at a time
        HintMode = True
        Try
            SolvePuzzle()
        Catch ex As Exception
            MessageBox.Show("Please undo your move", "Invalid Move", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    '==================================================
    ' Steps to solve the puzzle
    '==================================================
    Public Function SolvePuzzle() As Boolean
        Dim changes As Boolean
        Dim ExitLoop As Boolean = False
        Try
            Do '---Look for Triplets in Columns
                Do '---Look for Triplets in Rows
                    Do '---Look for Triplets in Minigrids
                        Do '---Look for Twins in Columns
                            Do '---Look for Twins in Rows
                                Do '---Look for Twins in Minigrids
                                    Do '---Look for Lone Ranger in Columns
                                        Do '---Look for Lone Ranger in Rows
                                            Do  '---Look for Lone Ranger in 
                                                ' Minigrids 
                                                Do '---Perform Col/Row and 
                                                    ' Minigrid Elimination
                                                    changes = _
                                                       CheckColumnsAndRows()
                                                    If (HintMode AndAlso changes) _
                                                       OrElse IsPuzzleSolved() Then
                                                        ExitLoop = True
                                                        Exit Do
                                                    End If
                                                Loop Until Not changes

                                                If ExitLoop Then Exit Do
                                                '---Look for Lone Ranger in 
                                                ' Minigrids
                                                changes = _
                                                   LookForLoneRangersinMinigrids()
                                                If (HintMode AndAlso changes) _
                                                   OrElse IsPuzzleSolved() Then
                                                    ExitLoop = True
                                                    Exit Do
                                                End If
                                            Loop Until Not changes

                                            If ExitLoop Then Exit Do
                                            '---Look for Lone Ranger in Rows
                                            changes = LookForLoneRangersinRows()
                                            If (HintMode AndAlso changes) OrElse _
                                               IsPuzzleSolved() Then
                                                ExitLoop = True
                                                Exit Do
                                            End If
                                        Loop Until Not changes

                                        If ExitLoop Then Exit Do
                                        '---Look for Lone Ranger in Columns
                                        changes = LookForLoneRangersinColumns()
                                        If (HintMode AndAlso changes) OrElse _
                                           IsPuzzleSolved() Then
                                            ExitLoop = True
                                            Exit Do
                                        End If
                                    Loop Until Not changes

                                    If ExitLoop Then Exit Do
                                    '---Look for Twins in Minigrids
                                    changes = LookForTwinsinMinigrids()
                                    If (HintMode AndAlso changes) OrElse _
                                       IsPuzzleSolved() Then
                                        ExitLoop = True
                                        Exit Do
                                    End If
                                Loop Until Not changes

                                If ExitLoop Then Exit Do
                                '---Look for Twins in Rows
                                changes = LookForTwinsinRows()
                                If (HintMode AndAlso changes) OrElse _
                                   IsPuzzleSolved() Then
                                    ExitLoop = True
                                    Exit Do
                                End If
                            Loop Until Not changes

                            If ExitLoop Then Exit Do
                            '---Look for Twins in Columns
                            changes = LookForTwinsinColumns()
                            If (HintMode AndAlso changes) OrElse _
                               IsPuzzleSolved() Then
                                ExitLoop = True
                                Exit Do
                            End If
                        Loop Until Not changes

                        If ExitLoop Then Exit Do
                        '---Look for Triplets in Minigrids
                        changes = LookForTripletsinMinigrids()
                        If (HintMode AndAlso changes) OrElse IsPuzzleSolved() Then
                            ExitLoop = True
                            Exit Do
                        End If
                    Loop Until Not changes

                    If ExitLoop Then Exit Do
                    '---Look for Triplets in Rows
                    changes = LookForTripletsinRows()
                    If (HintMode AndAlso changes) OrElse IsPuzzleSolved() Then
                        ExitLoop = True
                        Exit Do
                    End If
                Loop Until Not changes

                If ExitLoop Then Exit Do
                '---Look for Triplets in Columns
                changes = LookForTripletsinColumns()
                If (HintMode AndAlso changes) OrElse IsPuzzleSolved() Then
                    ExitLoop = True
                    Exit Do
                End If
            Loop Until Not changes

        Catch ex As Exception
            Throw New Exception("Invalid Move")
        End Try

        If IsPuzzleSolved() Then
            Timer1.Enabled = False
            Beep()
            ToolStripStatusLabel1.Text = "*****Puzzle Solved*****"
            MsgBox("Puzzle solved")
            Return True
        Else
            Return False
        End If
    End Function

    '==================================================
    ' Look for lone rangers in Minigrids
    '==================================================
    Public Function LookForLoneRangersinMinigrids() As Boolean
        Dim changes As Boolean = False
        Dim NextMiniGrid As Boolean
        Dim occurence As Integer
        Dim cPos, rPos As Integer

        '---check for each number from 1 to 9---
        For n As Integer = 1 To 9

            '---check the 9 mini-grids---
            For r As Integer = 1 To 9 Step 3
                For c As Integer = 1 To 9 Step 3

                    NextMiniGrid = False
                    '---check within the mini-grid---
                    occurence = 0
                    For rr As Integer = 0 To 2
                        For cc As Integer = 0 To 2
                            If actual(c + cc, r + rr) = 0 AndAlso _
                               possible(c + cc, r + rr).Contains(n.ToString()) Then
                                occurence += 1
                                cPos = c + cc
                                rPos = r + rr
                                If occurence > 1 Then
                                    NextMiniGrid = True
                                    Exit For
                                End If
                            End If
                        Next
                        If NextMiniGrid Then Exit For
                    Next
                    If (Not NextMiniGrid) AndAlso occurence = 1 Then
                        '---that means number is confirmed---
                        SetCell(cPos, rPos, n, 1)
                        SetToolTip(cPos, rPos, n.ToString())
                        '---saves the move into the stack
                        Moves.Push(cPos & rPos & n.ToString())
                        DisplayActivity("Look for Lone Rangers in Minigrids", False)
                        DisplayActivity("===========================", False)
                        DisplayActivity("Inserted value " & n.ToString() & _
                                        " in " & "(" & cPos & "," & rPos & ")", False)
                        Application.DoEvents()
                        changes = True
                        '---if user clicks the Hint button, exit the function---
                        If HintMode Then Return True
                    End If
                Next
            Next
        Next
        Return changes
    End Function

    '=========================================================
    'Look for Lone Rangers in Rows
    '=========================================================
    Public Function LookForLoneRangersinRows() As Boolean
        Dim changes As Boolean = False
        Dim occurence As Integer
        Dim cPos, rPos As Integer

        '---check by row----
        For r As Integer = 1 To 9
            For n As Integer = 1 To 9
                occurence = 0
                For c As Integer = 1 To 9
                    If actual(c, r) = 0 AndAlso possible(c, r).Contains(n.ToString()) Then
                        occurence += 1
                        '---if multiple occurence, not a lone ranger anymore
                        If occurence > 1 Then Exit For
                        cPos = c
                        rPos = r
                    End If
                Next
                If occurence = 1 Then
                    '--number is confirmed---
                    SetCell(cPos, rPos, n, 1)
                    SetToolTip(cPos, rPos, n.ToString())
                    '---saves the move into the stack---
                    Moves.Push(cPos & rPos & n.ToString())
                    DisplayActivity("Look for Lone Rangers in Rows", False)
                    DisplayActivity("=========================", False)
                    DisplayActivity("Inserted value " & n.ToString() & _
                                    " in " & "(" & cPos & "," & rPos & ")", False)
                    Application.DoEvents()
                    changes = True
                    '---if user clicks the Hint button, exit the function---
                    If HintMode Then Return True
                End If
            Next
        Next
        Return changes
    End Function

    '=========================================================
    'Look for Lone Rangers in Columns
    '=========================================================
    Public Function LookForLoneRangersinColumns() As Boolean
        Dim changes As Boolean = False
        Dim occurence As Integer
        Dim cPos, rPos As Integer

        '----check by column----
        For c As Integer = 1 To 9
            For n As Integer = 1 To 9
                occurence = 0
                For r As Integer = 1 To 9
                    If actual(c, r) = 0 AndAlso possible(c, r).Contains(n.ToString()) Then
                        occurence += 1
                        '---if multiple occurence, not a lone ranger anymore
                        If occurence > 1 Then Exit For
                        cPos = c
                        rPos = r
                    End If
                Next
                If occurence = 1 Then
                    '--number is confirmed---
                    SetCell(cPos, rPos, n, 1)
                    SetToolTip(cPos, rPos, n.ToString())
                    '---saves the move into the stack
                    Moves.Push(cPos & rPos & n.ToString())
                    DisplayActivity("Look for Lone Rangers in Columns", False)
                    DisplayActivity("===========================", False)
                    DisplayActivity("Inserted value " & n.ToString() & _
                                    " in " & "(" & cPos & "," & rPos & ")", False)
                    Application.DoEvents()
                    changes = True
                    '---if user clicks the Hint button, exit the function---
                    If HintMode Then Return True
                End If
            Next
        Next
        Return changes
    End Function

    '=========================================================
    'Look for Twins in Columns
    '=========================================================
    Public Function LookForTwinsinColumns() As Boolean
        Dim changes As Boolean = False

        '---for each column, check each row in the column---
        For c As Integer = 1 To 9
            For r As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '--scan rows in this column---
                    For rr As Integer = r + 1 To 9

                        If (possible(c, rr) = possible(c, r)) Then

                            '--twins found---
                            DisplayActivity("Twins found in column at: (" & _
                               c & "," & r & ") and (" & c & "," & rr & ")", False)

                            '---remove the twins from all the other possible 
                            ' values in the row---
                            For rrr As Integer = 1 To 9

                                If (actual(c, rrr) = 0) AndAlso (rrr <> r) _
                                   AndAlso (rrr <> rr) Then

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

                                    '---set the ToolTip---
                                    SetToolTip(c, rrr, possible(c, rrr))

                                    '---if the possible values are modified, then 
                                    ' set the changes variable to True to indicate 
                                    ' that the possible values of cells in the
                                    ' minigrid have been modified---
                                    If original_possible <> possible(c, rrr) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string,
                                    ' then the user has placed a move that results
                                    ' in the puzzle being not solvable---
                                    If possible(c, rrr) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(c, rrr).Length = 1 Then
                                        SetCell(c, rrr, CInt(possible(c, rrr)), 1)
                                        SetToolTip(c, rrr, possible(c, rrr))
                                        '---saves the move into the stack
                                        Moves.Push(c & rrr & possible(c, rrr))
                                        DisplayActivity( _
                                           "Looking for twins (by column)", False)
                                        DisplayActivity( _
                                           "============================", False)
                                        DisplayActivity( _
                                           "Inserted value " & actual(c, rrr) & _
                                           " in " & "(" & c & "," & rrr & ")", _
                                           False)
                                        Application.DoEvents()

                                        '---if user clicks the Hint button,
                                        'exit the function---
                                        If HintMode Then Return True
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

    '=========================================================
    'Look for Twins in Minigrids
    '=========================================================
    Public Function LookForTwinsinMinigrids() As Boolean
        Dim changes As Boolean = False

        '---look for twins in each cell---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '---scan by the mini-grid that the current cell is in
                    Dim startC, startR As Integer
                    startC = c - ((c - 1) Mod 3)
                    startR = r - ((r - 1) Mod 3)

                    For rr As Integer = startR To startR + 2
                        For cc As Integer = startC To startC + 2

                            '---for cells other than the pair of twins---
                            If (Not ((cc = c) AndAlso (rr = r))) AndAlso _
                               possible(cc, rr) = possible(c, r) Then

                                '---twins found---
                                DisplayActivity("Twins found in minigrid at: (" & _
                                                c & "," & r & ") and (" & _
                                                cc & "," & rr & ")", False)

                                '---remove the twins from all the other possible 
                                ' values in the minigrid---
                                For rrr As Integer = startR To startR + 2
                                    For ccc As Integer = startC To startC + 2

                                        '---only check for empty cells---
                                        If actual(ccc, rrr) = 0 AndAlso _
                                           possible(ccc, rrr) <> possible(c, r) _
                                           Then

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

                                            '---set the ToolTip---
                                            SetToolTip( _
                                               ccc, rrr, possible(ccc, rrr))

                                            '---if the possible values are 
                                            ' modified, then set the changes 
                                            ' variable to True to indicate that 
                                            ' the possible values of cells in the 
                                            ' minigrid have been modified---
                                            If original_possible <> _
                                               possible(ccc, rrr) Then
                                                changes = True
                                            End If

                                            '---if possible value reduces to empty 
                                            ' string, then the user has placed a 
                                            ' move that results in the puzzle being
                                            ' not solvable---
                                            If possible(ccc, rrr) = String.Empty _
                                               Then
                                                Throw New Exception("Invalid Move")
                                            End If

                                            '---if left with 1 possible value for 
                                            ' the current cell, cell is 
                                            ' confirmed---
                                            If possible(ccc, rrr).Length = 1 Then
                                                SetCell(ccc, rrr, _
                                                   CInt(possible(ccc, rrr)), 1)
                                                SetToolTip( _
                                                   ccc, rrr, possible(ccc, rrr))
                                                '---saves the move into the stack
                                                Moves.Push( _
                                                   ccc & rrr & possible(ccc, rrr))
                                                DisplayActivity( _
                                                   "Look For Twins in Minigrids", _
                                                   False)
                                                DisplayActivity( _
                                                   "===========================", _
                                                   False)
                                                DisplayActivity( _
                                                   "Inserted value " & _
                                                   actual(ccc, rrr) & _
                                                   " in " & "(" & ccc & "," & _
                                                   rrr & ")", False)
                                                Application.DoEvents()
                                                '---if user clicks the Hint button,
                                                'exit the function---
                                                If HintMode Then Return True
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

    '=========================================================
    'Look for Twins in Rows
    '=========================================================
    Public Function LookForTwinsinRows() As Boolean
        Dim changes As Boolean = False

        '---for each row, check each column in the row---
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '---if two possible values, check for twins---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 2 Then

                    '--scan columns in this row---
                    For cc As Integer = c + 1 To 9
                        If (possible(cc, r) = possible(c, r)) Then

                            '--twins found---
                            DisplayActivity("Twins found in row at: (" & _
                               c & "," & r & ") and (" & cc & "," & r & ")", _
                               False)

                            '---remove the twins from all the other possible 
                            ' values in the row---
                            For ccc As Integer = 1 To 9
                                If (actual(ccc, r) = 0) AndAlso (ccc <> c) _
                                   AndAlso (ccc <> cc) Then

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

                                    '---set the ToolTip---
                                    SetToolTip(ccc, r, possible(ccc, r))

                                    '---if the possible values are modified, then 
                                    ' set the changes variable to True to indicate 
                                    ' that the possible values of cells in the 
                                    ' minigrid have been modified---
                                    If original_possible <> possible(ccc, r) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string, 
                                    ' then the user has placed a move that results 
                                    ' in the puzzle being not solvable---
                                    If possible(ccc, r) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(ccc, r).Length = 1 Then
                                        SetCell(ccc, r, CInt(possible(ccc, r)), 1)
                                        SetToolTip(ccc, r, possible(ccc, r))

                                        '---saves the move into the stack
                                        Moves.Push(ccc & r & possible(ccc, r))
                                        DisplayActivity("Look For Twins in Rows)", _
                                           False)
                                        DisplayActivity("=======================", _
                                           False)
                                        DisplayActivity("Inserted value " & _
                                           actual(ccc, r) & " in " & "(" & _
                                           ccc & "," & r & ")", False)
                                        Application.DoEvents()

                                        '---if user clicks the Hint button, exit 
                                        ' the function---
                                        If HintMode Then Return True
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

    '=========================================================
    'Look for Triplets in Minigrids
    '=========================================================
    Public Function LookForTripletsinMinigrids() As Boolean
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

                                '---save the coordinates of the triplets
                                tripletsLocation &= cc.ToString() & rr.ToString()
                            End If
                        Next
                    Next

                    '--found 3 cells as triplets; remove all from the other
                    ' cells---
                    If tripletsLocation.Length = 6 Then

                        '--triplets found---
                        DisplayActivity("Triplets found in " & _
                           tripletsLocation, False)

                        '---remove each cell's possible values containing the
                        ' triplet---
                        For rrr As Integer = startR To startR + 2
                            For ccc As Integer = startC To startC + 2

                                '---look for the cell that is not part of the 3 
                                ' cells found---
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

                                    '---set the ToolTip---
                                    SetToolTip(ccc, rrr, possible(ccc, rrr))

                                    '---if the possible values are modified, then 
                                    ' set the changes variable to True to indicate 
                                    ' that the possible values of cells in the
                                    ' minigrid have been modified---
                                    If original_possible <> possible(ccc, rrr) Then
                                        changes = True
                                    End If

                                    '---if possible value reduces to empty string, 
                                    'then the user has placed a move that results
                                    ' in the puzzle being not solvable---
                                    If possible(ccc, rrr) = String.Empty Then
                                        Throw New Exception("Invalid Move")
                                    End If

                                    '---if left with 1 possible value for the 
                                    ' current cell, cell is confirmed---
                                    If possible(ccc, rrr).Length = 1 Then
                                        SetCell(ccc, rrr, _
                                           CInt(possible(ccc, rrr)), 1)
                                        SetToolTip(ccc, rrr, possible(ccc, rrr))
                                        '---saves the move into the stack
                                        Moves.Push(ccc & rrr & possible(ccc, rrr))
                                        DisplayActivity( _
                                           "Look For Triplets in Minigrids)", _
                                           False)
                                        DisplayActivity( _
                                           "==============================", False)
                                        DisplayActivity( _
                                           "Inserted value " & actual(ccc, rrr) & _
                                           " in " & "(" & ccc & "," & rrr & ")", _
                                           False)
                                        Application.DoEvents()

                                        '---if user clicks the Hint button, exit 
                                        ' the function---
                                        If HintMode Then Return True
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

    '=========================================================
    'Look for Triplets in Columns
    '=========================================================
    Public Function LookForTripletsinColumns() As Boolean
        Dim changes As Boolean = False

        '---for each column, check each row in the column
        For c As Integer = 1 To 9
            For r As Integer = 1 To 9

                '--- three possible values; check for triplets---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 3 Then

                    '---first potential triplet found---
                    Dim tripletsLocation As String = c.ToString() & r.ToString()

                    '--scans rows in this column---
                    For rr As Integer = 1 To 9
                        If (rr <> r) AndAlso _
                           ((possible(c, rr) = possible(c, r)) OrElse _
                           (possible(c, rr).Length = 2 AndAlso _
                            possible(c, r).Contains( _
                            possible(c, rr)(0).ToString()) AndAlso _
                            possible(c, r).Contains( _
                            possible(c, rr)(1).ToString()))) Then

                            '---save the coordinates of the triplet---
                            tripletsLocation += c.ToString() & rr.ToString()
                        End If
                    Next

                    '--found 3 cells as triplets; remove all from the other 
                    ' cells---
                    If tripletsLocation.Length = 6 Then

                        '--triplets found---
                        DisplayActivity("Triplets found in " & tripletsLocation, _
                           False)

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

                                '---set the ToolTip---
                                SetToolTip(c, rrr, possible(c, rrr))

                                '---if the possible values are modified, then set 
                                ' the changes variable to True to indicate that 
                                ' the possible values of cells in the minigrid 
                                ' have been modified---
                                If original_possible <> possible(c, rrr) Then
                                    changes = True
                                End If

                                '---if possible value reduces to empty string, then
                                ' the user has placed a move that results in the 
                                ' puzzle beingnot solvable---
                                If possible(c, rrr) = String.Empty Then
                                    Throw New Exception("Invalid Move")
                                End If

                                '---if left with 1 possible value for the current 
                                ' cell, cell is confirmed---
                                If possible(c, rrr).Length = 1 Then
                                    SetCell(c, rrr, CInt(possible(c, rrr)), 1)
                                    SetToolTip(c, rrr, possible(c, rrr))

                                    '---saves the move into the stack
                                    Moves.Push(c & rrr & possible(c, rrr))
                                    DisplayActivity( _
                                       "Look For Triplets in Columns)", False)
                                    DisplayActivity( _
                                       "=============================", False)
                                    DisplayActivity( _
                                       "Inserted value " & actual(c, rrr) & _
                                       " in " & "(" & c & "," & rrr & ")", False)
                                    Application.DoEvents()

                                    '---if user clicks the Hint button, exit the 
                                    ' function---
                                    If HintMode Then Return True
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
    'Look for Triplets in Rows
    '=========================================================
    Public Function LookForTripletsinRows() As Boolean
        Dim changes As Boolean = False

        '---for each row, check each column in the row
        For r As Integer = 1 To 9
            For c As Integer = 1 To 9

                '--- three possible values; check for triplets---
                If actual(c, r) = 0 AndAlso possible(c, r).Length = 3 Then

                    '---first potential triplet found---
                    Dim tripletsLocation As String = c.ToString() & r.ToString()

                    '--scans columns in this row---
                    For cc As Integer = 1 To 9
                        '---look for other triplets---
                        If (cc <> c) AndAlso _
                           ((possible(cc, r) = possible(c, r)) OrElse _
                            (possible(cc, r).Length = 2 AndAlso _
                             possible(c, r).Contains( _
                             possible(cc, r)(0).ToString()) AndAlso _
                             possible(c, r).Contains( _
                             possible(cc, r)(1).ToString()))) Then
                            '---save the coordinates of the triplet---
                            tripletsLocation &= cc.ToString() & r.ToString()
                        End If
                    Next

                    '--found 3 cells as triplets; remove all from the other 
                    ' cells---
                    If tripletsLocation.Length = 6 Then

                        '--triplets found---
                        DisplayActivity("Triplets found in " & tripletsLocation, _
                           False)

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
                                   possible(ccc, r).Replace(possible(c, r)(0), _
                                   String.Empty)

                                '---remove second triplet number from possible
                                ' values---
                                possible(ccc, r) = _
                                   possible(ccc, r).Replace(possible(c, r)(1), _
                                   String.Empty)

                                '---remove third triplet number from possible
                                ' values---
                                possible(ccc, r) = _
                                   possible(ccc, r).Replace(possible(c, r)(2), _
                                   String.Empty)

                                '---set the ToolTip---
                                SetToolTip(ccc, r, possible(ccc, r))

                                '---if the possible values are modified, then set 
                                ' the changes variable to True to indicate that the
                                ' possible values of cells in the minigrid have 
                                ' been modified---
                                If original_possible <> possible(ccc, r) Then
                                    changes = True
                                End If

                                '---if possible value reduces to empty string, then
                                ' the user has placed a move that results in the 
                                ' puzzle being not solvable---
                                If possible(ccc, r) = String.Empty Then
                                    Throw New Exception("Invalid Move")
                                End If

                                '---if left with 1 possible value for the current 
                                ' cell, cell is confirmed---
                                If possible(ccc, r).Length = 1 Then
                                    SetCell(ccc, r, CInt(possible(ccc, r)), 1)
                                    SetToolTip(ccc, r, possible(ccc, r))

                                    '---saves the move into the stack---
                                    Moves.Push(ccc & r & possible(ccc, r))
                                    DisplayActivity("Look For Triplets in Rows", _
                                       False)
                                    DisplayActivity("=========================", _
                                       False)
                                    DisplayActivity("Inserted value " & _
                                       actual(ccc, r) & " in " & "(" & _
                                       ccc & "," & r & ")", False)
                                    Application.DoEvents()

                                    '---if user clicks the Hint button, exit the 
                                    ' function---
                                    If HintMode Then Return True
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
    Public Sub FindCellWithFewestPossibleValues(ByRef col As Integer, ByRef row As Integer)
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

    '=========================================================
    ' Randomly swap the list of possible values
    '=========================================================
    Public Sub RandomizePossibleValues(ByRef str As String)
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

    '---used for brute-force---
    Private BruteForceStop As Boolean = False
    Private ActualStack As New Stack(Of Integer(,))()
    Private PossibleStack As New Stack(Of String(,))()

    Public Sub SolvePuzzleByBruteForce()
        Dim c, r As Integer

        '---find out which cell has the smallest number of possible values---
        FindCellWithFewestPossibleValues(c, r)

        '---get the possible values for the chosen cell
        Dim possibleValues As String = possible(c, r)

        '---comment out this section if solving systematically---
        '---randomize the possible values----
        ''RandomizePossibleValues(possibleValues)
        '-------------------

        '---push the actual and possible stacks into the stack---
        ActualStack.Push(CType(actual.Clone(), Integer(,)))
        PossibleStack.Push(CType(possible.Clone(), String(,)))

        '---select one value and try---
        For i As Integer = 0 To possibleValues.Length - 1

            '---saves the move into the stack---
            Moves.Push(c & r & possibleValues(i).ToString())
            SetCell(c, r, CInt(possibleValues(i).ToString()), 1)
            DisplayActivity("Solve Puzzle By Brute Force", False)
            DisplayActivity("=======================", False)
            DisplayActivity("Trying to insert value " & actual(c, r) & _
                            " in " & "(" & c & "," & r & ")", False)
            Try
                If SolvePuzzle() Then
                    '---if the puzzle is solved, the recursion can stop now---
                    BruteForceStop = True
                    Return
                Else
                    '---no problem with current selection, proceed with next cell---
                    SolvePuzzleByBruteForce()
                    If BruteForceStop Then Return
                End If
            Catch ex As Exception
                DisplayActivity("Invalid move; Backtracking...", False)
                actual = ActualStack.Pop()
                possible = PossibleStack.Pop()
            End Try
        Next
    End Sub

    '=========================================================
    ' Check if the puzzle is solved
    '=========================================================
    Public Function IsPuzzleSolved() As Boolean
        '---check row by row
        Dim pattern As String
        Dim r, c As Integer
        For r = 1 To 9
            pattern = "123456789"
            For c = 1 To 9
                pattern = pattern.Replace(actual(c, r).ToString(), String.Empty)
            Next
            If pattern.Length > 0 Then
                Return False
            End If
        Next

        '---check col by col
        For c = 1 To 9
            pattern = "123456789"
            For r = 1 To 9
                pattern = pattern.Replace(actual(c, r).ToString(), String.Empty)
            Next
            If pattern.Length > 0 Then
                Return False
            End If
        Next

        '---check by mini-grid
        For c = 1 To 9 Step 3
            pattern = "123456789"
            For r = 1 To 9 Step 3
                For cc As Integer = 0 To 2
                    For rr As Integer = 0 To 2
                        pattern = pattern.Replace(actual(c + cc, r + rr).ToString(), String.Empty)
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
    ' The Easy menu item
    '=========================================================
    Private Sub EasyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EasyToolStripMenuItem.Click
        MediumToolStripMenuItem.Checked = False
        DifficultToolStripMenuItem.Checked = False
        ExtremelyDifficultToolStripMenuItem.Checked = False
    End Sub

    '=========================================================
    ' The Medium menu item
    '=========================================================
    Private Sub MediumToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MediumToolStripMenuItem.Click
        EasyToolStripMenuItem.Checked = False
        DifficultToolStripMenuItem.Checked = False
        ExtremelyDifficultToolStripMenuItem.Checked = False
    End Sub

    '=========================================================
    ' The Difficult menu item
    '=========================================================
    Private Sub DifficultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DifficultToolStripMenuItem.Click
        EasyToolStripMenuItem.Checked = False
        MediumToolStripMenuItem.Checked = False
        ExtremelyDifficultToolStripMenuItem.Checked = False
    End Sub

    '=========================================================
    ' The Extremely Difficult menu item
    '=========================================================
    Private Sub ExtremelyDifficultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtremelyDifficultToolStripMenuItem.Click
        EasyToolStripMenuItem.Checked = False
        MediumToolStripMenuItem.Checked = False
        DifficultToolStripMenuItem.Checked = False
    End Sub
End Class