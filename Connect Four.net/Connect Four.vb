'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Drawing
Imports System.Environment
Imports System.Math
Imports System.Text
Imports System.Windows.Forms

'This module contains this program's core procedures.
Public Module ConnectFourModule
   'This enumeration lists the disk colors used in the game.
   Public Enum DiskColorsE As Integer
      DCOutsideField          'Indicates that the referred location is outside the game's playing field.
      DCNone                  'Indicates that there is no disk.
      DCRed                   'Indicates a red disk.
      DCYellow                'Indicates a yellow disk.
   End Enum

   'This enumeration lists the states the game can be in.
   Private Enum GameStatesE As Integer
      GSNeitherPlaying        'Indicates that neither player is playing.
      GSRedPlaying            'Indicates that the player with the red disks is playing.
      GSRedWon                'Indicates that the player with the red disks has won.
      GSTied                  'Indicates that no more moves are possible and neither player has won.
      GSYellowPlaying         'Indicates that the player with the yellow disks is playing.
      GSYellowWon             'Indicates that the player with the yellow disks has won.
   End Enum

   'This structure defines the players' setup.
   Public Structure PlayersSetupStr
      Public ComputerColor As DiskColorsE     'The color of the disks the computer player plays with.
      Public FirstColor As DiskColorsE        'The color of the disks of the player who makes the first move.
      Public HumanColor As DiskColorsE        'The color of the disks the human player plays with when the computer player is enabled.
   End Structure

   Public Const SLOT_SIZE As Integer = 100       'The size of a disk slot in pixels.
   Private Const FIRST_COLUMN As Integer = 0      'The first column of disks.
   Private Const FIRST_ROW As Integer = 0         'The first row of disks.
   Private Const LAST_COLUMN As Integer = 6       'The last column of disks.
   Private Const LAST_ROW As Integer = 5          'The last row of disks.
   Private Const NO_COLUMN As Integer = -1        'Indicates that no column has been selected.
   Private Const WINNING_LENGTH As Integer = 4    'The number of disks of the same color that must be in one line to win.

   Private WithEvents ComputerPlayer As New Timer With {.Enabled = False, .Interval = 100}   'This lets the computer make a move when this game is played against the computer.
   Private WithEvents DiskDropper As New Timer With {.Enabled = False, .Interval = 100}      'This lets a player's disk drop into the slot selected by the active player.

   'This procedure manages the active player.
   Private Function ActivePlayerColor(Optional NewPlayer As DiskColorsE = DiskColorsE.DCNone, Optional ChangeTurns As Boolean = False, Optional ResetPlayers As Boolean = False) As DiskColorsE
      Try
         Static CurrentPlayer As DiskColorsE = DiskColorsE.DCNone

         If ChangeTurns Then
            Select Case CurrentPlayer
               Case DiskColorsE.DCRed
                  CurrentPlayer = DiskColorsE.DCYellow
               Case DiskColorsE.DCYellow
                  CurrentPlayer = DiskColorsE.DCRed
            End Select
         ElseIf Not NewPlayer = DiskColorsE.DCNone Then
            CurrentPlayer = NewPlayer
         ElseIf ResetPlayers Then
            CurrentPlayer = DiskColorsE.DCNone
         End If

         Return CurrentPlayer
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure determines which moves the computer player will make.
   Private Sub ComputerMakeMove()
      Try
         Dim Column As New Integer
         Dim MovesFound() As Boolean = FindMoves(PlayersSetup().ComputerColor, WINNING_LENGTH, AllowHelpingOpponent:=True)
         Dim Selector As New Random

         If Not FoundMoves(MovesFound) Then
            MovesFound = FindMoves(PlayersSetup().HumanColor, WINNING_LENGTH, AllowHelpingOpponent:=False)
            If Not FoundMoves(MovesFound) Then
               For TriggerLength As Integer = WINNING_LENGTH To 0 Step -1
                  MovesFound = FindMoves(PlayersSetup().ComputerColor, TriggerLength, AllowHelpingOpponent:=False)
                  If FoundMoves(MovesFound) Then Exit For
               Next TriggerLength

               If Not FoundMoves(MovesFound) Then
                  For TriggerLength As Integer = WINNING_LENGTH To 0 Step -1
                     MovesFound = FindMoves(PlayersSetup().HumanColor, TriggerLength, AllowHelpingOpponent:=True)
                     If FoundMoves(MovesFound) Then Exit For
                  Next TriggerLength
               End If
            End If
         End If

         If FoundMoves(MovesFound) Then
            Do While Application.OpenForms.Count > 0
               Column = Selector.Next(FIRST_COLUMN, LAST_COLUMN + 1)
               If MovesFound(Column) Then
                  SelectedColumn(NewSelectedColumn:=Column)
                  DiskDropper.Enabled = True
                  Exit Do
               End If
            Loop
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure lets the computer make a move when this game is played against the computer.
   Private Sub ComputerPlayer_Tick() Handles ComputerPlayer.Tick
      Try
         Select Case GetGameState()
            Case GameStatesE.GSRedPlaying, GameStatesE.GSYellowPlaying
               If ActivePlayerColor() = PlayersSetup().ComputerColor Then
                  If Not DiskDropper().Enabled Then ComputerMakeMove()
               End If
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure returns the number of disks of the specified color found using the specified position and direction.
   Private Function CountDisks(StartColumn As Integer, StartRow As Integer, DiskColor As DiskColorsE, XDirection As Integer, YDirection As Integer) As Integer
      Try
         Dim CheckCount As Integer = 0
         Dim Column As Integer = StartColumn
         Dim DiskCount As Integer = 0
         Dim Row As Integer = StartRow

         Do Until CheckCount = WINNING_LENGTH OrElse Application.OpenForms.Count = 0
            Select Case Disks(Column, Row)
               Case DiskColor
                  DiskCount += 1
               Case Not DiskColorsE.DCNone
                  Exit Do
            End Select

            Column += XDirection
            Row += YDirection
            CheckCount += 1
         Loop

         Return DiskCount
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure drops the current player's disk into current the selected column.
   Private Sub DiskDropper_Tick() Handles DiskDropper.Tick
      Try
         Static Row As Integer = FIRST_ROW

         If Disks(SelectedColumn(), Row) = DiskColorsE.DCOutsideField OrElse Not Disks(SelectedColumn(), Row) = DiskColorsE.DCNone Then
            DiskDropper.Enabled = False
            ActivePlayerColor(, ChangeTurns:=True)
            DisplayStatus()
            Row = FIRST_ROW
         Else
            If Row > FIRST_ROW Then Disks(SelectedColumn(), Row - 1, DiskColorsE.DCNone, ResetDisks:=False, RemoveDisk:=True)
            Disks(SelectedColumn(), Row, NewDisk:=ActivePlayerColor())
            InterfaceO().Invalidate()
            Row += 1
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the disks inside the game's playing field.
   Private Function Disks(Optional Column As Integer = FIRST_COLUMN, Optional Row As Integer = FIRST_ROW, Optional NewDisk As DiskColorsE = DiskColorsE.DCNone, Optional ResetDisks As Boolean = False, Optional RemoveDisk As Boolean = False) As DiskColorsE
      Try
         Dim Disk As DiskColorsE = DiskColorsE.DCOutsideField
         Static CurrentDisks(0 To LAST_COLUMN, 0 To LAST_ROW) As DiskColorsE

         If ResetDisks Then
            Disk = DiskColorsE.DCNone

            For Column = FIRST_COLUMN To LAST_COLUMN
               For Row = FIRST_ROW To LAST_ROW
                  CurrentDisks(Column, Row) = DiskColorsE.DCNone
               Next Row
            Next Column
         ElseIf RemoveDisk Then
            CurrentDisks(Column, Row) = DiskColorsE.DCNone
         Else
            If Column >= FIRST_COLUMN AndAlso Column <= LAST_COLUMN Then
               If Row >= FIRST_ROW AndAlso Row <= LAST_ROW Then
                  If Not NewDisk = DiskColorsE.DCNone Then CurrentDisks(Column, Row) = NewDisk
                  Disk = CurrentDisks(Column, Row)
               End If
            End If
         End If

         Return Disk
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure displays the help.
   Private Sub DisplayHelp()
      Try
         Dim HelpText As New StringBuilder

         HelpText.Append($"{My.Application.Info.Description}{NewLine}")
         HelpText.Append($"Keys:{NewLine}")
         HelpText.Append($"A = Restart game.{NewLine}")
         HelpText.Append($"C = No computer player.{NewLine}")
         HelpText.Append($"H = This help.{NewLine}")
         HelpText.Append($"I = Information.{NewLine}")
         HelpText.Append($"R = Red plays first.{NewLine}")
         HelpText.Append($"S = Computer plays as red.{NewLine}")
         HelpText.Append($"Y = Yellow plays first.{NewLine}")
         HelpText.Append("Z = Computer plays as yellow.")

         MessageBox.Show(HelpText.ToString(), My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure displays the status in the interface window's title bar.
   Public Sub DisplayStatus()
      Try
         InterfaceO().Text = $"{My.Application.Info.Title} - {StateText()} H = Help"
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure draws a disk of the specified color at the specified position.
   Private Sub DrawDisk(Column As Integer, Row As Integer, DiskColor As DiskColorsE, Canvas As Graphics)
      Try
         Canvas.DrawRectangle(Pens.Black, Column * SLOT_SIZE, Row * SLOT_SIZE, SLOT_SIZE, SLOT_SIZE)
         Canvas.FillEllipse(New SolidBrush({Color.Blue, Color.Cyan, Color.Red, Color.Yellow}(DiskColor)), CInt((Column + 0.1) * SLOT_SIZE), CInt((Row + 0.1) * SLOT_SIZE), CInt(SLOT_SIZE / 1.25), CInt(SLOT_SIZE / 1.25))
         Canvas.DrawEllipse(Pens.Black, CInt((Column + 0.1) * SLOT_SIZE), CInt((Row + 0.1) * SLOT_SIZE), CInt(SLOT_SIZE / 1.25), CInt(SLOT_SIZE / 1.25))
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure draws the disks inside the game's playing field.
   Public Sub DrawDisks(Canvas As Graphics)
      Try
         Canvas.Clear(InterfaceO().BackColor)

         For Column As Integer = FIRST_COLUMN To LAST_COLUMN
            For Row As Integer = FIRST_ROW To LAST_ROW
               DrawDisk(Column, Row, DiskColor:=Disks(Column, Row), Canvas:=Canvas)
            Next Row
         Next Column

         Select Case GetGameState()
            Case GameStatesE.GSNeitherPlaying, GameStatesE.GSRedWon, GameStatesE.GSYellowWon, GameStatesE.GSTied
               GreyOutDisks(Canvas)
         End Select
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procdure determines which columns can be used by the computer player to make a move.
   Private Function FindMoves(DiskColor As DiskColorsE, TriggerLength As Integer, AllowHelpingOpponent As Boolean) As Boolean()
      Try
         Dim Column As New Integer
         Dim CheckCount As New Integer
         Dim DiskCount As New Integer
         Dim FoundMove As New Integer
         Dim MovesFound(0 To LAST_COLUMN) As Boolean
         Dim Row As New Integer

         For MoveColumn As Integer = FIRST_COLUMN To LAST_COLUMN
            MovesFound(MoveColumn) = False
         Next MoveColumn

         For StartColumn As Integer = FIRST_COLUMN To LAST_COLUMN
            For StartRow As Integer = FIRST_ROW To LAST_ROW
               For XDirection As Integer = -1 To 1
                  For YDirection As Integer = -1 To 1
                     If Not (XDirection = 0 AndAlso YDirection = 0) Then
                        Column = StartColumn
                        CheckCount = 0
                        DiskCount = 0
                        FoundMove = NO_COLUMN
                        Row = StartRow

                        Do Until CheckCount = TriggerLength OrElse Application.OpenForms.Count = 0
                           Select Case Disks(Column, Row)
                              Case DiskColor
                                 DiskCount += 1
                              Case DiskColorsE.DCNone
                                 If Not Disks(Column, Row + 1) = DiskColorsE.DCNone Then FoundMove = Column
                              Case Else
                                 Exit Do
                           End Select

                           Column += XDirection
                           Row += YDirection
                           CheckCount += 1
                        Loop

                        If DiskCount = TriggerLength - 1 Then
                           If Not FoundMove = NO_COLUMN Then
                              If AllowHelpingOpponent Then
                                 MovesFound(FoundMove) = True
                              Else
                                 If Not MoveHelpsOpponent(PlayersSetup().HumanColor, FoundMove) Then MovesFound(FoundMove) = True
                              End If
                           End If
                        End If
                     End If
                  Next YDirection
               Next XDirection
            Next StartRow
         Next StartColumn

         Return MovesFound
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure checks whether the specified
   Private Function FoundMoves(MovesFound() As Boolean) As Boolean
      Try
         Return Not Array.TrueForAll(MovesFound, Function(Item As Boolean) Item = False)
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns whether any more moves can be made by the players.
   Private Function GameDone() As Boolean
      Try
         For Column As Integer = FIRST_COLUMN To LAST_COLUMN
            If Disks(Column, FIRST_ROW) = DiskColorsE.DCNone Then Return False
         Next Column

         Return True
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns the game's state.
   Private Function GetGameState() As GameStatesE
      Try
         Dim GameState As GameStatesE = GameStatesE.GSNeitherPlaying

         Select Case WinningPlayer()
            Case DiskColorsE.DCNone
               If GameDone() Then
                  GameState = GameStatesE.GSTied
               Else
                  Select Case ActivePlayerColor()
                     Case DiskColorsE.DCNone
                        GameState = GameStatesE.GSNeitherPlaying
                     Case DiskColorsE.DCRed
                        GameState = GameStatesE.GSRedPlaying
                     Case DiskColorsE.DCYellow
                        GameState = GameStatesE.GSYellowPlaying
                  End Select
               End If
            Case DiskColorsE.DCRed
               GameState = GameStatesE.GSRedWon
            Case DiskColorsE.DCYellow
               GameState = GameStatesE.GSYellowWon
         End Select

         Return GameState
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure greys out all the disks inside game's playing field.
   Private Sub GreyOutDisks(Canvas As Graphics)
      Try
         For Column As Integer = FIRST_COLUMN To LAST_COLUMN
            For Row As Integer = FIRST_ROW To LAST_ROW
               Canvas.DrawRectangle(Pens.Black, Column * SLOT_SIZE, Row * SLOT_SIZE, SLOT_SIZE, SLOT_SIZE)
               Canvas.FillEllipse(New SolidBrush({Color.DarkBlue, Color.DarkCyan, Color.DarkRed, Color.Goldenrod}(Disks(Column, Row))), CInt((Column + 0.1) * SLOT_SIZE), CInt((Row + 0.1) * SLOT_SIZE), CInt(SLOT_SIZE / 1.25), CInt(SLOT_SIZE / 1.25))
               Canvas.DrawEllipse(Pens.Black, CInt((Column + 0.1) * SLOT_SIZE), CInt((Row + 0.1) * SLOT_SIZE), CInt(SLOT_SIZE / 1.25), CInt(SLOT_SIZE / 1.25))
            Next Row
         Next Column
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure handles any errors that occur.
   Public Sub HandleError(ExceptionO As Exception)
      Try
         MessageBox.Show(ExceptionO.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
      Catch
         Application.Exit()
      End Try
   End Sub

   'This procedure process a move made by a human.
   Public Sub HumanMakeMove(Column As Integer)
      Try
         If Not DiskDropper.Enabled Then
            If Column >= FIRST_COLUMN AndAlso Column <= LAST_COLUMN Then
               Select Case GetGameState()
                  Case GameStatesE.GSRedPlaying, GameStatesE.GSYellowPlaying
                     If ActivePlayerColor() = PlayersSetup().HumanColor OrElse PlayersSetup().ComputerColor = DiskColorsE.DCNone Then
                        SelectedColumn(NewSelectedColumn:=Column)
                        DiskDropper.Enabled = True
                     End If
               End Select
            End If
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure initializes the game.
   Public Sub InitializeGame()
      Try
         Static CurrentFirstColor As DiskColorsE = DiskColorsE.DCNone

         ActivePlayerColor(, , ResetPlayers:=True)
         Disks(, , , ResetDisks:=True)
         InterfaceO.Invalidate()
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the game's interface window.
   Public Function InterfaceO(Optional NewInterface As Form = Nothing) As Form
      Try
         Static CurrentInterface As Form = Nothing

         If NewInterface IsNot Nothing Then
            CurrentInterface = NewInterface

            With CurrentInterface
               .BackColor = Color.Blue
               .Width = (Abs(LAST_COLUMN - FIRST_COLUMN) + 1) * SLOT_SIZE
               .Height = (Abs(LAST_ROW - FIRST_ROW) + 1) * SLOT_SIZE
               .Width += (.Width - .ClientRectangle.Width)
               .Height += (.Height - .ClientRectangle.Height)
               .Text = Nothing
            End With
         End If

         Return CurrentInterface
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure indicates whether dropping a disk at the specified column can help the specified opponent.
   Private Function MoveHelpsOpponent(OpponentColor As DiskColorsE, Column As Integer) As Boolean
      Try
         Dim Row As Integer = 0

         Do Until Row = LAST_ROW OrElse Not Disks(Column, Row + 1) = DiskColorsE.DCNone OrElse Application.OpenForms.Count = 0
            Row += 1
         Loop

         If Row > FIRST_ROW Then
            Row -= 1

            For XDirection As Integer = -1 To 1
               For YDirection As Integer = -1 To 1
                  If Not (XDirection = 0 AndAlso YDirection = 0) Then
                     If CountDisks(Column, Row, OpponentColor, XDirection, YDirection) = WINNING_LENGTH - 1 Then
                        Return True
                     End If
                  End If
               Next YDirection
            Next XDirection
         End If

         Return False
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the players' setup.
   Public Function PlayersSetup(Optional NewComputerColor As DiskColorsE = DiskColorsE.DCNone, Optional NewFirstColor As DiskColorsE = DiskColorsE.DCNone, Optional NoComputerPlayer As Boolean = False) As PlayersSetupStr
      Try
         Static CurrentPlayersSetup As New PlayersSetupStr With {.ComputerColor = DiskColorsE.DCNone, .FirstColor = DiskColorsE.DCRed, .HumanColor = DiskColorsE.DCYellow}

         With CurrentPlayersSetup
            If NoComputerPlayer Then
               .ComputerColor = DiskColorsE.DCNone
            Else
               If Not NewComputerColor = DiskColorsE.DCNone Then .ComputerColor = NewComputerColor
            End If

            If Not NewFirstColor = DiskColorsE.DCNone Then .FirstColor = NewFirstColor

            Select Case .ComputerColor
               Case DiskColorsE.DCNone
                  .HumanColor = DiskColorsE.DCNone
               Case DiskColorsE.DCRed
                  .HumanColor = DiskColorsE.DCYellow
               Case DiskColorsE.DCYellow
                  .HumanColor = DiskColorsE.DCRed
            End Select

            ComputerPlayer.Enabled = Not (.ComputerColor = DiskColorsE.DCNone)
         End With

         Return CurrentPlayersSetup
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure manages the game.
   Public Sub PlayGame(KeyStroke As Keys)
      Try
         Select Case GetGameState()
            Case GameStatesE.GSRedWon, GameStatesE.GSYellowWon, GameStatesE.GSTied
               DrawDisks(InterfaceO().CreateGraphics)
         End Select

         If Not KeyStroke = Nothing Then
            Select Case GetGameState()
               Case GameStatesE.GSNeitherPlaying
                  ActivePlayerColor(NewPlayer:=PlayersSetup().FirstColor)
                  InterfaceO.Invalidate()
               Case GameStatesE.GSRedWon, GameStatesE.GSYellowWon, GameStatesE.GSTied
                  InitializeGame()
               Case Else
                  Select Case KeyStroke
                     Case Keys.A
                        InitializeGame()
                     Case Keys.C
                        PlayersSetup(, , NoComputerPlayer:=True)
                        InitializeGame()
                     Case Keys.H
                        DisplayHelp()
                     Case Keys.I
                        With My.Application.Info
                           MessageBox.Show(String.Format("{0} v{1} - by: {2}", .Title, .Version.ToString, .CompanyName), .Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End With
                     Case Keys.R
                        PlayersSetup(, NewFirstColor:=DiskColorsE.DCRed)
                        InitializeGame()
                     Case Keys.S
                        PlayersSetup(NewComputerColor:=DiskColorsE.DCRed)
                        InitializeGame()
                     Case Keys.Y
                        PlayersSetup(, NewFirstColor:=DiskColorsE.DCYellow)
                        InitializeGame()
                     Case Keys.Z
                        PlayersSetup(NewComputerColor:=DiskColorsE.DCYellow)
                        InitializeGame()
                  End Select
            End Select
         End If
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try
   End Sub

   'This procedure manages the selected column.
   Private Function SelectedColumn(Optional NewSelectedColumn As Integer = NO_COLUMN) As Integer
      Try
         Static CurrentSelectedColumn As Integer = NO_COLUMN

         If Not NewSelectedColumn = NO_COLUMN Then CurrentSelectedColumn = NewSelectedColumn

         Return CurrentSelectedColumn
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns a text description of the game's state.
   Private Function StateText() As String
      Try
         Return {"Inactive. - Press any key.", "Red's turn.", "Red won.", "Game is tied.", "Yellow's turn.", "Yellow won."}(GetGameState())
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function

   'This procedure returns which player has won.
   Private Function WinningPlayer() As DiskColorsE
      Try
         For Each DiskColor As DiskColorsE In {DiskColorsE.DCRed, DiskColorsE.DCYellow}
            For Column As Integer = FIRST_COLUMN To LAST_COLUMN
               For Row As Integer = FIRST_ROW To LAST_ROW
                  For XDirection As Integer = -1 To 1
                     For YDirection As Integer = -1 To 1
                        If Not (XDirection = 0 AndAlso YDirection = 0) Then
                           If CountDisks(Column, Row, DiskColor, XDirection, YDirection) = WINNING_LENGTH Then
                              Return DiskColor
                           End If
                        End If
                     Next YDirection
                  Next XDirection
               Next Row
            Next Column
         Next DiskColor

         Return DiskColorsE.DCNone
      Catch ExceptionO As Exception
         HandleError(ExceptionO)
      End Try

      Return Nothing
   End Function
End Module
