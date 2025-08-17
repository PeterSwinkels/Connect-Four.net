'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Windows.Forms

'This module contains this program's main interface window.
Public Class InterfaceWindow
   'This procedure initializes this window.
   Public Sub New()
      Try
         InitializeComponent()
         InterfaceO(NewInterface:=Me)

         PlayersSetup(NewComputerColor:=DiskColorsE.DCYellow, NewFirstColor:=DiskColorsE.DCRed)
         InitializeGame()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to set the most recently pressed key.
   Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
      Try
         PlayGame(e.KeyCode)
      Catch ExceptionO As Exception
         DisplayException(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to set the most recently selected column of disks.
   Private Sub Form_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
      Try
         HumanMakeMove(Column:=(e.X \ SLOT_SIZE))
      Catch ExceptionO As Exception
         DisplayException(ExceptionO)
      End Try
   End Sub

   'This procedure gives the command to draw the graphics.
   Private Sub Form_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
      Try
         DrawDisks(e.Graphics)
         DisplayStatus()
      Catch ExceptionO As Exception
         DisplayException(ExceptionO)
      End Try
   End Sub
End Class
