Public Class Form1
    Dim CorClicado As Color = Color.FromArgb(50, 50, 50)
    Dim CorNaoClicado As Color = Color.FromArgb(64, 64, 64)
    Dim Jogada As Decimal
    Dim Campo1 As String
    Dim Campo2 As String
    Dim Campo3 As String
    Dim Campo4 As String
    Dim Campo5 As String
    Dim Campo6 As String
    Dim Campo7 As String
    Dim Campo8 As String
    Dim Campo9 As String
    Dim Jajogou1 As Decimal
    Dim Jajogou2 As Decimal
    Dim Jajogou3 As Decimal
    Dim Jajogou4 As Decimal
    Dim Jajogou5 As Decimal
    Dim Jajogou6 As Decimal
    Dim Jajogou7 As Decimal
    Dim Jajogou8 As Decimal
    Dim Jajogou9 As Decimal
    Dim Vez As String

    Private Sub picboxCampo1_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo1.MouseEnter
        picboxCampo1.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo1_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo1.MouseLeave
        picboxCampo1.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo2_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo2.MouseEnter
        picboxCampo2.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo2_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo2.MouseLeave
        picboxCampo2.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo3_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo3.MouseEnter
        picboxCampo3.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo3_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo3.MouseLeave
        picboxCampo3.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo4_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo4.MouseEnter
        picboxCampo4.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo4_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo4.MouseLeave
        picboxCampo4.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo5_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo5.MouseEnter
        picboxCampo5.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo5_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo5.MouseLeave
        picboxCampo5.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo6_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo6.MouseEnter
        picboxCampo6.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo6_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo6.MouseLeave
        picboxCampo6.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo7_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo7.MouseEnter
        picboxCampo7.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo7_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo7.MouseLeave
        picboxCampo7.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo8_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo8.MouseEnter
        picboxCampo8.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo8_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo8.MouseLeave
        picboxCampo8.BackColor = CorNaoClicado
    End Sub
    Private Sub picboxCampo9_MouseEnter(sender As Object, e As EventArgs) Handles picboxCampo9.MouseEnter
        picboxCampo9.BackColor = CorClicado
    End Sub
    Private Sub picboxCampo9_MouseLeave(sender As Object, e As EventArgs) Handles picboxCampo9.MouseLeave
        picboxCampo9.BackColor = CorNaoClicado
    End Sub
    Sub VerificarSeHaGanhador()
        If Campo1 = "X" And Campo2 = "X" And Campo3 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo4 = "X" And Campo5 = "X" And Campo6 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo7 = "X" And Campo8 = "X" And Campo9 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo1 = "X" And Campo5 = "X" And Campo9 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo2 = "X" And Campo5 = "X" And Campo8 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo3 = "X" And Campo5 = "X" And Campo7 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo1 = "X" And Campo4 = "X" And Campo7 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo3 = "X" And Campo6 = "X" And Campo9 = "X" Then
            MsgBox("X é o ganhador")
            Restart()
        ElseIf Campo1 = "O" And Campo2 = "O" And Campo3 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo4 = "O" And Campo5 = "O" And Campo6 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo7 = "O" And Campo8 = "O" And Campo9 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo1 = "O" And Campo5 = "O" And Campo9 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo2 = "O" And Campo5 = "O" And Campo8 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo3 = "O" And Campo5 = "O" And Campo7 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo1 = "O" And Campo4 = "O" And Campo7 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Campo3 = "O" And Campo6 = "O" And Campo9 = "O" Then
            MsgBox("O é o ganhador")
            Restart()
        ElseIf Jogada = 10 Then
            MsgBox("Deu Velha")
            Restart()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Jogada = 1
        PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
    End Sub

    Private Sub picboxCampo1_Click(sender As Object, e As MouseEventArgs) Handles picboxCampo1.Click
        If Jajogou1 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo1.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo1 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo1.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo1 = "O"
            End If
            Jogada += 1
            Jajogou1 = 1
            VerificarSeHaGanhador()

        End If

    End Sub
    Private Sub picboxCampo2_Click(sender As Object, e As EventArgs) Handles picboxCampo2.Click
        If Jajogou2 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo2.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo2 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo2.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo2 = "O"

            End If
            Jogada += 1
            Jajogou2 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo3_Click(sender As Object, e As EventArgs) Handles picboxCampo3.Click
        If Jajogou3 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo3 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo3 = "O"
            End If
            Jogada += 1
            Jajogou3 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo4_Click(sender As Object, e As EventArgs) Handles picboxCampo4.Click
        If Jajogou4 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo4.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo4 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo4.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo4 = "O"
            End If
            Jogada += 1
            Jajogou4 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo5_Click(sender As Object, e As EventArgs) Handles picboxCampo5.Click
        If Jajogou5 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo5.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo5 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo5.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo5 = "O"
            End If
            Jogada += 1
            Jajogou5 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo6_Click(sender As Object, e As EventArgs) Handles picboxCampo6.Click
        If Jajogou6 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo6.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo6 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo6.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo6 = "O"
            End If
            Jogada += 1
            Jajogou6 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo7_Click(sender As Object, e As EventArgs) Handles picboxCampo7.Click
        If Jajogou7 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo7.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo7 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo7.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo7 = "O"
            End If
            Jogada += 1
            Jajogou7 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo8_Click(sender As Object, e As EventArgs) Handles picboxCampo8.Click
        If Jajogou8 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo8.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo8 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo8.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo8 = "O"
            End If
            Jogada += 1
            Jajogou8 = 1
            VerificarSeHaGanhador()

        End If

    End Sub

    Private Sub picboxCampo9_Click(sender As Object, e As EventArgs) Handles picboxCampo9.Click
        If Jajogou9 <> 1 Then
            If Jogada Mod 2 = 1 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                picboxCampo9.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                Campo9 = "X"
            ElseIf Jogada Mod 2 = 0 Then
                PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
                picboxCampo9.Image = TicTacToeTeste.My.Resources.Resources.TTT___O_black
                Campo9 = "O"
            End If
            Jogada += 1
            Jajogou9 = 1
            VerificarSeHaGanhador()

        End If

    End Sub
    Sub Restart()
        Jogada = 1
        PictureBox3.Image = TicTacToeTeste.My.Resources.Resources.TTT___X_black
        picboxCampo1.Image = Nothing
        picboxCampo2.Image = Nothing
        picboxCampo3.Image = Nothing
        picboxCampo4.Image = Nothing
        picboxCampo5.Image = Nothing
        picboxCampo6.Image = Nothing
        picboxCampo7.Image = Nothing
        picboxCampo8.Image = Nothing
        picboxCampo9.Image = Nothing
        Campo1 = ""
        Campo2 = ""
        Campo3 = ""
        Campo4 = ""
        Campo5 = ""
        Campo6 = ""
        Campo7 = ""
        Campo8 = ""
        Campo9 = ""
        Jajogou1 = 0
        Jajogou2 = 0
        Jajogou3 = 0
        Jajogou4 = 0
        Jajogou5 = 0
        Jajogou6 = 0
        Jajogou7 = 0
        Jajogou8 = 0
        Jajogou9 = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Restart()
    End Sub
End Class
