Public Class 활성화키
    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Focus()
        TextBox1.PasswordChar = "*"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub
    Private Sub KEY_Down_EVENT(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = 13 Then      '13 = Enter
            Button1_Click(sender, New System.EventArgs())
        End If

    End Sub
End Class