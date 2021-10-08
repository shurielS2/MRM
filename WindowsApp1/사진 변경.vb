Public Class Form12
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click       '취소
        DialogResult = 2
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click       '삭제
        DialogResult = 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click       '확인
        DialogResult = 3
    End Sub
End Class