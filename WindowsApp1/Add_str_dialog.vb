Public Class Add_str_dialog
    Dim Description As String
    Dim ans_msgbox As Integer
    Public ans_textbox As String
    Private Sub Add_str_dialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = Form1.add_str_value(0)
        TextBox1.Text = ""
        TextBox1.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ans_textbox = TextBox1.Text
        Me.DialogResult = Windows.Forms.DialogResult.OK             'dialog result 값 정하고 dialog 닫기
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ans_msgbox = MsgBox("추가 내용 기입을 취소 하시겠습니까?", 4, "추가 내용 기입 취소")
        Select Case ans_msgbox
            Case 6      'yes
                Me.Close()

            Case 7      'no

        End Select
    End Sub
    Private Sub KEY_Down_EVENT(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = 13 Then      '13 = Enter
            If e.Shift Then

            Else
                Button1_Click(sender, New System.EventArgs())
            End If

        End If
    End Sub
End Class