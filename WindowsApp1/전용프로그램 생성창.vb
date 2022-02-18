Public Class Form8
    Public Ans As Integer
    Public Change_Name As String
    Public default_name As String
    Dim Select_list As String
    Const prohibit_name As String = "Mitutoyo Result Matcher"
    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles Me.Load
        Ans = 2
        Select_list = Form1.ListBox1.SelectedItem.ToString()
        TextBox1.Text = default_name
        Me.Location = New Point(Form1.Location.X + 80, Form1.Location.Y + 80)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Ans = 1         '취소
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Ans = 0         '확인
        Change_Name = TextBox1.Text

        If Change_Name = prohibit_name Then
            MsgBox("기본 이름과 동일하게 전용 프로그램을 생성할수 없습니다.", 64, "Error Occured")
            Ans = 2
        End If

        Me.Close()
    End Sub


    ' 텍스트박스에서 엔터 눌렀을때 키코드에 대응해서 이벤트 발생 ↓↓↓↓
    Private Sub KEY_DOWN_EVENT(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = 13 Then
            Button1_Click(sender, New System.EventArgs())
        End If

    End Sub
End Class