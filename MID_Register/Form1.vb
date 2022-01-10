Public Class Form1

    Dim MID As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)
        MID = Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        Button2.Enabled = True

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        Button2.Enabled = False
        'TextBox1.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sfd_For_manual As New SaveFileDialog()
        Dim text_for_save As String

        If TextBox2.Text = "" Then
            MsgBox("업체명을 적어주세요")
            GoTo skip
        End If

        If TextBox3.Text = "" Then
            MsgBox("담당자를 적어주세요")
            GoTo skip
        End If

        If TextBox4.Text = "" Then
            MsgBox("소재지를 적어주세요(ex: 경기도)")
            GoTo skip
        End If

        If TextBox5.Text = "" Then
            MsgBox("사용장비를 적어주세요")
            GoTo skip
        End If

        If TextBox6.Text = "" Then
            MsgBox("사용장비의 시리얼번호를 적어주세요")
            GoTo skip
        End If

        If TextBox7.Text = "" Then
            MsgBox("E-mail을 적어주세요")
            GoTo skip
        End If

        text_for_save = TextBox2.Text & "," & TextBox3.Text & "," & TextBox4.Text & "," & TextBox5.Text & "," & TextBox6.Text & "," & TextBox7.Text & "," & TextBox1.Text
        With sfd_For_manual
            .InitialDirectory = "C:\desktop\"
            .Filter = "MRM REG|.MREG"
            .FilterIndex = 1
            .Title = "MRM Register"
            .FileName = TextBox2.Text & "_" & TextBox6.Text
            .RestoreDirectory = True
            .CheckFileExists = False
            .CheckPathExists = True
        End With


        If sfd_For_manual.ShowDialog() = Windows.Forms.DialogResult.OK Then

            My.Computer.FileSystem.WriteAllText(sfd_For_manual.FileName, text_for_save, False, System.Text.Encoding.Default)

        End If



skip:
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> MID Then
            TextBox2.Text = ""
            TextBox2.Enabled = False
            TextBox3.Text = ""
            TextBox3.Enabled = False
            TextBox4.Text = ""
            TextBox4.Enabled = False
            TextBox5.Text = ""
            TextBox5.Enabled = False
            TextBox6.Text = ""
            TextBox6.Enabled = False
            TextBox7.Text = ""
            TextBox7.Enabled = False
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class
