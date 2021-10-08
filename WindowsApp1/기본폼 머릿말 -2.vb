Public Class Form9
    Public Product_Name As String
    Public Machine_Name As String
    Public Request_Dept As String
    Public Request_Date As String
    Public Drawing_Num As String
    Public Program_Name As String
    Public Player_Name As String
    Public Measure_Date As String

    Public Check_date_1 As String
    Public Check_date_2 As String

    Public select_pic_name As String

    Dim select_image As Bitmap
    Dim origin_image As Bitmap
    Dim change_image As Bitmap

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click           '취소

        Me.Close()
    End Sub
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles Me.Load
        origin_image = My.Resources._1920px_Mitutoyo_company_logo_sample
        Me.Location = New Point(Form1.Location.X + 80, Form1.Location.Y + 80)

        If Form1.New_Fix_check = 0 Then
            TextBox1.Text = ""


            TextBox5.Text = ""

            CheckBox1.Checked = True
            CheckBox2.Checked = True

        End If

        If Form1.New_Fix_check = 1 Then
            '====================================== 다른 성적서 폼에서도 동일 내용 불러오기위한 과정   2021.01.11 변경
            TextBox1.Text = Form1.user_info_temp(0)
            'TextBox2.Text = Form1.user_info_temp(1)
            'TextBox3.Text = Form1.user_info_temp(2)
            TextBox4.Text = Form1.user_info_temp(3)
            TextBox5.Text = Form1.user_info_temp(4)
            'TextBox6.Text = Form1.user_info_temp(5)
            'TextBox7.Text = Form1.user_info_temp(6)
            TextBox8.Text = Form1.user_info_temp(7)
            CheckBox1.Checked = Form1.user_info_temp(8)
            CheckBox2.Checked = Form1.user_info_temp(9)
            'select_pic_name = Form1.user_info_temp(10)
            'select_pic_name_2 = Form1.user_info_temp(11)

            '    If select_pic_name = "" Then
            '    PictureBox2.Image = origin_image
            '' Else
            '    PictureBox2.Image = New Bitmap(select_pic_name)
            ' End If
            ''
            '     If select_pic_name_2 = "" Then
            '     PictureBox3.Image = origin_image
            ' Else
            '     PictureBox3.Image = New Bitmap(select_pic_name_2)
            ' End If

            '====================================== 이전 구문 2021.01.11 변경

            ' TextBox1.Text = Product_Name
            'TextBox2.Text = Machine_Name
            'TextBox3.Text = Request_Dept
            ' TextBox4.Text = Request_Date
            ' TextBox5.Text = Drawing_Num
            'TextBox6.Text = Program_Name
            'TextBox7.Text = Player_Name
            ' TextBox8.Text = Measure_Date

            'PictureBox2.Image = select_image
            'PictureBox3.Image = select_image_2
        End If


        If Form1.New_Fix_check = 2 Then
            TextBox1.Text = 수정창.User_Info_Value(0)

            TextBox5.Text = 수정창.User_Info_Value(4)


            If 수정창.User_Info_Value(8) = "True" Then
                CheckBox1.Checked = True
            Else
                TextBox4.Text = 수정창.User_Info_Value(3)
            End If

            If 수정창.User_Info_Value(9) = "True" Then
                CheckBox2.Checked = True
            Else

                TextBox8.Text = 수정창.User_Info_Value(7)
            End If

            If 수정창.User_Info_Value(10) = "" Then

                select_pic_name = ""
            Else

            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           '확인

        Form1.user_info_temp(0) = TextBox1.Text
        Form1.user_info_temp(1) = ""
        Form1.user_info_temp(2) = ""
        Form1.user_info_temp(3) = TextBox4.Text
        Form1.user_info_temp(4) = TextBox5.Text
        Form1.user_info_temp(5) = ""
        Form1.user_info_temp(6) = ""
        Form1.user_info_temp(7) = TextBox8.Text
        Form1.user_info_temp(8) = CheckBox1.Checked.ToString
        Form1.user_info_temp(9) = CheckBox2.Checked.ToString
        Form1.user_info_temp(10) = ""
        Form1.user_info_temp(11) = ""


        Product_Name = TextBox1.Text
        Machine_Name = ""
        Request_Dept = ""
        Request_Date = TextBox4.Text
        Drawing_Num = TextBox5.Text
        Program_Name = ""
        Player_Name = ""
        Measure_Date = TextBox8.Text
        Check_date_1 = CheckBox1.Checked.ToString
        Check_date_2 = CheckBox2.Checked.ToString

        select_pic_name = ""

        select_image = Nothing

        Form1.New_Fix_check = 1
        수정창.user_info_count = 1
        Me.Close()

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        TextBox4.Text = DateTimePicker1.Value.Date
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        TextBox8.Text = DateTimePicker2.Value.Date
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox4.Enabled = False
            DateTimePicker1.Enabled = False

            TextBox4.Text = DateTime.Now.Date

        End If

        If CheckBox1.Checked = False Then
            TextBox4.Enabled = True
            DateTimePicker1.Enabled = True

            TextBox4.Text = ""

        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            TextBox8.Enabled = False
            DateTimePicker2.Enabled = False

            TextBox8.Text = DateTime.Now.Date

        End If

        If CheckBox2.Checked = False Then
            TextBox8.Enabled = True
            DateTimePicker2.Enabled = True

            TextBox8.Text = ""

        End If
    End Sub


End Class