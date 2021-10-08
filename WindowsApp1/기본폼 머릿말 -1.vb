Public Class Form7
    Const MRM_root_dir As String = "C:\MitutoyoApp"

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
    Public select_pic_name_2 As String


    Dim select_image As Bitmap
    Dim select_image_2 As Bitmap
    Dim origin_image As Bitmap
    Dim origin_image_2 As Bitmap
    Dim change_image As Bitmap

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click           '취소

        Me.Close()
    End Sub
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles Me.Load
        origin_image = My.Resources._1920px_Mitutoyo_company_logo_sample
        Me.Location = New Point(Form1.Location.X + 80, Form1.Location.Y + 80)

        If Form1.New_Fix_check = 0 Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""

            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox7.Text = ""
            CheckBox1.Checked = True
            CheckBox2.Checked = True

            PictureBox2.Image = origin_image
            PictureBox3.Image = origin_image

        End If

        If Form1.New_Fix_check = 1 Then
            '====================================== 다른 성적서 폼에서도 동일 내용 불러오기위한 과정   2021.01.11 변경
            TextBox1.Text = Form1.user_info_temp(0)
            TextBox2.Text = Form1.user_info_temp(1)
            TextBox3.Text = Form1.user_info_temp(2)
            TextBox4.Text = Form1.user_info_temp(3)
            TextBox5.Text = Form1.user_info_temp(4)
            TextBox6.Text = Form1.user_info_temp(5)
            TextBox7.Text = Form1.user_info_temp(6)
            TextBox8.Text = Form1.user_info_temp(7)
            CheckBox1.Checked = Form1.user_info_temp(8)
            CheckBox2.Checked = Form1.user_info_temp(9)
            select_pic_name = Form1.user_info_temp(10)
            select_pic_name_2 = Form1.user_info_temp(11)


            If select_pic_name = "" Then
                PictureBox2.Image = origin_image
            Else
                PictureBox2.Image = New Bitmap(select_pic_name)
            End If

            If select_pic_name_2 = "" Then
                PictureBox3.Image = origin_image
            Else
                PictureBox3.Image = New Bitmap(select_pic_name_2)
            End If

            '====================================== 이전 구문 2021.01.11 변경

            '  TextBox1.Text = Product_Name
            '  TextBox2.Text = Machine_Name
            '  TextBox3.Text = Request_Dept
            '  TextBox4.Text = Request_Date
            '  TextBox5.Text = Drawing_Num
            '  TextBox6.Text = Program_Name
            '  TextBox7.Text = Player_Name
            '  TextBox8.Text = Measure_Date

            '  PictureBox2.Image = select_image
        End If


        If Form1.New_Fix_check = 2 Then
            TextBox1.Text = 수정창.User_Info_Value(0)
            TextBox2.Text = 수정창.User_Info_Value(1)
            TextBox3.Text = 수정창.User_Info_Value(2)

            TextBox5.Text = 수정창.User_Info_Value(4)
            TextBox6.Text = 수정창.User_Info_Value(5)
            TextBox7.Text = 수정창.User_Info_Value(6)


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
                PictureBox2.Image = origin_image
                select_pic_name = ""
            Else
                PictureBox2.Image = New Bitmap(수정창.User_Info_Value(10))
                select_pic_name = 수정창.User_Info_Value(10)
            End If

            If 수정창.User_Info_Value(11) = "" Then
                PictureBox3.Image = origin_image
                select_pic_name_2 = ""
            Else
                PictureBox3.Image = New Bitmap(수정창.User_Info_Value(11))
                select_pic_name_2 = 수정창.User_Info_Value(11)
            End If

        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           '확인

        '===================================================================== form1 변수에 저장  다른 폼 공유용
        Form1.user_info_temp(0) = TextBox1.Text
        Form1.user_info_temp(1) = TextBox2.Text
        Form1.user_info_temp(2) = TextBox3.Text
        Form1.user_info_temp(3) = TextBox4.Text
        Form1.user_info_temp(4) = TextBox5.Text
        Form1.user_info_temp(5) = TextBox6.Text
        Form1.user_info_temp(6) = TextBox7.Text
        Form1.user_info_temp(7) = TextBox8.Text
        Form1.user_info_temp(8) = CheckBox1.Checked.ToString
        Form1.user_info_temp(9) = CheckBox2.Checked.ToString

        If select_pic_name = "" Then

        Else
            Form1.user_info_temp(10) = select_pic_name
        End If

        If select_pic_name = "" Then

        Else
            Form1.user_info_temp(11) = select_pic_name_2
        End If


        '===================================================================== 본래 form변수에 저장
        Product_Name = TextBox1.Text
        Machine_Name = TextBox2.Text
        Request_Dept = TextBox3.Text
        Request_Date = TextBox4.Text
        Drawing_Num = TextBox5.Text
        Program_Name = TextBox6.Text
        Player_Name = TextBox7.Text
        Measure_Date = TextBox8.Text
        Check_date_1 = CheckBox1.Checked.ToString
        Check_date_2 = CheckBox2.Checked.ToString
        If select_pic_name = "" Then
            select_image = origin_image
        Else
            select_image = New Bitmap(select_pic_name)
        End If

        If select_pic_name_2 = "" Then
            select_image_2 = origin_image
        Else
            select_image_2 = New Bitmap(select_pic_name_2)
        End If

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


    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        origin_image = PictureBox2.Image

        If select_pic_name = "" Then

            Dim ofd_For_Csv As New OpenFileDialog()
            With ofd_For_Csv
                .InitialDirectory = MRM_root_dir & "\MRM\data"
                .Filter = "PNG(*.png)|*.png|JPEG(*.jpg,*.jpeg,*.jpe,*.jfif)|*.jpg;*.jpeg;*.jpe;*.jfif|GIF(*.gif)|*gif|BMP(*.bmp,*.dib)|*.bmp;*.dib|TIFF(*.tiff,*.tif)|*.tiff;*.tif|All File(*.*)|*.*"
                .FilterIndex = 1
                .Title = "Change Picture "
                .RestoreDirectory = True
                .CheckFileExists = True
                .CheckPathExists = True
            End With
            If ofd_For_Csv.ShowDialog() = Windows.Forms.DialogResult.OK Then
                change_image = New Bitmap(ofd_For_Csv.FileName)
                PictureBox2.Image = change_image
                select_pic_name = ofd_For_Csv.FileName
            End If

        Else

            Form12.ShowDialog()

            Select Case Form12.DialogResult

                Case 1      '삭제

                    PictureBox2.Image = My.Resources._1920px_Mitutoyo_company_logo_sample
                    select_pic_name = ""

                Case 2      '취소

                Case 3      '확인

                    Dim ofd_For_Csv As New OpenFileDialog()
                    With ofd_For_Csv
                        .InitialDirectory = MRM_root_dir & "\MRM\data"
                        .Filter = "PNG(*.png)|*.png|JPEG(*.jpg,*.jpeg,*.jpe,*.jfif)|*.jpg;*.jpeg;*.jpe;*.jfif|GIF(*.gif)|*gif|BMP(*.bmp,*.dib)|*.bmp;*.dib|TIFF(*.tiff,*.tif)|*.tiff;*.tif|All File(*.*)|*.*"
                        .FilterIndex = 1
                        .Title = "Change Picture "
                        .RestoreDirectory = True
                        .CheckFileExists = True
                        .CheckPathExists = True
                    End With
                    If ofd_For_Csv.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        change_image = New Bitmap(ofd_For_Csv.FileName)
                        PictureBox2.Image = change_image
                        select_pic_name = ofd_For_Csv.FileName
                    End If

            End Select

        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        origin_image_2 = PictureBox3.Image

        If select_pic_name_2 = "" Then

            Dim ofd_For_Csv As New OpenFileDialog()
            With ofd_For_Csv
                .InitialDirectory = MRM_root_dir & "\MRM\data"
                .Filter = "PNG(*.png)|*.png|JPEG(*.jpg,*.jpeg,*.jpe,*.jfif)|*.jpg;*.jpeg;*.jpe;*.jfif|GIF(*.gif)|*gif|BMP(*.bmp,*.dib)|*.bmp;*.dib|TIFF(*.tiff,*.tif)|*.tiff;*.tif|All File(*.*)|*.*"
                .FilterIndex = 1
                .Title = "Change Picture "
                .RestoreDirectory = True
                .CheckFileExists = True
                .CheckPathExists = True
            End With

            If ofd_For_Csv.ShowDialog() = Windows.Forms.DialogResult.OK Then
                change_image = New Bitmap(ofd_For_Csv.FileName)
                PictureBox3.Image = change_image
                select_pic_name_2 = ofd_For_Csv.FileName
            End If

        Else
            Form12.ShowDialog()
            Select Case Form12.DialogResult

                Case 1      '삭제

                    PictureBox3.Image = My.Resources._1920px_Mitutoyo_company_logo_sample
                    select_pic_name_2 = ""
                Case 2      '취소

                Case 3      '확인

                    Dim ofd_For_Csv As New OpenFileDialog()
                    With ofd_For_Csv
                        .InitialDirectory = MRM_root_dir & "\MRM\data"
                        .Filter = "PNG(*.png)|*.png|JPEG(*.jpg,*.jpeg,*.jpe,*.jfif)|*.jpg;*.jpeg;*.jpe;*.jfif|GIF(*.gif)|*gif|BMP(*.bmp,*.dib)|*.bmp;*.dib|TIFF(*.tiff,*.tif)|*.tiff;*.tif|All File(*.*)|*.*"
                        .FilterIndex = 1
                        .Title = "Change Picture "
                        .RestoreDirectory = True
                        .CheckFileExists = True
                        .CheckPathExists = True
                    End With

                    If ofd_For_Csv.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        change_image = New Bitmap(ofd_For_Csv.FileName)
                        PictureBox3.Image = change_image
                        select_pic_name_2 = ofd_For_Csv.FileName
                    End If

            End Select

        End If
    End Sub
End Class