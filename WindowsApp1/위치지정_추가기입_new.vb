Public Class 위치지정_추가기입_new
    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Const MRM_root_dir As String = "C:\MitutoyoApp"

    Dim add_Str_section As String
    Public add_str_keyname(20) As String
    Public add_str_value(20) As String
    Dim ini_dir As String

    Public tab_names() As String
    Public tab_count As Integer


    Private Sub 위치지정_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim section_num As Integer

        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        For i = 1 To tab_count          'mknew에서 tab_count값 넣어줌
            ComboBox4.Items.Add(tab_names(i))
            ComboBox5.Items.Add(tab_names(i))
            ComboBox6.Items.Add(tab_names(i))
        Next

        If MkNew.add_str_count = 0 Then



        ElseIf MkNew.add_str_count = 1 Then
            For section_num = 1 To 3
                Select Case section_num
                    Case 1

                        If add_str_value(4) <> "" Then

                            TextBox1.Text = MkNew.Temp_User_Info(0)
                            TextBox2.Text = MkNew.Temp_User_Info(1)
                            TextBox3.Text = MkNew.Temp_User_Info(2)
                            ComboBox1.Text = MkNew.Temp_User_Info(3)
                            CheckBox1.Checked = MkNew.Temp_User_Info(4)
                            ComboBox4.Text = MkNew.Temp_User_Info(15)
                            ComboBox7.Text = MkNew.Temp_User_Info(18)
                        Else
                            TextBox1.Text = ""
                            TextBox2.Text = ""
                            TextBox3.Text = ""
                            ComboBox1.Text = ""
                            CheckBox1.Checked = False
                            ComboBox4.Text = ""

                        End If


                    Case 2
                        If add_str_value(9) <> "" Then

                            TextBox4.Text = MkNew.Temp_User_Info(5)
                            TextBox5.Text = MkNew.Temp_User_Info(6)
                            TextBox6.Text = MkNew.Temp_User_Info(7)
                            ComboBox2.Text = MkNew.Temp_User_Info(8)

                            CheckBox2.Checked = MkNew.Temp_User_Info(9)
                            ComboBox5.Text = MkNew.Temp_User_Info(16)
                            ComboBox8.Text = MkNew.Temp_User_Info(19)
                        Else

                            TextBox4.Text = ""
                            TextBox5.Text = ""
                            TextBox6.Text = ""
                            ComboBox2.Text = ""
                            CheckBox2.Checked = False
                            ComboBox5.Text = ""
                        End If
                    Case 3
                        If add_str_value(14) <> "" Then

                            TextBox7.Text = MkNew.Temp_User_Info(10)
                            TextBox8.Text = MkNew.Temp_User_Info(11)
                            TextBox9.Text = MkNew.Temp_User_Info(12)
                            ComboBox3.Text = MkNew.Temp_User_Info(13)
                            CheckBox3.Checked = MkNew.Temp_User_Info(14)
                            ComboBox6.Text = MkNew.Temp_User_Info(17)
                            ComboBox9.Text = MkNew.Temp_User_Info(20)
                        Else
                            TextBox7.Text = ""
                            TextBox8.Text = ""
                            TextBox9.Text = ""
                            ComboBox3.Text = ""
                            CheckBox3.Checked = False
                            ComboBox6.Text = ""
                        End If

                End Select
            Next section_num
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click       ' 저장 버튼
        Dim i As Integer

        For i = 1 To 3
            ini_input(i)
        Next

        MkNew.add_str_count = 1

    End Sub


    Sub ini_input(panel_num As Integer)

        add_Str_section = "add_str_" & panel_num
        add_str_keyname(0) = "Description"
        add_str_keyname(1) = "value"
        add_str_keyname(2) = "loction"
        add_str_keyname(3) = "combo"
        add_str_keyname(4) = "use_check"
        add_str_keyname(5) = "apply_tab"
        add_str_keyname(6) = "input_type"

        Select Case panel_num
            Case 1
                If CheckBox1.Checked.ToString = True Then
                    add_str_value(0) = TextBox1.Text
                    add_str_value(1) = TextBox2.Text
                    add_str_value(2) = TextBox3.Text
                    add_str_value(3) = ComboBox1.Text
                    If add_str_value(3) = "매번 생성시" Then add_str_value(1) = ""
                    add_str_value(4) = CheckBox1.Checked.ToString
                    add_str_value(15) = ComboBox4.Text
                    add_str_keyname(18) = ComboBox7.Text
                Else
                    add_str_value(0) = ""
                    add_str_value(1) = ""
                    add_str_value(2) = ""
                    add_str_value(3) = ""
                    add_str_value(4) = CheckBox3.Checked.ToString
                    add_str_value(15) = ""
                    add_str_keyname(18) = ""
                End If

            Case 2
                If CheckBox2.Checked.ToString = True Then
                    add_str_value(5) = TextBox4.Text
                    add_str_value(6) = TextBox5.Text
                    add_str_value(7) = TextBox6.Text
                    add_str_value(8) = ComboBox2.Text
                    If add_str_value(8) = "매번 생성시" Then add_str_value(6) = ""
                    add_str_value(9) = CheckBox2.Checked.ToString
                    add_str_value(16) = ComboBox5.Text
                    add_str_keyname(19) = ComboBox8.Text
                Else
                    add_str_value(5) = ""
                    add_str_value(6) = ""
                    add_str_value(7) = ""
                    add_str_value(8) = ""
                    add_str_value(9) = CheckBox3.Checked.ToString
                    add_str_value(16) = ""
                    add_str_keyname(19) = ""
                End If

            Case 3
                If CheckBox3.Checked.ToString = True Then
                    add_str_value(10) = TextBox7.Text
                    add_str_value(11) = TextBox8.Text
                    add_str_value(12) = TextBox9.Text
                    add_str_value(13) = ComboBox3.Text
                    If add_str_value(13) = "매번 생성시" Then add_str_value(12) = ""
                    add_str_value(14) = CheckBox3.Checked.ToString
                    add_str_value(17) = ComboBox4.Text
                    add_str_keyname(20) = ComboBox9.Text
                Else
                    add_str_value(10) = ""
                    add_str_value(11) = ""
                    add_str_value(12) = ""
                    add_str_value(13) = ""
                    add_str_value(14) = CheckBox3.Checked.ToString
                    add_str_value(17) = ""
                    add_str_keyname(20) = ""
                End If

        End Select


    End Sub



    Function Restore_str(str As String) As String
        Dim origin_str As String
        origin_str = str
        Return origin_str
    End Function

    Function GetINIValue(lpApplicationName As String, lpKeyName As String, lpFileName As String) As String

        Dim INI_Return As Long, nSize As Long
        Dim lpReturnedString As String
        Dim lpDefault As String


        nSize = 255
        lpReturnedString = Space(nSize)
        lpDefault = ""
        INI_Return = GPPS(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)

        lpReturnedString = Trim$(lpReturnedString)
        lpReturnedString = lpReturnedString.Substring(0, Len(lpReturnedString) - 1)

        GetINIValue = lpReturnedString
    End Function ' GetINIValue

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox7.Text = "텍스트" Then
            If ComboBox1.Text = "매번 생성시" Then
                TextBox2.Enabled = False
            Else
                TextBox2.Enabled = True
            End If
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox8.Text = "텍스트" Then
            If ComboBox2.Text = "매번 생성시" Then
                TextBox5.Enabled = False
            Else
                TextBox5.Enabled = True
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox9.Text = "텍스트" Then
            If ComboBox3.Text = "매번 생성시" Then
                TextBox8.Enabled = False
            Else
                TextBox8.Enabled = True
            End If
        End If

    End Sub


    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Select Case ComboBox7.Text
            Case "텍스트"
                TextBox2.Enabled = True
                ComboBox1.Enabled = True
                TextBox2.Text = ""

            Case "날짜"

                TextBox2.Text = "&날짜"
                TextBox2.Enabled = False

                ComboBox1.Enabled = False
                ComboBox1.Text = "자동기입"
            Case "시간"
                TextBox2.Enabled = False
                TextBox2.Text = "&시간"

                ComboBox1.Enabled = False
                ComboBox1.Text = "자동기입"
            Case "날짜 + 시간"
                TextBox2.Enabled = False
                TextBox2.Text = "&날짜 + 시간"

                ComboBox1.Enabled = False
                ComboBox1.Text = "자동기입"
        End Select

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        Select Case ComboBox8.Text
            Case "텍스트"
                TextBox5.Enabled = True
                ComboBox2.Enabled = True
                TextBox5.Text = ""

            Case "날짜"

                TextBox5.Text = "&날짜"
                TextBox5.Enabled = False

                ComboBox2.Enabled = False
                ComboBox2.Text = "자동기입"
            Case "시간"
                TextBox5.Enabled = False
                TextBox5.Text = "&시간"

                ComboBox2.Enabled = False
                ComboBox2.Text = "자동기입"
            Case "날짜 + 시간"
                TextBox5.Enabled = False
                TextBox5.Text = "&날짜 + 시간"

                ComboBox2.Enabled = False
                ComboBox2.Text = "자동기입"
        End Select
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Select Case ComboBox9.Text
            Case "텍스트"
                TextBox8.Enabled = True
                ComboBox3.Enabled = True
                TextBox8.Text = ""

            Case "날짜"

                TextBox8.Text = "&날짜"
                TextBox8.Enabled = False

                ComboBox3.Enabled = False
                ComboBox3.Text = "자동기입"
            Case "시간"
                TextBox8.Enabled = False
                TextBox8.Text = "&시간"

                ComboBox3.Enabled = False
                ComboBox3.Text = "자동기입"
            Case "날짜 + 시간"
                TextBox8.Enabled = False
                TextBox8.Text = "&날짜 + 시간"

                ComboBox3.Enabled = False
                ComboBox3.Text = "자동기입"
        End Select
    End Sub
End Class