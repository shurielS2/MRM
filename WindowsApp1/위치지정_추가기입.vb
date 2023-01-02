Public Class 위치지정_추가기입
    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Const MRM_root_dir As String = "C:\MitutoyoApp"

    Dim add_Str_section As String
    Public add_str_keyname(25) As String
    Public add_str_value(25) As String
    Dim ini_dir As String

    Dim origin_width As Integer = 360

    Public tab_names() As String
    Public tab_count As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           '확인 
        Dim i As Integer

        For i = 1 To 3
            ini_input(i)
        Next
        수정창.add_str_count = 1
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

                add_str_value(0) = TextBox1.Text
                add_str_value(1) = TextBox2.Text
                add_str_value(2) = TextBox3.Text
                add_str_value(3) = ComboBox1.Text
                If add_str_value(3) = "매번 생성시" Then add_str_value(1) = ""
                add_str_value(4) = CheckBox1.Checked.ToString
                add_str_value(15) = ComboBox4.Text
                add_str_value(18) = ComboBox7.Text


            Case 2

                add_str_value(5) = TextBox4.Text
                add_str_value(6) = TextBox5.Text
                add_str_value(7) = TextBox6.Text
                add_str_value(8) = ComboBox2.Text
                If add_str_value(8) = "매번 생성시" Then add_str_value(6) = ""
                add_str_value(9) = CheckBox2.Checked.ToString
                add_str_value(16) = ComboBox5.Text
                add_str_value(19) = ComboBox8.Text


            Case 3

                add_str_value(10) = TextBox7.Text
                add_str_value(11) = TextBox8.Text
                add_str_value(12) = TextBox9.Text
                add_str_value(13) = ComboBox3.Text
                If add_str_value(13) = "매번 생성시" Then add_str_value(11) = ""
                add_str_value(14) = CheckBox3.Checked.ToString
                add_str_value(17) = ComboBox4.Text
                add_str_value(20) = ComboBox9.Text

        End Select

    End Sub

    Private Sub 위치지정_추가기입_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim section_num As Integer
        ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & Form1.ListBox1.SelectedItem.ToString & ".ini"


        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        For i = 1 To tab_count          '수정창에서 탭페이지 만들때 넣어줌
            ComboBox4.Items.Add(tab_names(i))
            ComboBox5.Items.Add(tab_names(i))
            ComboBox6.Items.Add(tab_names(i))
        Next

        If 수정창.add_str_count = 0 Then           '0 첫 시작시 로드

            For section_num = 1 To 3
                add_Str_section = "add_str_" & section_num
                add_str_keyname(0) = "Description"
                add_str_keyname(1) = "value"
                add_str_keyname(2) = "loction"
                add_str_keyname(3) = "combo"
                add_str_keyname(4) = "use_check"
                add_str_keyname(5) = "apply_tab"
                add_str_keyname(6) = "input_type"


                add_str_value(0) = GetINIValue(add_Str_section, add_str_keyname(0), Restore_str(ini_dir))
                add_str_value(1) = GetINIValue(add_Str_section, add_str_keyname(1), Restore_str(ini_dir))
                add_str_value(2) = GetINIValue(add_Str_section, add_str_keyname(2), Restore_str(ini_dir))
                add_str_value(3) = GetINIValue(add_Str_section, add_str_keyname(3), Restore_str(ini_dir))
                add_str_value(4) = GetINIValue(add_Str_section, add_str_keyname(4), Restore_str(ini_dir))
                add_str_value(5) = GetINIValue(add_Str_section, add_str_keyname(5), Restore_str(ini_dir))
                add_str_value(6) = GetINIValue(add_Str_section, add_str_keyname(6), Restore_str(ini_dir))

                If add_str_value(4) = "" Then add_str_value(4) = "false"

                Select Case section_num
                    Case 1


                        TextBox1.Text = add_str_value(0)
                        TextBox2.Text = add_str_value(1)
                        TextBox3.Text = add_str_value(2)
                        ComboBox1.Text = add_str_value(3)
                        CheckBox1.Checked = add_str_value(4)
                        ComboBox4.Text = add_str_value(5)
                            ComboBox7.Text = add_str_value(6)

                            Select Case ComboBox7.Text
                                Case "텍스트"
                                    TextBox2.Enabled = True
                                    ComboBox1.Enabled = True


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

                                Case ""
                                    TextBox2.Enabled = True
                                    TextBox2.Text = ""

                                    ComboBox1.Enabled = True
                                    ComboBox1.Text = ""
                            End Select




                    Case 2

                        TextBox4.Text = add_str_value(0)
                        TextBox5.Text = add_str_value(1)
                        TextBox6.Text = add_str_value(2)
                        ComboBox2.Text = add_str_value(3)

                        CheckBox2.Checked = add_str_value(4)
                        ComboBox5.Text = add_str_value(5)
                        ComboBox8.Text = add_str_value(6)
                            Select Case ComboBox8.Text
                                Case "텍스트"
                                    TextBox5.Enabled = True
                                    ComboBox2.Enabled = True
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

                                Case ""
                                    TextBox5.Enabled = True
                                    TextBox5.Text = ""

                                    ComboBox2.Enabled = True
                                    ComboBox2.Text = ""
                            End Select


                    Case 3

                        TextBox7.Text = add_str_value(0)
                        TextBox8.Text = add_str_value(1)
                        TextBox9.Text = add_str_value(2)
                        ComboBox3.Text = add_str_value(3)
                        CheckBox3.Checked = add_str_value(4)
                        ComboBox6.Text = add_str_value(5)
                        ComboBox9.Text = add_str_value(6)

                        Select Case ComboBox9.Text
                                Case "텍스트"
                                    TextBox8.Enabled = True
                                    ComboBox3.Enabled = True


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
                                Case ""
                                    TextBox8.Enabled = True
                                    TextBox8.Text = ""

                                ComboBox3.Enabled = True
                                ComboBox3.Text = ""
                            End Select


                End Select


            Next section_num

        ElseIf 수정창.add_str_count = 1 Then          '1 수정 후 로드
            For section_num = 1 To 3
                Select Case section_num
                    Case 1

                        TextBox1.Text = 수정창.Temp_User_Info(0)
                        TextBox2.Text = 수정창.Temp_User_Info(1)
                        TextBox3.Text = 수정창.Temp_User_Info(2)
                        ComboBox1.Text = 수정창.Temp_User_Info(3)
                        CheckBox1.Checked = 수정창.Temp_User_Info(4)
                        ComboBox4.Text = 수정창.Temp_User_Info(15)

                        ComboBox7.Text = 수정창.Temp_User_Info(18)


                    Case 2

                        TextBox4.Text = 수정창.Temp_User_Info(5)
                        TextBox5.Text = 수정창.Temp_User_Info(6)
                        TextBox6.Text = 수정창.Temp_User_Info(7)
                        ComboBox2.Text = 수정창.Temp_User_Info(8)

                        CheckBox2.Checked = 수정창.Temp_User_Info(9)
                        ComboBox5.Text = 수정창.Temp_User_Info(16)
                        ComboBox8.Text = 수정창.Temp_User_Info(19)

                    Case 3


                        TextBox7.Text = 수정창.Temp_User_Info(10)
                        TextBox8.Text = 수정창.Temp_User_Info(11)
                        TextBox9.Text = 수정창.Temp_User_Info(12)
                        ComboBox3.Text = 수정창.Temp_User_Info(13)
                        CheckBox3.Checked = 수정창.Temp_User_Info(14)
                        ComboBox6.Text = 수정창.Temp_User_Info(17)
                        ComboBox9.Text = 수정창.Temp_User_Info(20)

                End Select

            Next section_num

        End If
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
        If ComboBox1.Text = "매번 생성시" Then
            TextBox2.Enabled = False
        Else
            TextBox2.Enabled = True
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "매번 생성시" Then
            TextBox5.Enabled = False
        Else
            TextBox5.Enabled = True
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "매번 생성시" Then
            TextBox8.Enabled = False
        Else
            TextBox8.Enabled = True
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