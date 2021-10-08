Public Class 위치지정_추가기입
    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Const MRM_root_dir As String = "C:\MitutoyoApp"

    Dim add_Str_section As String
    Public add_str_keyname(20) As String
    Public add_str_value(20) As String
    Dim ini_dir As String

    Dim origin_width As Integer = 360

    Dim panel_count As Integer



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

        Select Case panel_num
            Case 1
                If CheckBox1.Checked.ToString = True Then
                    add_str_value(0) = TextBox1.Text
                    add_str_value(1) = TextBox2.Text
                    add_str_value(2) = TextBox3.Text
                    add_str_value(3) = ComboBox1.Text
                    If add_str_value(3) = "매번 생성시" Then add_str_value(1) = ""
                    add_str_value(4) = CheckBox1.Checked.ToString

                Else
                    add_str_value(0) = ""
                    add_str_value(1) = ""
                    add_str_value(2) = ""
                    add_str_value(3) = ""
                    add_str_value(4) = CheckBox3.Checked.ToString
                End If

            Case 2
                If CheckBox2.Checked.ToString = True Then
                    add_str_value(5) = TextBox4.Text
                    add_str_value(6) = TextBox5.Text
                    add_str_value(7) = TextBox6.Text
                    add_str_value(8) = ComboBox2.Text
                    If add_str_value(8) = "매번 생성시" Then add_str_value(6) = ""
                    add_str_value(9) = CheckBox2.Checked.ToString

                Else
                    add_str_value(5) = ""
                    add_str_value(6) = ""
                    add_str_value(7) = ""
                    add_str_value(8) = ""
                    add_str_value(9) = CheckBox3.Checked.ToString
                End If

            Case 3
                If CheckBox3.Checked.ToString = True Then
                    add_str_value(10) = TextBox7.Text
                    add_str_value(11) = TextBox8.Text
                    add_str_value(12) = TextBox9.Text
                    add_str_value(13) = ComboBox3.Text
                    If add_str_value(13) = "매번 생성시" Then add_str_value(12) = ""
                    add_str_value(14) = CheckBox3.Checked.ToString

                Else
                    add_str_value(10) = ""
                    add_str_value(11) = ""
                    add_str_value(12) = ""
                    add_str_value(13) = ""
                    add_str_value(14) = CheckBox3.Checked.ToString

                End If

        End Select

    End Sub

    Private Sub 위치지정_추가기입_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim section_num As Integer

        ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & Form1.ListBox1.SelectedItem.ToString & ".ini"


        For section_num = 1 To 3
            add_Str_section = "add_str_" & section_num
            add_str_keyname(0) = "Description"
            add_str_keyname(1) = "value"
            add_str_keyname(2) = "loction"
            add_str_keyname(3) = "combo"
            add_str_keyname(4) = "use_check"
            add_str_keyname(5) = "panel_count"


            add_str_value(0) = GetINIValue(add_Str_section, add_str_keyname(0), Restore_str(ini_dir))
            add_str_value(1) = GetINIValue(add_Str_section, add_str_keyname(1), Restore_str(ini_dir))
            add_str_value(2) = GetINIValue(add_Str_section, add_str_keyname(2), Restore_str(ini_dir))
            add_str_value(3) = GetINIValue(add_Str_section, add_str_keyname(3), Restore_str(ini_dir))
            add_str_value(4) = GetINIValue(add_Str_section, add_str_keyname(4), Restore_str(ini_dir))
            add_str_value(5) = GetINIValue(add_Str_section, add_str_keyname(5), Restore_str(ini_dir))

            Select Case section_num
                Case 1

                    If add_str_value(4) <> "" Then

                        TextBox1.Text = add_str_value(0)
                        TextBox2.Text = add_str_value(1)
                        TextBox3.Text = add_str_value(2)
                        ComboBox1.Text = add_str_value(3)
                        CheckBox1.Checked = add_str_value(4)
                    Else
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        TextBox3.Text = ""
                        ComboBox1.Text = ""
                        CheckBox1.Checked = False
                    End If


                Case 2
                    If add_str_value(4) <> "" Then

                        TextBox4.Text = add_str_value(0)
                        TextBox5.Text = add_str_value(1)
                        TextBox6.Text = add_str_value(2)
                        ComboBox2.Text = add_str_value(3)

                        CheckBox2.Checked = add_str_value(4)

                    Else

                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        ComboBox2.Text = ""
                        CheckBox2.Checked = False

                    End If
                Case 3
                    If add_str_value(4) <> "" Then

                        TextBox7.Text = add_str_value(0)
                        TextBox8.Text = add_str_value(1)
                        TextBox9.Text = add_str_value(2)
                        ComboBox3.Text = add_str_value(3)
                        CheckBox3.Checked = add_str_value(4)


                    Else
                        TextBox7.Text = ""
                        TextBox8.Text = ""
                        TextBox9.Text = ""
                        ComboBox3.Text = ""
                        CheckBox3.Checked = False
                    End If
            End Select

        Next section_num
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
End Class