Public Class 위치지정_추가기입_new
    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Const MRM_root_dir As String = "C:\MitutoyoApp"

    Dim add_Str_section As String
    Public add_str_keyname(20) As String
    Public add_str_value(20) As String
    Dim ini_dir As String

    Dim origin_width As Integer = 360

    Dim panel_count As Integer


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim i As Integer

        For i = 1 To 3
            ini_input(i)
        Next

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
                    If add_str_value(13) = "매번 생성시" Then add_str_value(11) = ""
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