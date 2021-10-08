Public Class MkNew

    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Const MRM_root_dir As String = "C:\MitutoyoApp"
    Public Logic_value As Integer
    Public Matching_name As String
    Public SaveFile_Dir As String
    Dim ini_section As String
    Dim ini_Dir As String
    Dim ini_KeyName() As String
    Dim ini_Value() As String
    Dim ini_dir_preserve As String

    Dim Result_Section As String
    Dim Result_Keyname() As String
    Dim Result_Value() As String
    Dim check_section As String
    Dim check_value() As String
    Dim check_keyname() As String

    Dim User_Info_Section As String
    Dim User_Info_Keyname() As String
    Dim User_Info_Value() As String
    Dim add_str_section(3) As String
    Dim add_str_keyname() As String
    Dim add_str_value() As String

    Public Temp_User_Info() As String
    Public Temp_User_Info_bitmap As Bitmap



    Private Sub MkNew_Load(sender As Object, e As EventArgs) Handles Me.Load
        RadioButton1.Checked = True
        CheckBox11.Checked = True
        Form1.New_Fix_check = 0

        Me.Location = New Point(Form1.Location.X + 50, Form1.Location.Y + 50)
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ReDim Temp_User_Info(25)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           'csv 파일열기
        Dim ofd_For_Csv As New OpenFileDialog()
        With ofd_For_Csv
            .InitialDirectory = MRM_root_dir & "\MRM\Data"
            .Filter = "CSV(*.CSV)/ASC(*.ASC)|*.csv;*.asc|All File(*.*)|*.*"
            .FilterIndex = 1
            .Title = "Select CSV File "
            .RestoreDirectory = True
            .CheckFileExists = True
            .CheckPathExists = True
        End With

        Dim fFindFolder As New System.IO.DirectoryInfo(ofd_For_Csv.InitialDirectory)
        If fFindFolder.Exists = False Then
            MkDir(MRM_root_dir & "\MRM\data")
        End If

        If ofd_For_Csv.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = ofd_For_Csv.FileName
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click       'save 파일 경로

        Dim save_For_Result As New FolderBrowserDialog()
        save_For_Result.SelectedPath = MRM_root_dir & "\MRM\Result"
        Dim fFindFolder As New System.IO.DirectoryInfo(save_For_Result.SelectedPath)           '폴더 존재 유무 확인
        If fFindFolder.Exists = False Then
            MkDir(MRM_root_dir & "\MRM\Result")
        End If


        save_For_Result.ShowDialog()

        TextBox2.Text = save_For_Result.SelectedPath


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click       '저장
        ReDim ini_KeyName(15)
        ReDim ini_Value(15)

        ReDim check_value(15)
        ReDim check_keyname(15)
        ReDim Result_Keyname(15)
        ReDim Result_Value(15)
        ReDim add_str_keyname(25)
        ReDim add_str_value(25)
        ReDim User_Info_Keyname(15)
        ReDim User_Info_Value(15)


        If TextBox1.Text = "" Then
            MsgBox("매칭할 CSV파일을 선택해 주세요")
            Exit Sub
        End If

        If TextBox2.Text = "" Then                          '성적서 파일 세이브 경로 공백 확인
            MsgBox("파일 경로를 입력해 주세요")
            Exit Sub
        End If

        If ComboBox1.Text = "" Then
            MsgBox("파일 저장 유형을 선택해 주세요")
            Exit Sub
        End If

        If TextBox3.Text = "" Then              '매칭프로그램 이름 공백 확인용
            MsgBox("매칭 이름을 적어주세요")
            Exit Sub
        End If

        Matching_name = TextBox3.Text           '매칭 리스트 이름

        If TextBox4.Text = "" Then                      '성적서 파일 공백 확인
            MsgBox("성적서 파일 이름을 입력해주세요")
            Exit Sub
        End If

        'SaveFile_Dir = TextBox2.Text & "\" & TextBox3.Text & ComboBox1.Text          '파일 세이브 경로(이름.확장자 포함)

        Dim fFindDir As New System.IO.DirectoryInfo(MRM_root_dir & "\MRM\Data\Resources\ini")
        If fFindDir.Exists = False Then

            MkDir(MRM_root_dir & "\MRM\Data\Resources\ini")
        End If

        'Logic_value = 1                         '확인버튼 누른후 form1에서 판단용
        'MsgBox(SaveFile_Dir)               '경로 확인용

        ini_Dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & Matching_name & ".ini"                    'ini 파일 경로
        ini_dir_preserve = ini_Dir

        '============================================================================ini 값입력 & 생성
        ini_section = "Matching_Info"

        ini_KeyName(0) = "CSV_file_Path"
        ini_KeyName(1) = "Save_File_Path"
        ini_KeyName(2) = "Save_File_Name"
        ini_KeyName(3) = "Save_Type"
        ini_KeyName(4) = "ini_dir"
        ini_KeyName(5) = "Result_Form"
        ini_KeyName(6) = "Last_Paly_date"
        ini_KeyName(7) = "check_date"
        ini_KeyName(8) = "check_time"
        ini_KeyName(9) = "auto_save"
        ini_KeyName(10) = "Basic_form_seleted"

        Result_Section = "custom_match_info"
        Result_Keyname(0) = "label"
        Result_Keyname(1) = "measure_value"
        Result_Keyname(2) = "Design_value"
        Result_Keyname(3) = "error"
        Result_Keyname(4) = "UP_tol"
        Result_Keyname(5) = "Low_tol"
        Result_Keyname(6) = "judge"
        Result_Keyname(7) = "line_count"
        Result_Keyname(8) = "Result_Form_Dir"
        Result_Keyname(9) = "component"
        Result_Keyname(10) = "input_direction"

        check_section = "check"

        check_keyname(0) = "label"
        check_keyname(1) = "measure_value"
        check_keyname(2) = "Design_value"
        check_keyname(3) = "error"
        check_keyname(4) = "UP_tol"
        check_keyname(5) = "Low_tol"
        check_keyname(6) = "judge"
        check_keyname(7) = "component"

        User_Info_Section = "User_info"

        User_Info_Keyname(0) = "Product_Name"
        User_Info_Keyname(1) = "Machine_Name"
        User_Info_Keyname(2) = "Request_Dept"
        User_Info_Keyname(3) = "Request_Date"
        User_Info_Keyname(4) = "Drawing_Num"
        User_Info_Keyname(5) = "Program_Name"
        User_Info_Keyname(6) = "Player_Name"
        User_Info_Keyname(7) = "Measure_Date"
        User_Info_Keyname(8) = "Check_date_1"
        User_Info_Keyname(9) = "Check_date_2"
        User_Info_Keyname(10) = "Select_pic_name"
        User_Info_Keyname(11) = "Select_pic_name_2"
        User_Info_Keyname(12) = "Select_pic_name_3"


        add_str_section(0) = "add_str_1"
        add_str_section(1) = "add_str_2"
        add_str_section(2) = "add_str_3"

        add_str_keyname(0) = "Description"
        add_str_keyname(1) = "value"
        add_str_keyname(2) = "loction"
        add_str_keyname(3) = "combo"
        add_str_keyname(4) = "use_check"

        '============================================================================ini 값입력 & 생성

        ini_Value(0) = TextBox1.Text        'csv파일 경로
        ini_Value(1) = TextBox2.Text      '성적서 저장 경로(이름.확장자 제외)
        ini_Value(2) = TextBox4.Text       '성적서 저장 이름
        ini_Value(3) = ComboBox1.Text       '성적서 저장 타입
        ini_Value(4) = ini_Dir          'ini 저장경로

        If RadioButton1.Checked = True Then
            ini_Value(5) = RadioButton1.Text        '그룹박스 선택 번호 혹은 스트링
            ini_Value(10) = TabControl1.SelectedTab.Text
        Else
            ini_Value(5) = RadioButton2.Text      '그룹박스 선택 번호 혹은 스트링
            ini_Value(10) = "위치 지정"
        End If

        ini_Value(6) = ""                '현재 시간
        ini_Value(7) = CheckBox1.Checked.ToString   '이름 날짜 체크
        ini_Value(8) = CheckBox2.Checked.ToString   ' 이름 시간 체크
        ini_Value(9) = CheckBox11.Checked.ToString  '자동저장 유무 체크

        '============================================================================

        Dim fFindFile As New System.IO.FileInfo(ini_Dir)             'ini 존재 여부 확인용
        If fFindFile.Exists = True Then             'ini 존재 여부 확인용
            MsgBox("동일이름의 파일이 존재 합니다." & Environment.NewLine & "이름을 다시 지정해주세요.")
            Exit Sub
        Else
            Form1.ListBox1.Items.Add(Matching_name)                 '리스트 박스에 목록 추가 
        End If

        '============================================================================
        '매칭 성적서 폼 기본 <> 위치 지정
        '============================================================================

        If RadioButton2.Checked = True Then             '위치 지정

            Result_Value(0) = TextBox5.Text
            Result_Value(1) = TextBox6.Text
            Result_Value(2) = TextBox7.Text
            Result_Value(3) = TextBox8.Text
            Result_Value(4) = TextBox9.Text
            Result_Value(5) = TextBox10.Text
            Result_Value(6) = TextBox11.Text
            Result_Value(7) = TextBox12.Text
            Result_Value(9) = TextBox14.Text
            Result_Value(10) = ComboBox2.Text
            Result_Value(8) = TextBox13.Text


            check_value(0) = CheckBox3.Checked
            check_value(1) = CheckBox4.Checked
            check_value(2) = CheckBox5.Checked
            check_value(3) = CheckBox6.Checked
            check_value(4) = CheckBox7.Checked
            check_value(5) = CheckBox8.Checked
            check_value(6) = CheckBox9.Checked
            check_value(7) = CheckBox10.Checked

            Dim p As Integer
            For p = 0 To 7
                If check_value(p) = True Then
                    If Result_Value(p) = "" Then
                        MsgBox("셀 주소에 공백이 있습니다. 셀 주소를 확인해 주세요",, "셀 주소 기입오류")
                        Exit Sub
                    End If
                End If
            Next


            If Result_Value(7) = "" Then        '페이지당 줄수 
                MsgBox("페이지당 줄수는 필수 입니다. 기입해 주세요",, "페이지 줄 수 기입오류")
                Exit Sub
            End If

            If Result_Value(8) = "" Then         '원본 성적서 선택 유무 확인 
                MsgBox("원본 성적서 폼을 선택해 주세요",, "원본 성적서 선택 오류")
                Exit Sub
            End If

            WPPS(Result_Section, Result_Keyname(0), Result_Value(0), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(1), Result_Value(1), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(2), Result_Value(2), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(3), Result_Value(3), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(4), Result_Value(4), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(5), Result_Value(5), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(6), Result_Value(6), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(7), Result_Value(7), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(8), Result_Value(8), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(9), Result_Value(9), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(10), Result_Value(10), Restore_str(ini_Dir))


            WPPS(check_section, check_keyname(0), check_value(0), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(1), check_value(1), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(2), check_value(2), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(3), check_value(3), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(4), check_value(4), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(5), check_value(5), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(6), check_value(6), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(7), check_value(7), Restore_str(ini_Dir))

            '========================================================================
            '위치지정시 기본 유저 정보 공백 처리
            '========================================================================
            User_Info_Value(0) = ""
            User_Info_Value(1) = ""
            User_Info_Value(2) = ""
            User_Info_Value(3) = ""
            User_Info_Value(4) = ""
            User_Info_Value(5) = ""
            User_Info_Value(6) = ""
            User_Info_Value(7) = ""
            User_Info_Value(8) = ""
            User_Info_Value(9) = ""
            User_Info_Value(10) = ""
            User_Info_Value(11) = ""
            User_Info_Value(12) = ""

            WPPS(User_Info_Section, User_Info_Keyname(0), User_Info_Value(0), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(1), User_Info_Value(1), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(2), User_Info_Value(2), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(3), User_Info_Value(3), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(4), User_Info_Value(4), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(5), User_Info_Value(5), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(6), User_Info_Value(6), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(7), User_Info_Value(7), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(8), User_Info_Value(8), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(9), User_Info_Value(9), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(10), User_Info_Value(10), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(11), User_Info_Value(11), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(12), User_Info_Value(12), Restore_str(ini_Dir))


            add_str_value(0) = Temp_User_Info(0)
            add_str_value(1) = Temp_User_Info(1)
            add_str_value(2) = Temp_User_Info(2)
            add_str_value(3) = Temp_User_Info(3)
            add_str_value(4) = Temp_User_Info(4)
            add_str_value(5) = Temp_User_Info(5)
            add_str_value(6) = Temp_User_Info(6)
            add_str_value(7) = Temp_User_Info(7)
            add_str_value(8) = Temp_User_Info(8)
            add_str_value(9) = Temp_User_Info(9)
            add_str_value(10) = Temp_User_Info(10)
            add_str_value(11) = Temp_User_Info(11)
            add_str_value(12) = Temp_User_Info(12)
            add_str_value(13) = Temp_User_Info(13)
            add_str_value(14) = Temp_User_Info(14)
            add_str_value(15) = Temp_User_Info(15)
            add_str_value(16) = Temp_User_Info(16)


            WPPS(add_str_section(0), add_str_keyname(0), add_str_value(0), Restore_str(ini_Dir))
            WPPS(add_str_section(0), add_str_keyname(1), add_str_value(1), Restore_str(ini_Dir))
            WPPS(add_str_section(0), add_str_keyname(2), add_str_value(2), Restore_str(ini_Dir))
            WPPS(add_str_section(0), add_str_keyname(3), add_str_value(3), Restore_str(ini_Dir))
            WPPS(add_str_section(0), add_str_keyname(4), add_str_value(4), Restore_str(ini_Dir))
            WPPS(add_str_section(1), add_str_keyname(0), add_str_value(5), Restore_str(ini_Dir))
            WPPS(add_str_section(1), add_str_keyname(1), add_str_value(6), Restore_str(ini_Dir))
            WPPS(add_str_section(1), add_str_keyname(2), add_str_value(7), Restore_str(ini_Dir))
            WPPS(add_str_section(1), add_str_keyname(3), add_str_value(8), Restore_str(ini_Dir))
            WPPS(add_str_section(1), add_str_keyname(4), add_str_value(9), Restore_str(ini_Dir))
            WPPS(add_str_section(2), add_str_keyname(0), add_str_value(10), Restore_str(ini_Dir))
            WPPS(add_str_section(2), add_str_keyname(1), add_str_value(11), Restore_str(ini_Dir))
            WPPS(add_str_section(2), add_str_keyname(2), add_str_value(12), Restore_str(ini_Dir))
            WPPS(add_str_section(2), add_str_keyname(3), add_str_value(13), Restore_str(ini_Dir))
            WPPS(add_str_section(2), add_str_keyname(4), add_str_value(14), Restore_str(ini_Dir))



        Else    '기본 선택시 ini에 위치선택 속성 공백 지정하여 넣기

            Result_Value(0) = ""
            Result_Value(1) = ""
            Result_Value(2) = ""
            Result_Value(3) = ""
            Result_Value(4) = ""
            Result_Value(5) = ""
            Result_Value(6) = ""
            Result_Value(7) = ""
            Result_Value(9) = ""

            Result_Value(8) = ""

            check_value(0) = False
            check_value(1) = False
            check_value(2) = False
            check_value(3) = False
            check_value(4) = False
            check_value(5) = False
            check_value(6) = False
            check_value(7) = False

            WPPS(Result_Section, Result_Keyname(0), Result_Value(0), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(1), Result_Value(1), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(2), Result_Value(2), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(3), Result_Value(3), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(4), Result_Value(4), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(5), Result_Value(5), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(6), Result_Value(6), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(7), Result_Value(7), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(8), Result_Value(8), Restore_str(ini_Dir))
            WPPS(Result_Section, Result_Keyname(9), Result_Value(9), Restore_str(ini_Dir))

            WPPS(check_section, check_keyname(0), check_value(0), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(1), check_value(1), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(2), check_value(2), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(3), check_value(3), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(4), check_value(4), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(5), check_value(5), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(6), check_value(6), Restore_str(ini_Dir))
            WPPS(check_section, check_keyname(7), check_value(7), Restore_str(ini_Dir))

            '========================================================================
            '기본 서식 사용시 기본 유저 정보 기입
            '========================================================================
            For i = 0 To 7
                If Temp_User_Info(i) = Nothing Then
                    Temp_User_Info(i) = ""
                End If
            Next i

            User_Info_Value(0) = Temp_User_Info(0)
            User_Info_Value(1) = Temp_User_Info(1)
            User_Info_Value(2) = Temp_User_Info(2)
            User_Info_Value(3) = Temp_User_Info(3)
            User_Info_Value(4) = Temp_User_Info(4)
            User_Info_Value(5) = Temp_User_Info(5)
            User_Info_Value(6) = Temp_User_Info(6)
            User_Info_Value(7) = Temp_User_Info(7)
            User_Info_Value(8) = Temp_User_Info(8)
            User_Info_Value(9) = Temp_User_Info(9)
            User_Info_Value(10) = Temp_User_Info(10)
            User_Info_Value(11) = Temp_User_Info(11)
            User_Info_Value(12) = Temp_User_Info(12)

            WPPS(User_Info_Section, User_Info_Keyname(0), User_Info_Value(0), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(1), User_Info_Value(1), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(2), User_Info_Value(2), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(3), User_Info_Value(3), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(4), User_Info_Value(4), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(5), User_Info_Value(5), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(6), User_Info_Value(6), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(7), User_Info_Value(7), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(8), User_Info_Value(8), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(9), User_Info_Value(9), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(10), User_Info_Value(10), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(11), User_Info_Value(11), Restore_str(ini_Dir))
            WPPS(User_Info_Section, User_Info_Keyname(12), User_Info_Value(12), Restore_str(ini_Dir))
        End If

        '==============================================================기본 ini 내용 작성
        WPPS(ini_section, ini_KeyName(0), ini_Value(0), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(1), ini_Value(1), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(2), ini_Value(2), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(3), ini_Value(3), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(4), ini_Value(4), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(5), ini_Value(5), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(6), ini_Value(6), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(7), ini_Value(7), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(8), ini_Value(8), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(9), ini_Value(9), Restore_str(ini_Dir))
        WPPS(ini_section, ini_KeyName(10), ini_Value(10), Restore_str(ini_Dir))
        '==============================================================기본 ini 내용 작성



        Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click       '취소
        Logic_value = 2
        MsgBox("성적서 매칭 정보 생성을 취소했습니다.")
        Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked Then

            TabControl1.Visible = True

            Button6.Text = "머릿말 정보 입력"
            Button6.Visible = True
            Panel1.Visible = False

        Else                ' 위치지정

            TabControl1.Visible = False

            Button6.Text = "추가 데이터 입력"
            Button6.Visible = True
            Panel1.Visible = True
        End If

        '   If Button6.Visible = False Then
        '   Button6.Visible = True
        '   Button6.Enabled = True
        '   End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = False Then
            TextBox5.Enabled = False
        Else
            TextBox5.Enabled = True
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = False Then
            TextBox6.Enabled = False
        Else
            TextBox6.Enabled = True
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = False Then
            TextBox7.Enabled = False
        Else
            TextBox7.Enabled = True
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = False Then
            TextBox8.Enabled = False
        Else
            TextBox8.Enabled = True
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = False Then
            TextBox9.Enabled = False
        Else
            TextBox9.Enabled = True
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = False Then
            TextBox10.Enabled = False
        Else
            TextBox10.Enabled = True
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = False Then
            TextBox11.Enabled = False
        Else
            TextBox11.Enabled = True
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim ofd_For_resultForm As New OpenFileDialog

        With ofd_For_resultForm
            .InitialDirectory = MRM_root_dir & "\MRM\Data\Resources"
            .Filter = "Xlsx(*.xlsx)|*.xlsx|All File(*.*)|*.*"
            .FilterIndex = 1
            .Title = "Select Result Form"
            .RestoreDirectory = True
            .CheckFileExists = True
            .CheckPathExists = True
        End With
        If ofd_For_resultForm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox13.Text = ofd_For_resultForm.FileName
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = False Then
            TextBox14.Enabled = False
        Else
            TextBox14.Enabled = True
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If CheckBox3.Checked = True Then
            TextBox5.Enabled = True
        End If
        If CheckBox4.Checked = True Then
            TextBox6.Enabled = True
        End If
        If CheckBox5.Checked = True Then
            TextBox7.Enabled = True
        End If
        If CheckBox6.Checked = True Then
            TextBox8.Enabled = True
        End If
        If CheckBox7.Checked = True Then
            TextBox9.Enabled = True
        End If
        If CheckBox8.Checked = True Then
            TextBox10.Enabled = True
        End If
        If CheckBox9.Checked = True Then
            TextBox11.Enabled = True
        End If
        If CheckBox10.Checked = True Then
            TextBox14.Enabled = True
        End If

        ' If Button6.Visible = True Then
        'Button6.Visible = False
        'Button6.Enabled = False
        'End If

    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If RadioButton1.Checked = True Then
            Select Case TabControl1.SelectedTab.Text
                Case "기본폼1"

                    Form7.ShowDialog()

                    Temp_User_Info(0) = Form7.Product_Name
                    Temp_User_Info(1) = Form7.Machine_Name
                    Temp_User_Info(2) = Form7.Request_Dept
                    Temp_User_Info(3) = Form7.Request_Date
                    Temp_User_Info(4) = Form7.Drawing_Num
                    Temp_User_Info(5) = Form7.Program_Name
                    Temp_User_Info(6) = Form7.Player_Name
                    Temp_User_Info(7) = Form7.Measure_Date
                    Temp_User_Info(8) = Form7.Check_date_1
                    Temp_User_Info(9) = Form7.Check_date_2
                    Temp_User_Info(10) = Form7.select_pic_name
                    Temp_User_Info(11) = Form7.select_pic_name_2
                    Temp_User_Info(12) = ""

                Case "기본폼2"

                    Form9.ShowDialog()

                    Temp_User_Info(0) = Form9.Product_Name
                    Temp_User_Info(1) = Form9.Machine_Name
                    Temp_User_Info(2) = Form9.Request_Dept
                    Temp_User_Info(3) = Form9.Request_Date
                    Temp_User_Info(4) = Form9.Drawing_Num
                    Temp_User_Info(5) = Form9.Program_Name
                    Temp_User_Info(6) = Form9.Player_Name
                    Temp_User_Info(7) = Form9.Measure_Date
                    Temp_User_Info(8) = Form9.Check_date_1
                    Temp_User_Info(9) = Form9.Check_date_2
                    Temp_User_Info(10) = ""
                    Temp_User_Info(11) = ""
                    Temp_User_Info(12) = ""
                Case "기본폼3"

                    Form10.ShowDialog()

                    Temp_User_Info(0) = Form10.Product_Name
                    Temp_User_Info(1) = Form10.Machine_Name
                    Temp_User_Info(2) = Form10.Request_Dept
                    Temp_User_Info(3) = Form10.Request_Date
                    Temp_User_Info(4) = Form10.Drawing_Num
                    Temp_User_Info(5) = Form10.Program_Name
                    Temp_User_Info(6) = Form10.Player_Name
                    Temp_User_Info(7) = Form10.Measure_Date
                    Temp_User_Info(8) = Form10.Check_date_1
                    Temp_User_Info(9) = Form10.Check_date_2
                    Temp_User_Info(10) = Form10.select_pic_name
                    Temp_User_Info(11) = Form10.select_pic_name_2
                    Temp_User_Info(12) = Form10.select_pic_name_3

                Case "기본폼4"

                    Form7.ShowDialog()

                    Temp_User_Info(0) = Form7.Product_Name
                    Temp_User_Info(1) = Form7.Machine_Name
                    Temp_User_Info(2) = Form7.Request_Dept
                    Temp_User_Info(3) = Form7.Request_Date
                    Temp_User_Info(4) = Form7.Drawing_Num
                    Temp_User_Info(5) = Form7.Program_Name
                    Temp_User_Info(6) = Form7.Player_Name
                    Temp_User_Info(7) = Form7.Measure_Date
                    Temp_User_Info(8) = Form7.Check_date_1
                    Temp_User_Info(9) = Form7.Check_date_2
                    Temp_User_Info(10) = Form7.select_pic_name
                    Temp_User_Info(11) = Form7.select_pic_name_2
                    Temp_User_Info(12) = ""

                Case "기본폼5"

                    Form10.ShowDialog()

                    Temp_User_Info(0) = Form10.Product_Name
                    Temp_User_Info(1) = Form10.Machine_Name
                    Temp_User_Info(2) = Form10.Request_Dept
                    Temp_User_Info(3) = Form10.Request_Date
                    Temp_User_Info(4) = Form10.Drawing_Num
                    Temp_User_Info(5) = Form10.Program_Name
                    Temp_User_Info(6) = Form10.Player_Name
                    Temp_User_Info(7) = Form10.Measure_Date
                    Temp_User_Info(8) = Form10.Check_date_1
                    Temp_User_Info(9) = Form10.Check_date_2
                    Temp_User_Info(10) = Form10.select_pic_name
                    Temp_User_Info(11) = Form10.select_pic_name_2
                    Temp_User_Info(12) = Form10.select_pic_name_3
            End Select

        Else

            위치지정_추가기입_new.ShowDialog()
            Temp_User_Info(0) = 위치지정_추가기입_new.add_str_value(0)
            Temp_User_Info(1) = 위치지정_추가기입_new.add_str_value(1)
            Temp_User_Info(2) = 위치지정_추가기입_new.add_str_value(2)
            Temp_User_Info(3) = 위치지정_추가기입_new.add_str_value(3)
            Temp_User_Info(4) = 위치지정_추가기입_new.add_str_value(4)
            Temp_User_Info(5) = 위치지정_추가기입_new.add_str_value(5)
            Temp_User_Info(6) = 위치지정_추가기입_new.add_str_value(6)
            Temp_User_Info(7) = 위치지정_추가기입_new.add_str_value(7)
            Temp_User_Info(8) = 위치지정_추가기입_new.add_str_value(8)
            Temp_User_Info(9) = 위치지정_추가기입_new.add_str_value(9)
            Temp_User_Info(10) = 위치지정_추가기입_new.add_str_value(10)
            Temp_User_Info(11) = 위치지정_추가기입_new.add_str_value(11)
            Temp_User_Info(12) = 위치지정_추가기입_new.add_str_value(12)
            Temp_User_Info(13) = 위치지정_추가기입_new.add_str_value(13)
            Temp_User_Info(14) = 위치지정_추가기입_new.add_str_value(14)





        End If

    End Sub
    Function Restore_str(str As String) As String
        Dim origin_str As String
        origin_str = str
        Return origin_str
    End Function


End Class