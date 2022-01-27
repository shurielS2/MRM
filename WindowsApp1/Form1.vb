'===============================================================어셈블리 버전 갱신용
'2020-07-02 ver 1.0.1       기본성적서 이외에 위치지정 성적서 폼 라디오 버튼 및 속성 추가
'2020-07-14 ver 1.0.2       form3 -리스트 수정시 form1 리스트 재로드 수정
'2020-07-14 ver 1.0.2       위치지정 성적서의 베이스가 될 성적서 파일 지정 추가
'2020-07-15 ver 1.0.3       성적서 save folder 위치 지정시 지정 폴더가 없으면 생성 기능 추가
'2020-07-15 ver 1.0.3       기본 성적서폼 선택시 위치지정 속성 공백 지정해서 저장하는 기능 추가
'2020-07-15 ver 1.0.4       위치지정 성적서 생성 루틴 추가 
'2020-07-15 ver 1.0.5       줄 수 지정 숫자만 입력하는 이벤트 추가
'2020-07-15 ver 1.0.6       Form1 셀 주소 지정 영문과 숫자 분리 (ad_NUM, ad_str 함수)
'2020-07-15 ver 1.0.7       수정 로드시 체크박스 상태에 따라서 텍스트 박스 비활성화
'2020-07-15 ver 1.0.8       자동저장 안할시에 임시파일로 백업본 저장
'2020-07-16 ver 1.0.9       기본 성적서 없을때 자동 생성시 오류 workbook >workbooks 수정
'2020-07-16 ver 1.0.10      위치지정 loop문 line_count 스텍 누락 추가
'2020-07-16 ver 1.0.11      성적서 생성, 수정 시 공백 확인 누락된것 추가 (csv 파일, 저장 유형)
'2020-08-13 Ver 1.0.12      프로그래스바 추가 
'2020-08-13 Ver 1.0.13      성적서 자동저장 체크박스 추가. -기본으로는 자동저장
'2020-08-13 Ver 1.0.14      A9:I36 범위 shrinktofit = true 설정 (셀맞춤 기능)
'2020-08-18 Ver 1.0.15      파일 경로 오류 수정 (False <> false 문제)
'2020-08-18 Ver 1.0.16      상위 폴더 경로 (MRM) 추가
'2020-08-19 Ver 1.0.17      복사방지 추가 & 액티브 키(프로그램) 제작
'2020-08-20 Ver 1.0.18      3차원 측정기용 ASC 파일 매칭 추가
'2020-08-20 Ver 1.0.19      fileopen - input 구현 방식으로 변경
'2020-08-20 Ver 1.0.20      중복실행 방지 구문 추가
'2020-08-21 Ver 1.0.21      수정시 리스트박스 리로드 구문 추가
'2020-08-21 Ver 1.0.22      contexstrip 추가 - 경로 폴더 열기
'2020-08-24 Ver 1.0.23      판정 사용자서식 추가
'2020-08-25 Ver 1.0.24      Extention_type3 덤핑 루틴 수정 덤핑 입력후 ""(공백)으로 처리
'2020-09-24 Ver 1.0.25      기본성적서 유저정보 입력 버튼 생성
'2020-09-24 Ver 1.0.26      유저정보 입력 버그수정 
'2020-09-25 Ver 1.0.27      유저정보 입력 세부 수정
'2020-10-07 Ver 1.0.28      유저정보 입력 날자 지정 체크박스 기능 추가
'2020-10-29 Ver 1.0.29      유저정보 성적서 사진 변경 기능 추가 , 페이지 자동입력 추가
'2020-11-04 Ver 1.0.30      PDF저장 기능 추가 , PDF저장시 temp파일 생성 추가
'2020-11-20 Ver 1.0.31      설정 하나 골라서 완전 자동화 실행 방법 추가 - ini폴더안에 실행파일과 동일한 이름을 가진 폴더 존재하면 해당폴더 안의 ini 파일 하나 읽어와서 바로 실행
'2020-11-23 Ver 1.0.32      텍스트 박스에서 키코드 이벤트 Form1, Form8에 추가
'2020-12-18 Ver 1.0.33      ini파일 경로 한글 사용하기위해 dir를 매번 초기값으로 반환하는 함수 적용
'2021-01-05 Ver 1.0.34      매칭 삭제시 확인 메세지 추가, QV-CSV에서 형상 공차 기입 위치 '오차 -> 설계값' 으로 자동 변경 구문추가
'2021-01-07 Ver 1.0.35      기본폼 3개 추가 조합폼 5개 추가하여 기본 제공 폼 5개 적용. 탭컨트롤 추가하여 각 탭마다 기본폼 선택 적용, 각 성적서 조합 마다 삽입 위치 조정
'2021-01-11 Ver 1.0.36      각 기본폼 마다 유저 기본정보 입력창 내용 공유 구문 정리.
'2021-01-15 Ver 1.0.37      Keycode event 추가 : delete키 - 삭제 버튼 클릭 추가
'2021-02-03 Ver 1.0.38      메인폼 로드시 MRM 폴더 유무 확인후 폴더 없으면 기본 폴더 생성하는 구문 추가 / 활성화 방법 변경 - 레지스트리 등록 방법 > mysettings 값으로 등록 하기. 복사후 컴퓨터 옮기면 재등록 필요
'2020-02-04 Ver 1.0.39      사용자 매뉴얼 프로그램 리소스에 추가, 다운로드 기능 추가
'2021-03-02 Ver 1.0.40      샘플 사진 제거 및 우측 위 미쓰도요 로고 >  업체 사진 기입 가능하도록 수정
'2021-03-04 Ver 1.0.41      사진 변경 삭제 선택폼 추가, 미세 버그 수정
'2021-03-10 Ver 1.0.42      메인 화면 폼 크기 고정 및 Label 자동 줄변경 추가
'2021-03-15 Ver 1.0.43      소프트웨어 설치유무 레지스트리 판단하여 활성화 제한 (소프트웨어 설치 컴퓨터에서만 활성화 가능하게 하기)
'2021-03-16 Ver 1.0.44      루트 directory 지정하여 어디서 열든 어떤 소프웨어와 연동하든 하나의 경로에서 설정파일 존재하도록 하기. curdir() -> MRM_root_dir 변경
'2021-03-16 Ver 1.0.45      활성화 성공적으로 한후 소프트웨어(QVPAK,MCOSMOS) 삭제 했을때 오류 문구 추가
'2021-03-30 Ver 1.0.46      전용프로그램 폴더 지정 후 경로 변경
'2021-04-02 Ver 1.0.47      QV배열 측정시Split기능사용 할때 공백 부분 인덱스 처리 limit 값을  10>-1, 11>-1로 변경 공백 부분도 배열 생성해줌 
'2021-04-06 Ver 1.0.48      활성화시에 QVClient 레지스트리 등록 및 root Dirtory에 MRM 프로그램 복사 -> QVClient 실행용 
'2021-04-16 Ver 1.0.49      위치지정 서식 사용시 가로 입력 추가 , 중복실행 방지 프로세스 이름에 기본이름 추가 데이터 입력 구간에 on error 추가 - 에러날때 엑셀 닫기
'2021-04-30 Ver 1.0.50      Label11 어셈블리 버전 자동 삽입, 스플레쉬 이미지 삽입, 링크라벨 이용 - 홈페이지 연결
'2021-05-06 Ver 1.0.51      설치파일 (MSI) 제작
'2021-05-07 Ver 1.1.0       유저테스트용 최종 빌드 - 유저테스트용 패스워드 지정, MID.dll파일 생성 - MID : 유저 컴퓨터 ID 
'2021-05-25 Ver 1.1.1       디자인 추가 수정.
'2021-06-03 Ver 1.1.2       위치지정 라벨 체크 안할시 데이터 끝 위치 검출 구문 셀 주소 nothing 오류 수정
'2021-06-09 Ver 1.1.3       메인화면 및 버튼 크기 조정
'2021-07-27 Ver 1.1.4       라벨 수정, 생성시 위치지정 서식 텍스트 박스 비활성화 수정,  위치지정 추가기입 프로젝트 추가(디자인)
'2021-08-24 Ver 1.1.5       위치지정 서식의 추가기입용 다이얼로그 추가 - 원하는 위치에 하나의 텍스트 문장을 측정후 성적서 생성 마다 혹은 매 측정시 자동으로 기입되는 기능 
'2021-08-25 Ver 1.2.0       대구 유저테스트를 위해 비밀번호 수정 TrialMRM0730 > TrialMRM0930 , 사용기한 9/30까지 연장
'2021-08-25 Ver 1.2.1       add_str_dialog에 Keycode 이벤트 추가 Shift+enter = 줄바꿈, Enter = 확인키 , 확인키 클릭시 dialog결과 OK(1)로 출력 하게 하기
'2021-09-06 Ver 1.2.2       병합셀 확인하여 병합셀의 크기 만큼 셀주소를 조정 해서 병합셀에 데이터가 중복 입력 안되도록 수정 >> 병합셀에서도 데이터가 하나만 들어갈수 있음
'2021-10-15 Ver 1.3.0       위치지정 서식 기초추가 >ini 파일준비 >>메인화면에서 설정정보 읽어오는 기능구현, 위치지정서식 동적컨트롤로 탭페이지로 표현 
'2021-10-21 Ver 1.3.1       동적 컨트롤 위치 지정, 및 property 지정 및 수정
'2021-10-24 Ver 1.3.2       Mknew, 수정창에 위치지정 서식 동적 컨트롤 대응 할 수 있도록 수정 > Mknew,수정창에 동적 컨드롤 생성 > 원본 성적서 불러와서 탭 갯수 만큼 tabpage생성 후 이름 지정.
'2021-10-24 Ver 1.3.3       위치지정 서식 추가기입창 임시 저장 기능 추가, 저장 후 다음 번 불러올떄는 저장된 데이터 불러오도록 하기. 동적 컨트롤 수정, 가로 기입 오류 수정
'2021-10-25 Ver 1.3.4       위치지정 서식 마무리 > 사용 성적서의 탭을 읽어와 tabpage 자동 생성 후 각 탭마다 집어넣을 값, 줄수 따로 지정하여 모든 탭에 각자 데이터 집어 넣을 수 있도로 변경, 위치지정 추가 기입창에 적용 탭 항목 추가 및 수정
'2021-11-08 Ver 1.3.5       일반, 전용 버튼 추가하여 전용프로그램 리스트 읽어와서 수정 가능하도록 함. > 전용프로그램의 수정 및 관리가 편리해짐.
'2021-11-11 Ver 1.3.6       수정창 > 원본 성적서 불러오기후 탭 재생성 기능 추가 / 새로만들기, 수정창 > 원본성적서 불러오는 다이얼로그 취소 누를 경우 에러 나는 현상 해경
'2021-11-15 Ver 1.3.7       수정창, Mknew, Form1에 위치지정_추가기입 input_type 속성 추가 > 날짜, 시간 등 선택하여 날짜, 시간 자동 입력
'2021-11-29 Ver 1.3.8       사용설명서 업데이트 Rev.004 업데이트 
'2021-12-13 Ver 1.3.9       상한,하한공차,오차 위치값 재배열
'2022-01-14 Ver 1.3.10      위치지정 서식 추가 기입창 입력 에러(누락 부분 18,19,20을 load에서 불러오지 않아서 누락 기입), 추가 기입창 안키고 수정 저장 하면 공백으로 사라지는 현상 수정
'=============================================================================================================

'=========================================
'New_Fix_Check = 0 신규 작성시
'New_Fix_Check = 1 폼 변경시 유저정보 동일 하게 불러오기 
'New_Fix_Check = 2 수정 버튼 눌러서 유저정보 수정시
'=========================================
Imports MID_CHECK.Class1
Public Class Form1

    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    '============================================================== Active Key
    Const active_key As String = "MRMV103"     '유저테스트용 빌드 mrm 활성화 비밀번호

    'Const MID As String = ""
    Const MRM_root_dir As String = "C:\MitutoyoApp"
    '============================================================== Active Key
    Dim Save_Dir As String
    Dim CSV_Dir As String
    Dim Open_Dir As String
    Dim strDate As String
    Dim strTime As String
    Dim Matching_List() As String
    Dim iniPath As String
    Dim logo_path As String
    Dim tempsave_dir As String

    Dim ini_Name As String
    Dim ini_dir As String

    Dim Cell_Address As Integer
    Dim cXL As Object
    Dim XL As Object
    Dim gXL As Object
    Dim Sheet_Count As Integer
    Dim Line_Count As Integer
    Dim Quetient As Double

    Dim Cell_Count As Integer
    Dim Cell_Count2 As Integer
    Dim TotalSheet As Integer
    Dim strData() As String
    Dim source_type As String

    Dim select_pic_name As String
    Dim select_pic_name_2 As String
    Dim select_pic_name_3 As String

    Dim select_image As Bitmap
    Dim select_image_2 As Bitmap
    Dim select_image_3 As Bitmap

    Dim error_count As Integer
    Dim auto_save_check As String

    Dim Data_error_occur As Integer              '소스파일 및 엑셀 누락 에러 검출용 변수

    Public List_check As Integer

    Public Product_Name As String
    Public Machine_Name As String
    Public Request_Dept As String
    Public Request_Date As String
    Public Drawing_Num As String
    Public Program_Name As String
    Public Player_Name As String
    Public Measure_Date As String

    Public specialized_dir As String
    Public specialized_ini As String

    Public process_name As String

    Public New_Fix_check As Integer

    Public user_info_temp() As String
    Public add_str_value(10) As String

    Dim for_trial As Date
    Dim for_trial2 As Long


    Structure address
        Dim Col As String
        Dim Row As String
    End Structure

    Structure check_value
        Dim label As String
        Dim Measure_value As String
        Dim Design_value As String
        Dim Error_check As String
        Dim UP_tol As String
        Dim Low_tol As String
        Dim judge As String
        Dim component As String
    End Structure
    Structure control_structure
        Dim Text_box() As TextBox
        Dim check_box() As CheckBox
        Dim label() As Label
    End Structure




    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load            '로드시 처음 행동

        Me.Location = New Point(50, 50)
        '==================================================== 머신 활성화 체크
        '==================================================== mysetting 초기화
        'My.Settings.등록여부 = False
        'MsgBox(My.Settings.등록_ID)
        'My.Settings.등록_ID = "2132314"
        'My.Settings.Save()
        'me.close()
        'My.Settings.Reset()
        '==================================================== mysetting 초기화
        ' My.Settings.Reload()

        '==================================================== 트라이얼 버전용

        'for_trial = "2022-11-01"
        'for_trial2 = DateDiff(DateInterval.Day, Now, for_trial) + 1
        'If for_trial2 <= 0 Then
        'MsgBox("테스트 사용기간이 종료 되었습니다." & Environment.NewLine & "정식판 사용을 위해서 한국미쓰도요 영업부에 연락 부탁드립니다." & Environment.NewLine & "대표 번호 : 031-361-4220", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
        'End
        'End If

        '==================================================== 트라이얼 버전용

        '==================================================== 활성화 체크
        Select Case Lib_Serial_Check()
            Case 1
                End
        End Select

        '  If Lib_Serial_Check() = False Then
        'Me.Close()
        '  End If
        '==================================================== 활성화 체크

        '==================================================== 중복실행 방지

        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            MsgBox("프로그램이 이미 실행중입니다!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, Me.Text)
            End
        End If

        If UBound(Diagnostics.Process.GetProcessesByName("Mitutoyo Result Matcher")) > 0 Then
            MsgBox("프로그램이 이미 실행중입니다!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, Me.Text)
            End
        End If


        '====================================================리스트 파일 로드

        Get_list()


        '==================================================== MRM 하위 폴더 유무 확인후 폴더 생성
        Dim Folder_Name() As String
        ReDim Folder_Name(15)
        Folder_Name(0) = MRM_root_dir
        Folder_Name(1) = "\MRM"
        Folder_Name(2) = "\Data"
        Folder_Name(3) = "\Result"
        Folder_Name(4) = "\Resources"
        Folder_Name(5) = "\ini"
        Folder_Name(6) = "\temp"
        Folder_Name(7) = "\전용프로그램"
        Folder_Name(10) = System.Reflection.Assembly.GetExecutingAssembly.Location
        Folder_Name(11) = "\MID"
        Dim folder_finder As New System.IO.DirectoryInfo(Folder_Name(0) & Folder_Name(1))

        If folder_finder.Exists = False Then
            MkDir(Folder_Name(0) & Folder_Name(1))
            MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2))
            MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4))
            MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(5))
            ' MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(11))

            MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(3))
            MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(3) & Folder_Name(6))

            'MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(7))

            'FileCopy(Folder_Name(10), Folder_Name(0) & Folder_Name(1) & "\Mitutoyo Result Matcher.exe")
            '  Call SetAttr(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(11), FileAttribute.Hidden)

        Else
            Dim folder_finder2 As New System.IO.DirectoryInfo(Folder_Name(0) & Folder_Name(1) & Folder_Name(2))
            If folder_finder2.Exists = False Then
                ' MkDir(Folder_Name(0) & Folder_Name(1))
                MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2))
                MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4))
                MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(5))
                ' MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(11))

                MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(3))
                MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(3) & Folder_Name(6))

                'MkDir(Folder_Name(0) & Folder_Name(1) & Folder_Name(7))

                '  FileCopy(Folder_Name(10), Folder_Name(0) & Folder_Name(1) & "\Mitutoyo Result Matcher.exe")
                '  Call SetAttr(Folder_Name(0) & Folder_Name(1) & Folder_Name(2) & Folder_Name(4) & Folder_Name(11), FileAttribute.Hidden)
            End If
        End If



        '====================================================
        '개요 select case 문으로 레지스트리 값 읽어 오는 판단에서 분기점
        '읽어 올때 레지스트리가 존재하면 정상 구동, 없으면 레지스트리 등록 절차로 이동
        '레지스트리 읽어 올 때 장비 id 랑 일치 여부 확인 복사 방지구문 구성
        '레지스트리 등록시 장비 id 등록 
        ' 결과예상 - 첫플레이시에만 장비아이디 등록 
        ' 이후 플레이시에는 레지스트리가 존재하기때문에 등록 절차없이 플레이
        '예상 문제점 컴퓨터를 옮겼는데 예전 컴퓨터에서 여전히 플레이 가능한 문제.
        'MsgBox(Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber))

        '================================전용화 판단용 파일 존재유무 판단

        process_name = System.Diagnostics.Process.GetCurrentProcess().ProcessName
        Dim specialized_string As String = "\MRM\Data\Resources\ini\" & process_name
        specialized_dir = MRM_root_dir & specialized_string
        specialized_ini = specialized_dir & "\" & process_name & ".ini"

        Dim specialized_folder As New System.IO.DirectoryInfo(specialized_dir)

        If specialized_folder.Exists = True Then
            If Specialized_play(specialized_dir, specialized_ini) = 0 Then
                Me.Close()
                End
            Else
                프로그레스바.Close()
                MsgBox("전용 프로그램 실행 중 오류가 발생하였습니다." & Environment.NewLine & " 전용 프로글램용 폴더 경로의 확인이나 전용 프로그램을 다시 생성해 주세요.",, "전용프로그램 실행 오류")
                End
            End If

        End If

        '================================전용화 판단용 파일 존재유무 판단

        CheckBox1.Checked = True

        ReDim user_info_temp(15)           '기본정보(머릿말) 내용 저장용 변수

        Label11.Text = System.String.Format(Label11.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build)

        ' MsgBox(My.Application.Info.Version.ToString)

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           '성적서 생성 버튼


        Data_error_occur = 0
        If ListBox1.SelectedItem IsNot Nothing Then
        Else
            MsgBox("성적서 매칭 리스트를 선택해주세요",, "매칭 리스트 선택 오류")
            Exit Sub
        End If

        ini_Name = ListBox1.SelectedItem.ToString()

        ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & ini_Name & ".ini"

        '======================================================

        ReDim strData(10)

        Open_Dir = MRM_root_dir & "\MRM\Data\Resources\Result_Source.xlsx"   '원본 엑셀
        Dim fFindFile As New System.IO.FileInfo(Open_Dir)
        If fFindFile.Exists = False Then

            Call Origin_form()
        End If

        XL = CreateObject("Excel.application")

        CSV_Dir = Label16.Text
        '=========================================================================
        '프로그래스바
        프로그레스바.Location = New Point(500, 500)
        프로그레스바.Show()
        '=========================================================================
        Dim type_value As String
        type_value = UCase(Strings.Right(CSV_Dir, 3))
        Select Case type_value
            Case "CSV"
                source_type = 1
            Case "ASC"
                source_type = 2

            Case Else
                MsgBox("데이터 파일의 형식이 잘못 되었습니다. 확인후 다시 시도해주세요", 48, "Error Occurred")
                프로그레스바.Close()
                Exit Sub
        End Select

        Select Case source_type

            Case 1
                Call Extension_type_1()     'csv
            Case 2
                Call Extension_type_2()     'asc   
        End Select

        If Data_error_occur = 1 Then
            MsgBox("측정 데이터 누락 혹은 성적서 서식 엑셀 누락 에러 발생" & Environment.NewLine & "측정 데이터 혹은 성적서 서식 엑셀의 이름, 존재 여부 확인 후 재시도 부탁드립니다. ", 48, "Error Occurred")
            프로그레스바.Close()
            Exit Sub
        End If

        Select Case strDate
            Case "True"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & DateString & "_" & Format(TimeOfDay, "HH-mm-ss") & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & DateString & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                End Select
            Case "False"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & Format(TimeOfDay, "HH-mm-ss''") & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))
                End Select

        End Select


        WPPS("Matching_Info", "Last_Paly_date", Now(), Restore_str(ini_dir))         '마지막 측정 시간 저장
        프로그레스바.Close()

        tempsave_dir = MRM_root_dir & "\MRM\Result\temp"
        Dim fFindFolder As New System.IO.DirectoryInfo(tempsave_dir)   '  temp폴더 존재 유무 확인용 선언

        Select Case Label22.Text
            Case "예"

                Select Case Label19.Text
                    Case ".xlsx"
                        With XL
                            .DisplayAlerts = False
                            '.Sheets(1).select
                            .workbooks(1).SaveAS(filename:=Restore_str(Save_Dir))          '다른이름으로 저장 위치 지정
                            .Workbooks(1).close
                            .quit
                        End With

                    Case ".PDF"
                        With XL
                            .DisplayAlerts = False
                            .workbooks(1).ExportAsFixedFormat(Type:=0, Filename:=Restore_str(Save_Dir), Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)
                            If fFindFolder.Exists = False Then
                                MkDir(MRM_root_dir & "\MRM\Result\temp")
                            End If
                            tempsave_dir = MRM_root_dir & "\MRM\Result\temp"
                            .workbooks(1).SaveAS(tempsave_dir & "\temp-auto.xlsx")
                            .Workbooks(1).close
                            .quit
                        End With
                End Select

                XL = Nothing
                '=============================================================================== 에러 구문 출력 
                If error_count <> 0 Then
                    Select Case error_count

                        Case 1          '링크 사진 경로 삭제 및 이동.
                            MsgBox("첨부한 그림의 경로가 변경되었거나 삭제되었습니다." & Environment.NewLine & "첨부 그림 경로나 파일 존재 유무를 다시 한번 확인해 주세요.",, "Error Occurred")

                    End Select
                End If
                '=============================================================================== 에러 구문 출력 
                If CheckBox1.Checked = True Then
                    Me.Close()
                End If

            Case "아니오"
                MsgBox("자동저장을 선택하지 않으셨습니다." & Environment.NewLine & "성적서가 화면에 켜집니다." & Environment.NewLine & "성적서를 따로 저장해주세요.",, "성적서 자동 저장 취소")
                XL.DisplayAlerts = False
                If fFindFolder.Exists = False Then
                    MkDir(MRM_root_dir & "\MRM\Result\temp")
                End If
                XL.workbooks(1).SaveAS(tempsave_dir & "\temp.xlsx")
                XL.visible = True
                '=============================================================================== 에러 구문 출력 
                If error_count <> 0 Then
                    Select Case error_count

                        Case 1          '링크 사진 경로 삭제 및 이동.
                            MsgBox("첨부한 그림의 경로가 변경되었거나 삭제되었습니다." & Environment.NewLine & "첨부 그림 경로나 파일 존재 유무를 다시 한번 확인해 주세요.",, "Error Occurred")

                    End Select
                End If
                '=============================================================================== 에러 구문 출력 
                If CheckBox1.Checked = True Then
                    Me.Close()
                End If
        End Select

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim MkNew As New MkNew()
        MkNew.ShowDialog()

    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
        MsgBox("성적서 생성을 취소 하였습니다.",, "성적서 생성 취소")
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        If ListBox1.SelectedItem IsNot Nothing Then
            ini_Name = ListBox1.SelectedItem.ToString()
        End If

        Select Case List_check
            Case 1

                ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & ini_Name & ".ini"

            Case 2
                ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & ini_Name & "\" & ini_Name & ".ini"
        End Select

        If ListBox1.SelectedItem = Nothing Then

            '   strDate = ""
            '   strTime = ""
            '   Label16.Text = ""
            '   Label17.Text = ""
            '   Label22.Text = ""
            '   Label19.Text = ""
            '   Label20.Text = ""
            '   iniPath = ""
            '   logo_path = ""
            '   Label18.Text = ""

        Else
            Call Input_property(ini_dir)

            'Call Input_property_t(ini_dir)            '테스트용


        End If
    End Sub
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
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim temp_kill_Dir As String
        Dim temp_kill_Name As String
        Dim temp_kill_exe As String
        Dim temp_kill_folder As String

        If ListBox1.SelectedItem IsNot Nothing Then
            If MsgBox("선택 리스트를 삭제하시겠습니까?", 4, "매칭 리스트 삭제") = 6 Then           'YES : 6, NO: 7

                On Error Resume Next
                Select Case List_check
                    Case 1          '일반
                        temp_kill_Name = ListBox1.SelectedItem.ToString
                        temp_kill_Dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & temp_kill_Name & ".ini"
                        ListBox1.Items.Remove(ListBox1.SelectedItem)
                        Kill(temp_kill_Dir)

                    Case 2          '전용


                        temp_kill_Name = ListBox1.SelectedItem.ToString
                        temp_kill_Dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & temp_kill_Name & "\" & temp_kill_Name & ".ini"
                        temp_kill_folder = MRM_root_dir & "\MRM\Data\Resources\ini\" & temp_kill_Name
                        temp_kill_exe = MRM_root_dir & "\MRM\전용프로그램\" & temp_kill_Name & ".exe"
                        ListBox1.Items.Remove(ListBox1.SelectedItem)
                        Kill(temp_kill_Dir)         '전용 ini삭제
                        RmDir(temp_kill_folder)     '전용 folder 삭제
                        Kill(temp_kill_exe)         '전용 exe 삭제

                End Select

                On Error GoTo 0

            End If
        Else
            MsgBox("삭제할 매칭 리스트를 선택해주세요",, "삭제 리스트 선택")
            Exit Sub

        End If

        Exit Sub

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedItem IsNot Nothing Then
            수정창.ShowDialog()
            Dim ReSelect As Integer
            ReSelect = ListBox1.SelectedIndex

            ListBox1.ClearSelected()
            ListBox1.SelectedIndex = ReSelect
        Else
            MsgBox("수정할 매칭 설정을 리스트에서 선택해주세요",, "수정 리스트 선택")
            Exit Sub
        End If

    End Sub
    Function Ad_NUM(ByVal Inputstr As String)
        Dim i As Integer
        Dim Nostr As String

        For i = 1 To Len(Inputstr)
            If IsNumeric(Mid(Inputstr, i, 1)) Then
                Nostr = Nostr & Mid(Inputstr, i, 1)
            End If
        Next i

        Ad_NUM = Nostr

    End Function
    Function Ad_Str(ByVal inputstr As String)
        Dim i As Integer
        Dim onlystr As String
        Dim tempstr As String

        For i = 1 To Len(inputstr)

            tempstr = Mid(inputstr, i, 1)

            Select Case tempstr
                Case "A" To "Z"
                    onlystr = onlystr & tempstr
                Case "a" To "z"
                    onlystr = onlystr & tempstr
            End Select

        Next i

        Ad_Str = onlystr
    End Function

    Function Lib_Serial_Check() As Long
        Dim Serial_NUM As String
        Dim reg_path As String
        Dim registry_value As Long
        Dim registry_time As String

        Dim software_chk_QV As String
        Dim QV_install_chk As String
        Dim QV_install_value As String
        Dim QV_version As String

        Dim software_chk_CMM As String
        'Dim CMM_install_chk As String
        Dim CMM_install_value As String
        'Dim CMM_version As String
        Dim active_ans As Integer



        Dim install_check As String

        'Dim UID_check As String
        reg_path = "HKEY_CURRENT_USER\Software\Mitutoyo\MRM"
        registry_value = My.Computer.Registry.GetValue(reg_path, "Active_Machine_ID", Nothing)
        Serial_NUM = Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)
        registry_time = Now()

        software_chk_QV = "HKEY_LOCAL_MACHINE\SOFTWARE\MEI\QVPak"               '폴더 존재 유무 확인용 
        QV_version = My.Computer.Registry.GetValue(software_chk_QV, "Current Version", Nothing)     'QV 버전 확인
        QV_install_chk = "HKEY_LOCAL_MACHINE\SOFTWARE\MEI\QVPak\" & QV_version                      '각 버전 폴더 진입
        QV_install_value = My.Computer.Registry.GetValue(QV_install_chk, "", Nothing)         '설치 유무 기본값  "" =기본값

        software_chk_CMM = "HKEY_CURRENT_USER\SOFTWARE\Mitutoyo\GEOPAK"               '폴더 존재 유무 확인용 
        CMM_install_value = My.Computer.Registry.GetValue(software_chk_CMM, "", "MCOSMOS")         '설치 유무 기본값
        '===================CMM 확인 루틴 설명
        '레지스트리에 GEOPAK이 있으면 GEOPAK의 기본값을 읽어 오지만 아무것도 없기 때문에 기본값 출력  -> 기본값 출력하면 소프트웨어 존재
        '레지스트리에 GEOPAK이 없으면 Nothing값 출력 -> Nothing값 출력하면 소프트웨어 존재 안함



        install_check = My.Computer.Registry.GetValue(reg_path, "install_check", "false")

        On Error GoTo MID_DLL_ERROR

        If install_check = "false" Then
            MsgBox("MRM이 설치 되어있지 않습니다." & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다.", 48, "MRM Activation")  '6 : YES  7 : NO
            Me.Close()
            Lib_Serial_Check = 1
            GoTo mid_check_skip

        ElseIf install_check = "true" Then

        Else
            MsgBox("MRM이 설치 되어있지 않습니다." & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다.", 48, "MRM Activation")  '6 : YES  7 : NO
            Me.Close()
            Lib_Serial_Check = 1
            GoTo mid_check_skip

        End If


        ' If registry_value <> Serial_NUM Then    '      
        If registry_value <> MID_CHECK_EH(Serial_NUM) Then    'dll파일에서 MID 읽어와서 일치하는지 확인
            '
            'active_ans = MsgBox("MRM 사용이 활성 되어있지 않습니다." & Environment.NewLine & Environment.NewLine & "MRM을 활성 하시겠습니까?", 4, "MRM Activation")
            'If active_ans = 6 Then  '6 : YES  7 : NO
            활성화키.ShowDialog()
            활성화키.Focus()

            If UCase(활성화키.TextBox1.Text) = UCase(active_key) Then

                If QV_install_value = "Install completed" Then 'qv 설치되어있음

                    'My.Settings.등록_ID = Serial_NUM
                    'My.Settings.등록여부 = True
                    'My.Settings.Save()
                    My.Computer.Registry.SetValue(reg_path, "Active_Machine_ID", Serial_NUM)
                    My.Computer.Registry.SetValue(reg_path, "Last_Active_Time", registry_time)
                    'My.Computer.Registry.SetValue(reg_path, "Active_UID", active_key)
                    My.Computer.Registry.SetValue(reg_path, "Active_Software", "QVPAK")

                    My.Computer.Registry.SetValue(QV_install_chk & "\QVClientMenu Config", "MenuName12", "MRM")
                    My.Computer.Registry.SetValue(QV_install_chk & "\QVClientMenu Config", "CommandLine12", MRM_root_dir & "\MRM\Mitutoyo Result Matcher.exe")
                    MsgBox("MRM을 성공적으로 활성 하였습니다.", 0, "MRM Activation")

                ElseIf CMM_install_value = "MCOSMOS" Then     'cmm 설치되어있음

                    'My.Settings.등록_ID = Serial_NUM
                    'My.Settings.등록여부 = True
                    'My.Settings.Save()
                    My.Computer.Registry.SetValue(reg_path, "Active_Machine_ID", Serial_NUM)
                    My.Computer.Registry.SetValue(reg_path, "Last_Active_Time", registry_time)
                    'My.Computer.Registry.SetValue(reg_path, "Active_UID", active_key)
                    My.Computer.Registry.SetValue(reg_path, "Active_Software", "MCOSMOS")
                    MsgBox("MRM을 성공적으로 활성 하였습니다.", 0, "MRM Activation")


                Else        'cmm 설치 안되어있음

                    MsgBox("MRM을 성공적으로 활성 하지 못하였습니다." & Environment.NewLine & "Mitutoyo 측정 소프트웨어가 설치되어있는 컴퓨터에서 활성화 시켜 주세요" & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다.", 48, "MRM Activation")  '6 : YES  7 : NO
                    Me.Close()
                    Lib_Serial_Check = 1


                End If

            Else
                MsgBox("MRM을 성공적으로 활성 하지 못하였습니다. 입력한 활성화 키를 다시한번 확인 부탁드립니다." & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다.", 48, "MRM Activation")  '6 : YES  7 : NO
                Me.Close()
                Lib_Serial_Check = 1
            End If
            ' Else

            'MsgBox("MRM 활성화를 취소하셨습니다.", 64, "MRM Activation")
            'Me.Close()
            'Lib_Serial_Check = 1

            'End If

        Else '
            If QV_install_value = "Install completed" Then 'qv 설치되어있음


            ElseIf CMM_install_value = "MCOSMOS" Then     'cmm 설치되어있음

            Else        'QV, cmm 설치 안되어있음

                MsgBox("MRM을 성공적으로 활성 하지 못하였습니다." & Environment.NewLine & "Mitutoyo 측정 소프트웨어가 설치되어있는 컴퓨터에서 활성화 시켜 주세요" & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다.", 48, "MRM Activation")  '6 : YES  7 : NO
                Me.Close()
                Lib_Serial_Check = 1

            End If

        End If





        '============================레지스트리 등록 확인법 
        'Lib_Serial_Check = True
        'Dim reg_path As String
        'Dim registry_value As Long
        'reg_path = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\MRM"
        'registry_value = My.Computer.Registry.GetValue(reg_path, "Active_Machine_ID", Nothing)
        'Select Case Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)
        ' Case registry_value  '레지스트리에 등록 되어있는 사용 장비 C드라이브 시리얼 번호 가져오기 

        'Case Else
        'Lib_Serial_Check = False
        'MsgBox("프로그램 활성화를 위해서 아래 하드웨어 ID 를 기록하여" & Chr(13) & "Mitutoyo Korea 로 문의 바랍니다." &
        'Chr(13) & Chr(13) & "HardWare ID : " & Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber), 16, "복사 방지 오류")
        ''MkDir("C:\")
        'Lib_Serial_Check = 1

        'End Select

mid_check_skip:

        Exit Function
MID_DLL_ERROR:
        'SplashScreen1.Close()
        MsgBox("        >>>>>   파일 복사 감지   <<<<<    " & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다." & Environment.NewLine & "문의전화(영업부) : 각 담당 영업사원" & Environment.NewLine & "문의전화 (영업기술부) : 031-361-4274" & Environment.NewLine & "ERROR CDOE : M-001", 48, "MRM Activation")          'M-001 시리얼 체크 에러 발생

        Me.Close()
        End
    End Function
    Sub Origin_form()
        '======================정렬 value
        'top: -4160
        'Bottom: -4107
        'Left: -4131
        'Center: -4108
        'Right: -4152

        '========================================외각선 value
        'xlDiagonalDown 5  Border running from the upper-left corner to the lower-right of each cell in the range.
        'xlDiagonalUp   6  Border running from the lower-left corner to the upper-right of each cell in the range.

        'xlEdgeLeft  7  Border at the left edge of the range.
        'xlEdgeTop   8  Border at the top of the range.
        'xlEdgeBottom   9  Border at the bottom of the range.
        'xlEdgeRight 10 Border at the right edge of the range.

        'xlInsideVertical  11 Vertical borders for all the cells in the range except borders on the outside of the range.
        'xlInsideHorizontal   12 Horizontal borders for all cells in the range except borders on the outside of the range.

        '======================외각선 스타일 value
        'xlContinuous   1  Continuous line.
        'xlDash   -4115 Dashed line.
        'xlDashDot   4  Alternating dashes and dots.
        'xlDashDotDot   5  Dash followed by two dots.
        'xlDot -4118 Dotted line.
        'xlDouble -4119 Double line.
        'xlLineStyleNone   -4142 No line.
        'xlSlantDashDot 13 Slanted dashes.
        '=========================================================================

        '=========================================================================
        cXL = CreateObject("Excel.application")
        'cXL.Visible = True
        cXL.Workbooks.add
        cXL.sheets(1).name = "기본폼1"
        Cell_Address = 8

        With cXL.Range("A1:I8")
            .HorizontalAlignment = -4108
            .Borders(7).LineStyle = 1
            .Borders(8).LineStyle = 1
            .Borders(9).LineStyle = -4119        '2중 라인
            .Borders(10).LineStyle = 1
            .Borders(11).LineStyle = 1
            .Borders(12).LineStyle = 1

        End With

        cXL.Range("A8:I8").Borders(8).LineStyle = -4119

        With cXL.Range("A9:I36")
            .Borders(7).LineStyle = 1
            .Borders(9).LineStyle = -4119        '2중 라인
            .Borders(10).LineStyle = 1
            .Borders(11).LineStyle = 1
            .Borders(12).LineStyle = -4118
            .shrinktofit = True
        End With

        With cXL.Range("A35:I41")
            .Borders(7).LineStyle = 1
            .Borders(9).LineStyle = 1
            .Borders(10).LineStyle = 1
        End With

        cXL.Range("A37").value = "<코멘트>"

        cXL.Columns("A:I").ColumnWidth = 8.5
        cXL.Columns("B:B").ColumnWidth = 5.5
        cXL.Rows("2:2").RowHeight = 50

        With cXL.Range("A1:F3")
            .merge
            .HorizontalAlignment = -4152
            .VerticalAlignment = -4107
            .Font.Name = "맑은 고딕"
            .Font.Size = 28
            .FormulaR1C1 = "검사성적서  "
        End With
        With cXL.Range("G1")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "승인"
        End With
        With cXL.Range("H1:I2")
            .merge
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
            .ShrinkToFit = True
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            ' .Value = "Mitutoyo Result Matcher"
        End With
        With cXL.Range("H3")
            .merge
            .HorizontalAlignment = -4108
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "페이지"
        End With

        With cXL.Range("A4:C4")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "품명"
        End With
        With cXL.Range("A5:C5")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "장비 이름"
        End With
        With cXL.Range("A6:C6")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "의뢰 부서"
        End With
        With cXL.Range("A7:C7")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "의뢰 날짜"
        End With

        With cXL.Range("F4:G4")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "도면 번호"
        End With
        With cXL.Range("F5:G5")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "프로그램 이름"
        End With
        With cXL.Range("F6:G6")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "측정자"
        End With
        With cXL.Range("F7:G7")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "측정 날짜"
        End With

        With cXL.Range("A8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "라벨명(ID)"
        End With
        With cXL.Range("B8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "요소"
        End With
        With cXL.Range("C8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "측정값"
        End With
        With cXL.Range("D8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "기준값"
        End With
        With cXL.Range("E8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "오차"
        End With
        With cXL.Range("F8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "상한공차"
        End With
        With cXL.Range("G8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "하한공차"
        End With
        With cXL.Range("H8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "판정"
        End With
        With cXL.Range("I8")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = "비고"
        End With


        With cXL.Range("D4:E4")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("D5:E5")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("D6:E6")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("D7:E7")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With


        With cXL.Range("H4:I4")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("H5:I5")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("H6:I6")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With
        With cXL.Range("H7:I7")
            .merge
            .Font.Name = "맑은 고딕"
            .Font.Size = 11
            .Font.Bold = True
            .Value = ""
        End With

        With cXL.Range("H9:H36")
            .formatconditions.add(Type:=9, String:="통과", TextOperator:=0)
            .FormatConditions(1).SetFirstPriority
            .FormatConditions(1).Interior.PatternColorIndex = -4105
            .FormatConditions(1).Interior.Color = 5287936
            .FormatConditions(1).Interior.TintAndShade = 0
            .FormatConditions(1).StopIfTrue = False
        End With

        With cXL.Range("H9:H36")
            .formatconditions.add(Type:=9, String:="실패", TextOperator:=0)
            .FormatConditions(2).Interior.PatternColorIndex = -4105
            .FormatConditions(2).Interior.Color = 255
            .FormatConditions(2).Interior.TintAndShade = 0
            .FormatConditions(2).StopIfTrue = False
        End With

        With cXL.Range("H9:H36")
            .formatconditions.add(Type:=9, String:="=""" & "-----+--->" & """", TextOperator:=0)
            .FormatConditions(3).Interior.PatternColorIndex = -4105
            .FormatConditions(3).Interior.Color = 255
            .FormatConditions(3).Interior.TintAndShade = 0
            .FormatConditions(3).StopIfTrue = False
        End With


        With cXL.Range("H9:H36")
            .formatconditions.add(Type:=9, String:="<<---+-----", TextOperator:=0)
            .FormatConditions(4).Interior.PatternColorIndex = -4105
            .FormatConditions(4).Interior.Color = 255
            .FormatConditions(4).Interior.TintAndShade = 0
            .FormatConditions(4).StopIfTrue = False
        End With

        cXL.range("I3").NumberFormatLocal = "@"

        'Dim picture_dir() As Object
        'ReDim picture_dir(2)
        'Dim Pic_L As Single
        'Dim Pic_T As Single 'top
        'Dim Pic_W As Single
        'Dim Pic_H As Single
        'Pic_L = cXL.sheets("기본폼1").Range("A1").Left
        'Pic_T = cXL.sheets("기본폼1").Range("A1").Top
        'Pic_W = 180
        'Pic_H = 45
        'My.Resources._1920px_Mitutoyo_company_logo_sample.Save(MRM_root_dir & "\MRM\Data\Resources\default_logo.png")
        'picture_dir(0) = MRM_root_dir & "\MRM\data\Resources\default_logo.png"
        'cXL.sheets("기본폼1").shapes.addpicture(fileName:=picture_dir(0), Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L, Top:=Pic_T, Width:=Pic_W, Height:=Pic_H).select

        ' With cXL.Selection.ShapeRange
        '.Width = 180.0
        '.Height = 45.0
        ' .IncrementTop(2)
        '.IncrementLeft(2)
        'End With

        'Pic_L = cXL.sheets("기본폼1").Range("H1").Left
        'Pic_T = cXL.sheets("기본폼1").Range("H1").Top
        'Pic_W = 105
        'Pic_H = 29
        'My.Resources._1920px_Mitutoyo_company_logo.Save(MRM_root_dir & "\MRM\Data\Resources\default_logo (2).png")
        'picture_dir(1) = MRM_root_dir & "\MRM\data\Resources\default_logo (2).png"
        'cXL.Range("H1").select
        'cXL.sheets("기본폼1").shapes.addpicture(fileName:=picture_dir(1), Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L, Top:=Pic_T, Width:=Pic_W, Height:=Pic_H).select
        ' With cXL.Selection.ShapeRange
        '.Width = 117.0
        '.Height = 29.0
        '.IncrementTop(17.75)
        '.IncrementLeft(2)

        'End With
        '==========================================================================         기본폼 2
        cXL.Sheets("기본폼1").Copy(After:=cXL.Sheets("기본폼1"))
        cXL.Sheets(2).name = "기본폼2"

        With cXL.sheets("기본폼2").rows("5:6")
            .select
            .delete(shift:=-4162)
        End With
        With cXL.sheets("기본폼2").range("A1:F3")
            .select
            .unmerge
        End With
        With cXL.Sheets("기본폼2").rows("1:2")
            .select
            .delete(shift:=-4162)
        End With

        With cXL.sheets("기본폼2").Range("A1:G1")
            .merge
            .Borders(8).LineStyle = 1
            .HorizontalAlignment = -4108        'center
            .VerticalAlignment = -4108          'center
            .Font.Name = "맑은 고딕"
            .Font.Size = 28
            .FormulaR1C1 = "검사성적서"
        End With

        ' With cXL.Sheets("기본폼2").Shapes.Range(1)
        '   .Select
        ' .Delete
        ' End With
        '  With cXL.Sheets("기본폼2").Shapes.Range(1)
        ' .Select
        '  .Delete
        '  End With

        cXL.sheets("기본폼2").rows("1:1").rowheight = 37.5

        With cXL.sheets("기본폼2").Rows("32:32")
            .Select

            .insert(Shift:=-4121, CopyOrigin:=0)
            .insert(Shift:=-4121, CopyOrigin:=0)
            .insert(Shift:=-4121, CopyOrigin:=0)
            .insert(Shift:=-4121, CopyOrigin:=0)
        End With

        '==========================================================================         기본폼 3
        cXL.Sheets("기본폼1").Copy(After:=cXL.Sheets("기본폼2"))
        cXL.Sheets(3).name = "기본폼3"

        With cXL.Sheets("기본폼3").rows("8:8")
            .select
            .copy
        End With
        With cXL.Sheets("기본폼3").rows("21:21")
            .select
        End With
        cXL.Sheets("기본폼3").paste
        With cXL.sheets("기본폼3").Range("A8:I20")
            .select
            .clearcontents
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4108
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ReadingOrder = -5002  'xlContext
            .MergeCells = False
            .merge
        End With

        '=========================================================================
        cXL.workbooks(1).SaveAS(MRM_root_dir & "\MRM\data\Resources\Result_Source.xlsx")
        cXL.DisplayAlerts = False
        cXL.workbooks(1).close
        cXL.quit
        cXL = Nothing
        'Kill(picture_dir(0))
        'Kill(picture_dir(1))

        '=========================================================================

    End Sub


    Sub Extension_type_1()          'CSV 파일 읽어오기 위치지정 다중 탭 실험용
        '  On Error GoTo XLC
        Dim input_string As String
        Dim merge_count As Long
        Dim cRow As Integer
        Dim i As Integer
        Dim dump_num As Integer
        Dim blank_Stack As Integer

        '기본 혹은 위치지정 선택 
        Select Case Label20.Text
            Case "기본폼1", "기본폼2", "기본폼3", "기본폼4", "기본폼5"

                Measure_data_basic_form(Label20.Text)

             '=========================================================================
            Case "위치 지정"
                '=========================================================================




                Dim Label_ad() As address
                Dim measure_ad() As address
                Dim component_ad() As address
                Dim design_ad() As address
                Dim error_ad() As address
                Dim UP_Tol_ad() As address
                Dim Low_Tol_ad() As address
                Dim judge_ad() As address
                Dim Line_count_ad() As Integer
                Dim column_count As Integer
                Dim Result_Form_dir As String
                Dim select_check_value() As check_value
                Dim ini_tab_index As String                'ini 섹션 값 구분용
                Dim tab_index As Integer                    ' 원본 엑셀 탭 선택용

                Dim origin_sheet_name() As String
                Dim Data_sheet As String
                Dim input_direction() As String
                Dim error_arry() As String
                Dim error_txt As String
                Dim Tab_count As Integer        ' 탭 개수 인식
                Dim tab_name() As String
                Dim sum_Line_count As Integer
                Dim sum_temp As Integer


                Result_Form_dir = GetINIValue("custom_match_info", "Result_Form_dir", ini_dir)

                Tab_count = GetINIValue("custom_match_info", "tab_count", ini_dir)

                ReDim Label_ad(Tab_count)
                ReDim measure_ad(Tab_count)
                ReDim component_ad(Tab_count)
                ReDim design_ad(Tab_count)
                ReDim error_ad(Tab_count)
                ReDim UP_Tol_ad(Tab_count)
                ReDim Low_Tol_ad(Tab_count)
                ReDim judge_ad(Tab_count)
                ReDim Line_count_ad(Tab_count)
                ReDim select_check_value(Tab_count)
                ReDim tab_name(Tab_count)
                ReDim origin_sheet_name(Tab_count)
                ReDim input_direction(Tab_count)



                For tab_index = 1 To Tab_count
                    ini_tab_index = "tab_" & tab_index
                    tab_name(tab_index) = GetINIValue(ini_tab_index, "tab_name", ini_dir)
                    Label_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Label", ini_dir))
                    component_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "component", ini_dir))
                    measure_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "measure_value", ini_dir))
                    design_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Design_value", ini_dir))
                    error_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "error", ini_dir))
                    UP_Tol_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "UP_Tol", ini_dir))
                    Low_Tol_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Low_Tol", ini_dir))
                    judge_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "judge", ini_dir))

                    Label_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Label", ini_dir))
                    component_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "component", ini_dir))
                    measure_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "measure_value", ini_dir))
                    design_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Design_value", ini_dir))
                    error_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "error", ini_dir))
                    UP_Tol_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "UP_Tol", ini_dir))
                    Low_Tol_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Low_Tol", ini_dir))
                    judge_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "judge", ini_dir))

                    Line_count_ad(tab_index) = GetINIValue(ini_tab_index, "Line_count", ini_dir)

                    input_direction(tab_index) = GetINIValue(ini_tab_index, "input_direction", ini_dir)


                    select_check_value(tab_index).label = GetINIValue(ini_tab_index, "label_check", ini_dir)                                                          ' 탭마다 ini 별도로 정보 저장후 읽어오기, 키 이름 _check 추가해서 다른것과 구별 
                    select_check_value(tab_index).Measure_value = GetINIValue(ini_tab_index, "Measure_value_check", ini_dir)
                    select_check_value(tab_index).Design_value = GetINIValue(ini_tab_index, "Design_value_check", ini_dir)
                    select_check_value(tab_index).Error_check = GetINIValue(ini_tab_index, "error_check", ini_dir)
                    select_check_value(tab_index).UP_tol = GetINIValue(ini_tab_index, "UP_tol_check", ini_dir)
                    select_check_value(tab_index).Low_tol = GetINIValue(ini_tab_index, "Low_tol_check", ini_dir)
                    select_check_value(tab_index).judge = GetINIValue(ini_tab_index, "judge_check", ini_dir)
                    select_check_value(tab_index).component = GetINIValue(ini_tab_index, "component_check", ini_dir)

                    origin_sheet_name(tab_index) = GetINIValue(ini_tab_index, "tab_name", ini_dir)

                Next tab_index

                For Each sum_temp In Line_count_ad
                    sum_Line_count = sum_Line_count + sum_temp
                Next


                error_arry = Split("TP (3D),원형,동심도,진직도,PA,VT,VG,런아웃,대칭,평면도,TP (2D)", ",")

                XL.Workbooks.open(Result_Form_dir)       '성적서 오픈        원본성적서 지정
                XL.DisplayAlerts = False

                'origin_sheet_name = XL.activesheet.name
                Data_sheet = "Data_sheet"
                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
                XL.activesheet.name = Data_sheet

                FileOpen(3, CSV_Dir, OpenMode.Input)

                Do Until EOF(3)
                    cRow = cRow + 1
                    input_string = LineInput(3)
                    Dim input_arry() As String = Split(input_string, ",", -1)

                    If input_arry(2).IndexOf("]") <> -1 Then
                        Dim input_arry2() As String = Split(input_string, ",", -1)
                        On Error Resume Next
                        input_arry(0) = input_arry2(0)
                        input_arry(1) = Mid(input_arry2(1), 2) & "," & Replace(input_arry2(2), """", " ")
                        input_arry(2) = input_arry2(3)
                        input_arry(3) = input_arry2(4)
                        input_arry(4) = input_arry2(5)
                        input_arry(5) = input_arry2(6)
                        input_arry(6) = input_arry2(7)
                        input_arry(7) = input_arry2(8)
                        input_arry(8) = input_arry2(9)
                        input_arry(9) = input_arry2(10)

                    End If

                    Select Case dump_num
                        Case 0      'dump
                            For i = 0 To UBound(input_arry)
                                XL.sheets(Data_sheet).cells(cRow, i + 1).value = input_arry(i)
                                XL.sheets(Data_sheet).cells(cRow, i + 1).value = ""
                            Next

                            dump_num = 1
                            cRow = 0
                        Case 1
                            For i = 0 To UBound(input_arry)
                                XL.sheets(Data_sheet).cells(cRow, i + 1).value = input_arry(i)
                            Next
                    End Select

                Loop
                FileClose(3)
                XL.visible = True

                tab_index = 1

                '===========================================위치지정 추가기입 삽입 위치.

                Call add_str()

                '===========================================위치지정 추가기입 삽입 위치.

                '   XL.Sheets(origin_sheet_name).Copy(After:=XL.Sheets("Data_sheet"))        '데이터 입력용 워크시트 복사
                'XL.activesheet.name = tab_name(1)                         '데이터 입력용 워크시트 선택

                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 1                     '데이터 입력 워크시트 카운트

                Cell_Count2 = XL.Sheets(Data_sheet).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(Data_sheet).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                'Quetient = Cell_Count / Line_count_ad              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                프로그레스바.ProgressBar1.Maximum = Cell_Count

                '  If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 1
                    column_count = 1


                    Do Until Line_Count > Line_count_ad(tab_index)

                        If select_check_value(tab_index).label = True Then
                            XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("B" & Cell_Address).value2 & "(" & XL.Sheets(Data_sheet).Range("C" & Cell_Address).value2 & ")"   '라벨명
                            If XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        Label_ad(tab_index).Row = Label_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).component = True Then
                            XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("D" & Cell_Address).value2     '구성요소
                            If XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        component_ad(tab_index).Row = component_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If

                        End If

                        error_txt = XL.Sheets(Data_sheet).Range("D" & Cell_Address).value

                        If select_check_value(tab_index).Measure_value = True Then
                            If error_arry.Contains(error_txt) Then
                                XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).value = XL.Sheets(Data_sheet).Range("G" & Cell_Address).value      '오차 > 측정값
                            Else
                                XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).value = XL.Sheets(Data_sheet).Range("E" & Cell_Address).value     '측정값 > 측정값
                            End If

                            If XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        measure_ad(tab_index).Row = measure_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If

                        End If

                        If select_check_value(tab_index).Design_value = True Then
                            XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("F" & Cell_Address).value2     '설계치
                            If XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        design_ad(tab_index).Row = design_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If


                        If select_check_value(tab_index).Error_check = True Then
                            XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("G" & Cell_Address).value2     '오차
                            If XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        error_ad(tab_index).Row = error_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).UP_tol = True Then
                            XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("H" & Cell_Address).value2     '상한
                            If XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        UP_Tol_ad(tab_index).Row = UP_Tol_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).Low_tol = True Then
                            XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("I" & Cell_Address).value2     '하한
                            If XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        Low_Tol_ad(tab_index).Row = Low_Tol_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).judge = True Then
                            XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("J" & Cell_Address).value2     '판정 // 통과/실패
                            If XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        judge_ad(tab_index).Row = judge_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        '====================================================================================================================================
                        '====================================================================================================================================
                        '라인 끝 빈공간 용
                        XL.Sheets(tab_name(tab_index)).Range("BA5000").value = XL.Sheets(Data_sheet).Range("B" & Cell_Address).value & "(" & XL.Sheets(Data_sheet).Range("C" & Cell_Address).value & ")"   '라벨명

                        If XL.Sheets(tab_name(tab_index)).Range("BA5000").value = "()" Then
                            XL.Sheets(tab_name(tab_index)).Range("BA5000").value = ""
                            XL.Sheets(tab_name(tab_index)).Range("BA5000").delete

                            Exit Do
                        End If
                        XL.Sheets(tab_name(tab_index)).Range("BA5000").delete
                        '====================================================================================================================================
                        '====================================================================================================================================
                        '셀주소 하나씩 내리기
                        Select Case input_direction(tab_index)
                            Case "세로"
                                component_ad(tab_index).Row = component_ad(tab_index).Row + 1
                                Label_ad(tab_index).Row = Label_ad(tab_index).Row + 1
                                measure_ad(tab_index).Row = measure_ad(tab_index).Row + 1
                                design_ad(tab_index).Row = design_ad(tab_index).Row + 1
                                error_ad(tab_index).Row = error_ad(tab_index).Row + 1
                                UP_Tol_ad(tab_index).Row = UP_Tol_ad(tab_index).Row + 1
                                Low_Tol_ad(tab_index).Row = Low_Tol_ad(tab_index).Row + 1
                                judge_ad(tab_index).Row = judge_ad(tab_index).Row + 1

                            Case "가로"

                                column_count = column_count + 1
                        End Select

                        '====================================================================================================================================
                        Cell_Address = Cell_Address + 1
                        Line_Count = Line_Count + 1
                        '====================================================================================================================================
                        '====================================================================================================================================
                        If 프로그레스바.ProgressBar1.Value = 프로그레스바.ProgressBar1.Maximum Then
                        Else
                            프로그레스바.ProgressBar1.Value += 1
                        End If
                        '====================================================================================================================================
                    Loop

                    If sum_Line_count = Cell_Address Then        '같은 줄 수 일때 나가기 
                        Exit Do
                    End If

                    tab_index = tab_index + 1         '탭 변경용 탭 인덱스 +1 해주기

                    If Tab_count < tab_index Then Exit Do

                    '    XL.Sheets(origin_sheet_name).Copy(After:=XL.Sheets(tab_index))
                    '  Sheet_Count = Sheet_Count + 1
                    '      XL.activesheet.name = origin_sheet_name & "-" & Sheet_Count


                Loop Until Sheet_Count > Tab_count
                'XL.visible = True
                'XL.Sheets(origin_sheet_name).delete
                XL.Sheets(Data_sheet).delete

                '=========================================================================
                '기본혹은 위치지정 끝
        End Select
        '=========================================================================

        Exit Sub
XLC:

        XL.Workbooks(1).close
        XL.Quit
        Data_error_occur = 1
    End Sub

    Sub Extension_type_2()          'ASC 파일 읽어오기
        'On Error GoTo XLC
        Dim input_string As String
        Dim merge_count As Long
        Dim cRow As Integer
        Dim i As Integer

        Select Case Label20.Text
        '=========================================================================

            Case "기본폼1", "기본폼2", "기본폼3", "기본폼4", "기본폼5"
                Dim error_arry() As String

                Dim Data_sheet As String
                Dim form_txt_1 As String
                Dim form_txt_2 As String
                Dim sheets_switch As Integer
                Dim sheets_del As Integer

                sheets_switch = 0

                Select Case Label20.Text

                    Case "기본폼1"        '말머리 전부 포함               1사용
                        form_txt_1 = "기본폼1"
                        form_txt_2 = form_txt_1
                        sheets_switch = 0

                    Case "기본폼2"        '말머리 전부 없음               2 사용
                        form_txt_1 = "기본폼2"
                        form_txt_2 = form_txt_1
                        sheets_switch = 1

                    Case "기본폼3"        '그림 전부 삽입                 3 사용
                        form_txt_1 = "기본폼3"
                        form_txt_2 = form_txt_1
                        sheets_switch = 2

                    Case "기본폼4"        '첫 페이지 말머리 있음 2 페이지 부터 말머리 없음            1,2사용
                        form_txt_1 = "기본폼1"
                        form_txt_2 = "기본폼2"
                        sheets_switch = 0
                        sheets_del = 1
                    Case "기본폼5"         '첫페이지 그림 삽입, 2 페이지부터 그림, 말머리 없음         2,3 사용
                        form_txt_1 = "기본폼3"
                        form_txt_2 = "기본폼2"
                        sheets_switch = 2
                        sheets_del = 1
                    Case Else

                        form_txt_1 = "기본폼1"
                        form_txt_2 = "기본폼1"

                End Select

                Data_sheet = "Data Sheet"

                error_arry = Split("평면도,위치도,동심도,평행도,직각도,동축도,면의 위치도,경사도", ",")

                '=========================================================================


                XL.Workbooks.open(Open_Dir)       '성적서 오픈
                'XL.visible = True
                XL.DisplayAlerts = False

                XL.Sheets.add(before:=XL.Sheets("기본폼1")) 'asc파일 가져올 워크시트 추가
                XL.activesheet.name = Data_sheet

                FileOpen(2, CSV_Dir, OpenMode.Input)


                Do Until EOF(2)
                    cRow = cRow + 1
                    input_string = LineInput(2)
                    Dim input_arry() As String = Split(input_string, ";")
                    For i = 0 To UBound(input_arry)
                        XL.sheets(Data_sheet).cells(cRow, i + 1).value = input_arry(i)
                    Next
                Loop

                FileClose(2)
                ' XL.visible = True
                '================================================================기본 유저 정보 입력
                Input_user_info()
                Sheet_del()
                '================================================================기본 유저 정보 입력
                XL.Sheets(form_txt_1).Copy(After:=XL.Sheets(Data_sheet))        '데이터 입력용 워크시트 복사
                XL.activesheet.name = "DATA-1"                                         '워크시트 선택
                XL.sheets("DATA-1").select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 1                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(Data_sheet).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(Data_sheet).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp


                Select Case Label20.Text                    '페이지수 계산

                    Case "기본폼1"
                        Quetient = Cell_Count / 28              '페이지수 계산  폼1 : 28, 폼2 : 32, 폼3 : 15
                    Case "기본폼2"        '말머리 전부 없음               2 사용
                        Quetient = Cell_Count / 32
                    Case "기본폼3"        '그림 전부 삽입                 3 사용
                        Quetient = Cell_Count / 15
                    Case "기본폼4"        '첫 페이지 말머리 있음 2 페이지 부터 말머리 없음            1,2사용
                        Quetient = ((Cell_Count - 28) / 32) + 1
                    Case "기본폼5"         '첫페이지 그림 삽입, 2 페이지부터 그림, 말머리 없음         2,3 사용
                        Quetient = ((Cell_Count - 15) / 32) + 1
                    Case Else

                End Select

                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                프로그레스바.ProgressBar1.Maximum = Cell_Count

                Dim count As Integer

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력
                    Select Case sheets_switch
                        Case 0      '기본폼1
                            Line_Count = 9
                        Case 1      '기본폼2
                            Line_Count = 5
                        Case 2      '기본폼 3
                            Line_Count = 22
                    End Select

                    Do Until Line_Count > 36

                        XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("A" & Cell_Address).value & ")"   '라벨명
                        XL.Sheets("DATA-" & Sheet_Count).Range("B" & Line_Count).value = XL.Sheets(1).Range("C" & Cell_Address).value     '구성요소
                        XL.Sheets("DATA-" & Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(1).Range("G" & Cell_Address).value     '측정값
                        XL.Sheets("DATA-" & Sheet_Count).Range("D" & Line_Count).value = XL.Sheets(1).Range("D" & Cell_Address).value     '설계치
                        XL.Sheets("DATA-" & Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(1).Range("H" & Cell_Address).value     '오차
                        XL.Sheets("DATA-" & Sheet_Count).Range("F" & Line_Count).value = XL.Sheets(1).Range("E" & Cell_Address).value     '상한
                        XL.Sheets("DATA-" & Sheet_Count).Range("G" & Line_Count).value = XL.Sheets(1).Range("F" & Cell_Address).value     '하한
                        XL.Sheets("DATA-" & Sheet_Count).Range("H" & Line_Count).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        If XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = "()" Then
                            XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = ""
                            Exit Do
                        End If
                        Line_Count = Line_Count + 1
                        Cell_Address = Cell_Address + 1
                        count += 1

                        '====================================================================================================================================
                        If 프로그레스바.ProgressBar1.Value = 프로그레스바.ProgressBar1.Maximum Then
                        Else
                            프로그레스바.ProgressBar1.Value += 1
                        End If
                        '====================================================================================================================================

                    Loop

                    Select Case sheets_switch
                        Case 0, 2           '기본폼1, 기본폼3
                            XL.sheets("DATA-" & Sheet_Count).range("I3").value = Sheet_Count & "/" & TotalSheet  '페이지 번호 입력

                        Case 1      '기본폼2
                            XL.sheets("DATA-" & Sheet_Count).range("I1").value = Sheet_Count & "/" & TotalSheet  '페이지 번호 입력
                    End Select

                    If Sheet_Count = TotalSheet Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If


                    If Sheet_Count = 0 Then
                        XL.Sheets(form_txt_1).Copy(After:=XL.Sheets("DATA-" & Sheet_Count))
                        XL.activesheet.name = "DATA-" & (Sheet_Count + 1)
                    Else
                        XL.Sheets(form_txt_2).Copy(After:=XL.Sheets("DATA-" & Sheet_Count))
                        XL.activesheet.name = "DATA-" & (Sheet_Count + 1)
                        Select Case Label20.Text
                            Case "기본폼4"
                                sheets_switch = 1
                            Case "기본폼5"
                                sheets_switch = 1
                        End Select

                    End If
                    Sheet_Count = Sheet_Count + 1



                Loop Until Sheet_Count > (TotalSheet)

                XL.Sheets(form_txt_1).delete
                If sheets_del = 1 Then
                    XL.Sheets(form_txt_2).delete
                End If
                XL.Sheets(Data_sheet).delete


             '=========================================================================
            Case "위치 지정"
                '=========================================================================

                Dim Label_ad() As address
                Dim measure_ad() As address
                Dim component_ad() As address
                Dim design_ad() As address
                Dim error_ad() As address
                Dim UP_Tol_ad() As address
                Dim Low_Tol_ad() As address
                Dim judge_ad() As address
                Dim Line_count_ad() As Integer
                Dim column_count As Integer
                Dim Result_Form_dir As String
                Dim select_check_value() As check_value
                Dim ini_tab_index As String                'ini 섹션 값 구분용
                Dim tab_index As Integer                    ' 원본 엑셀 탭 선택용
                Dim origin_sheet_name() As String
                Dim Data_sheet As String
                Dim input_direction() As String
                Dim Tab_count As Integer        ' 탭 개수 인식
                Dim tab_name() As String
                Dim sum_Line_count As Integer
                Dim sum_temp As Integer
                Dim error_arry() As String
                'Dim error_txt As String


                Result_Form_dir = GetINIValue("custom_match_info", "Result_Form_dir", ini_dir)

                Tab_count = GetINIValue("custom_match_info", "tab_count", ini_dir)

                ReDim Label_ad(Tab_count)
                ReDim measure_ad(Tab_count)
                ReDim component_ad(Tab_count)
                ReDim design_ad(Tab_count)
                ReDim error_ad(Tab_count)
                ReDim UP_Tol_ad(Tab_count)
                ReDim Low_Tol_ad(Tab_count)
                ReDim judge_ad(Tab_count)
                ReDim Line_count_ad(Tab_count)
                ReDim select_check_value(Tab_count)
                ReDim tab_name(Tab_count)
                ReDim origin_sheet_name(Tab_count)
                ReDim input_direction(Tab_count)


                For tab_index = 1 To Tab_count
                    ini_tab_index = "tab_" & tab_index
                    tab_name(tab_index) = GetINIValue(ini_tab_index, "tab_name", ini_dir)
                    Label_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Label", ini_dir))
                    component_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "component", ini_dir))
                    measure_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "measure_value", ini_dir))
                    design_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Design_value", ini_dir))
                    error_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "error", ini_dir))
                    UP_Tol_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "UP_Tol", ini_dir))
                    Low_Tol_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "Low_Tol", ini_dir))
                    judge_ad(tab_index).Col = Ad_Str(GetINIValue(ini_tab_index, "judge", ini_dir))

                    Label_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Label", ini_dir))
                    component_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "component", ini_dir))
                    measure_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "measure_value", ini_dir))
                    design_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Design_value", ini_dir))
                    error_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "error", ini_dir))
                    UP_Tol_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "UP_Tol", ini_dir))
                    Low_Tol_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "Low_Tol", ini_dir))
                    judge_ad(tab_index).Row = Ad_NUM(GetINIValue(ini_tab_index, "judge", ini_dir))

                    Line_count_ad(tab_index) = GetINIValue(ini_tab_index, "Line_count", ini_dir)

                    input_direction(tab_index) = GetINIValue(ini_tab_index, "input_direction", ini_dir)


                    select_check_value(tab_index).label = GetINIValue(ini_tab_index, "label_check", ini_dir)                                                          ' 탭마다 ini 별도로 정보 저장후 읽어오기, 키 이름 _check 추가해서 다른것과 구별 
                    select_check_value(tab_index).Measure_value = GetINIValue(ini_tab_index, "Measure_value_check", ini_dir)
                    select_check_value(tab_index).Design_value = GetINIValue(ini_tab_index, "Design_value_check", ini_dir)
                    select_check_value(tab_index).Error_check = GetINIValue(ini_tab_index, "error_check", ini_dir)
                    select_check_value(tab_index).UP_tol = GetINIValue(ini_tab_index, "UP_tol_check", ini_dir)
                    select_check_value(tab_index).Low_tol = GetINIValue(ini_tab_index, "Low_tol_check", ini_dir)
                    select_check_value(tab_index).judge = GetINIValue(ini_tab_index, "judge_check", ini_dir)
                    select_check_value(tab_index).component = GetINIValue(ini_tab_index, "component_check", ini_dir)

                    origin_sheet_name(tab_index) = GetINIValue(ini_tab_index, "tab_name", ini_dir)

                Next tab_index

                For Each sum_temp In Line_count_ad
                    sum_Line_count = sum_Line_count + sum_temp
                Next



                error_arry = Split("평면도,위치도,동심도,평행도,직각도,동축도,면의 위치도,경사도", ",")

                XL.Workbooks.open(Result_Form_dir)       '성적서 오픈        원본성적서 지정
                XL.DisplayAlerts = False

                'origin_sheet_name = XL.activesheet.name
                Data_sheet = "Data_sheet"
                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
                XL.activesheet.name = Data_sheet

                FileOpen(2, CSV_Dir, OpenMode.Input)


                Do Until EOF(2)
                    cRow = cRow + 1
                    input_string = LineInput(2)
                    Dim input_arry() As String = Split(input_string, ";")
                    For i = 0 To UBound(input_arry)
                        XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                    Next
                Loop

                FileClose(2)
                'XL.visible = True

                tab_index = 1

                '===========================================위치지정 추가기입 삽입 위치.

                Call add_str()

                '===========================================위치지정 추가기입 삽입 위치.

                '   XL.Sheets(origin_sheet_name).Copy(After:=XL.Sheets("Data_sheet"))        '데이터 입력용 워크시트 복사
                'XL.activesheet.name = tab_name(1)                         '데이터 입력용 워크시트 선택

                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 1                     '데이터 입력 워크시트 카운트

                Cell_Count2 = XL.Sheets(Data_sheet).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(Data_sheet).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                'Quetient = Cell_Count / Line_count_ad              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                프로그레스바.ProgressBar1.Maximum = Cell_Count

                '  If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 1
                    column_count = 1


                    Do Until Line_Count > Line_count_ad(tab_index)

                        If select_check_value(tab_index).label = True Then
                            XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("B" & Cell_Address).value2 & "(" & XL.Sheets(Data_sheet).Range("A" & Cell_Address).value2 & ")"   '라벨명
                            If XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        Label_ad(tab_index).Row = Label_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name).Range(Label_ad(tab_index).Col & Label_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).component = True Then
                            XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("C" & Cell_Address).value2     '구성요소
                            If XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        component_ad(tab_index).Row = component_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(component_ad(tab_index).Col & component_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If

                        End If


                        If select_check_value(tab_index).Measure_value = True Then
                            XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).value = XL.Sheets(Data_sheet).Range("G" & Cell_Address).value      '측정값
                            If XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        measure_ad(tab_index).Row = measure_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(measure_ad(tab_index).Col & measure_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If

                        End If

                        If select_check_value(tab_index).Design_value = True Then
                            XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("D" & Cell_Address).value2     '설계치
                            If XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        design_ad(tab_index).Row = design_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(design_ad(tab_index).Col & design_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        '   error_txt = XL.Sheets("DATA-" & Sheet_Count).Range("B" & Line_Count).value      'SC파일은 기하공차가 오차로 나오지 않아 필요 없음

                        If select_check_value(tab_index).Error_check = True Then
                            XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("H" & Cell_Address).value2     '오차
                            If XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        error_ad(tab_index).Row = error_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(error_ad(tab_index).Col & error_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).UP_tol = True Then
                            XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("E" & Cell_Address).value2     '상한
                            If XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        UP_Tol_ad(tab_index).Row = UP_Tol_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(UP_Tol_ad(tab_index).Col & UP_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).Low_tol = True Then
                            XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("F" & Cell_Address).value2     '하한
                            If XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        Low_Tol_ad(tab_index).Row = Low_Tol_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(Low_Tol_ad(tab_index).Col & Low_Tol_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        If select_check_value(tab_index).judge = True Then
                            XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).value2 = XL.Sheets(Data_sheet).Range("J" & Cell_Address).value2     '판정 // 통과/실패
                            If XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergecells = True Then
                                Select Case input_direction(tab_index)
                                    Case "세로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergearea.rows.count
                                        judge_ad(tab_index).Row = judge_ad(tab_index).Row + (merge_count - 1)
                                    Case "가로"
                                        merge_count = XL.Sheets(tab_name(tab_index)).Range(judge_ad(tab_index).Col & judge_ad(tab_index).Row).cells(1, column_count).mergearea.columns.count
                                        column_count = column_count + (merge_count - 1)
                                End Select
                            End If
                        End If

                        '====================================================================================================================================
                        '====================================================================================================================================
                        '라인 끝 빈공간 용
                        XL.Sheets(tab_name(tab_index)).Range("BA5000").value = XL.Sheets(Data_sheet).Range("B" & Cell_Address).value2 & "(" & XL.Sheets(Data_sheet).Range("A" & Cell_Address).value2 & ")"   '라벨명

                        If XL.Sheets(tab_name(tab_index)).Range("BA5000").value = "()" Then
                            XL.Sheets(tab_name(tab_index)).Range("BA5000").value = ""
                            XL.Sheets(tab_name(tab_index)).Range("BA5000").delete

                            Exit Do
                        End If
                        XL.Sheets(tab_name(tab_index)).Range("BA5000").delete
                        '====================================================================================================================================
                        '====================================================================================================================================
                        '셀주소 하나씩 내리기
                        Select Case input_direction(tab_index)
                            Case "세로"
                                component_ad(tab_index).Row = component_ad(tab_index).Row + 1
                                Label_ad(tab_index).Row = Label_ad(tab_index).Row + 1
                                measure_ad(tab_index).Row = measure_ad(tab_index).Row + 1
                                design_ad(tab_index).Row = design_ad(tab_index).Row + 1
                                error_ad(tab_index).Row = error_ad(tab_index).Row + 1
                                UP_Tol_ad(tab_index).Row = UP_Tol_ad(tab_index).Row + 1
                                Low_Tol_ad(tab_index).Row = Low_Tol_ad(tab_index).Row + 1
                                judge_ad(tab_index).Row = judge_ad(tab_index).Row + 1

                            Case "가로"

                                column_count = column_count + 1
                        End Select

                        '====================================================================================================================================
                        Cell_Address = Cell_Address + 1
                        Line_Count = Line_Count + 1
                        '====================================================================================================================================
                        '====================================================================================================================================
                        If 프로그레스바.ProgressBar1.Value = 프로그레스바.ProgressBar1.Maximum Then
                        Else
                            프로그레스바.ProgressBar1.Value += 1
                        End If
                        '====================================================================================================================================
                    Loop

                    If sum_Line_count = Cell_Address Then        '같은 줄 수 일때 나가기 
                        Exit Do
                    End If

                    tab_index = tab_index + 1         '탭 변경용 탭 인덱스 +1 해주기

                    If Tab_count < tab_index Then Exit Do

                    '    XL.Sheets(origin_sheet_name).Copy(After:=XL.Sheets(tab_index))
                    '  Sheet_Count = Sheet_Count + 1
                    '      XL.activesheet.name = origin_sheet_name & "-" & Sheet_Count


                Loop Until Sheet_Count > Tab_count
                'XL.visible = True
                'XL.Sheets(origin_sheet_name).delete
                XL.Sheets(Data_sheet).delete

                '=========================================================================
                '기본혹은 위치지정 끝
        End Select
        '=========================================================================

        Exit Sub
XLC:

        XL.Workbooks(1).close
        XL.Quit
        Data_error_occur = 1
    End Sub
    Private Sub Label20_Changed(sender As Object, e As EventArgs) Handles Label20.TextChanged
        Select Case Label20.Text
            Case "기본폼1", "기본폼2", "기본폼3", "기본폼4", "기본폼5"
                TabControl1.Visible = True

                TabControl2.Visible = False                 '   2021-10-15 penal 제거후 tabcontrol2로 제어


                '===========================================2021-07-27 Panel 추가 해서 Panel안에 다 넣음
                'CheckBox3.Visible = False
                'CheckBox4.Visible = False
                'CheckBox5.Visible = False
                'CheckBox6.Visible = False
                'CheckBox7.Visible = False
                'CheckBox8.Visible = False
                'CheckBox9.Visible = False
                'CheckBox10.Visible = False
                '
                'Label24.Visible = False
                'Label25.Visible = False
                'La bel26.Visible = False
                'Label27.Visible = False
                'Label28.Visible = False
                'Label29.Visible = False
                'Label30.Visible = False
                'Label31.Visible = False
                'Label32.Visible = False
                'Label33.Visible = False
                'Label34.Visible = False
                'Label35.Visible = False
                'Label36.Visible = False
                'Label47.Visible = False
                'Label48.Visible = False
                '
                'CheckBox3.Checked = False
                'CheckBox4.Checked = False
                'CheckBox5.Checked = False
                'CheckBox6.Checked = False
                'CheckBox7.Checked = False
                'CheckBox8.Checked = False
                'CheckBox9.Checked = False
                'CheckBox10.Checked = False
                '===========================================2021-07-27 Panel 추가 해서 Panel안에 다 넣음
                Select Case Label20.Text
                    Case "기본폼1"
                        TabControl1.SelectTab(0)
                    Case "기본폼2"
                        TabControl1.SelectTab(1)
                    Case "기본폼3"
                        TabControl1.SelectTab(2)
                    Case "기본폼4"
                        TabControl1.SelectTab(3)
                    Case "기본폼5"
                        TabControl1.SelectTab(4)
                End Select

            Case "위치 지정"
                TabControl1.Visible = False

                TabControl2.Visible = True

                '===========================================2021-07-27 Panel 추가 해서 Panel안에 다 넣음
                'CheckBox3.Visible = True
                ''CheckBox4.Visible = True
                ' CheckBox5.Visible = True
                ' CheckBox6.Visible = True
                ' CheckBox7.Visible = True
                ' CheckBox8.Visible = True
                ' CheckBox9.Visible = True
                ''CheckBox10.Visible = True
                '
                ' Label24.Visible = True
                ' Label25.Visible = True
                ' Label26.Visible = True
                ' Label27.Visible = True
                ' Label28.Visible = True
                ' Label29.Visible = True
                ' Label30.Visible = True
                ' Label31.Visible = True
                ' Label32.Visible = True
                ' Label33.Visible = True
                ' Label34.Visible = True
                ' Label35.Visible = True
                ' Label36.Visible = True
                ' Label47.Visible = True
                ' Label48.Visible = True
                '===========================================2021-07-27 Panel 추가 해서 Panel안에 다 넣음
        End Select

    End Sub


    Private Sub 폴더위치열기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 폴더위치열기ToolStripMenuItem.Click
        Dim file_path As String
        Dim file_len As Integer

        file_len = InStrRev(Label16.Text, "\")
        file_path = Strings.Left(Label16.Text, file_len - 1)

        If file_path <> "" Then
            Shell("explorer.exe " & file_path, AppWinStyle.NormalFocus)
        End If

    End Sub
    Private Sub Source파일열기ToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles Source파일열기ToolStripMenuItem.Click

        Dim file_path As String

        file_path = Label16.Text

        If file_path <> "N/A" Then
            Shell("explorer.exe " & file_path, AppWinStyle.NormalFocus)
        End If
    End Sub

    Private Sub 폴더위치열기ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 폴더위치열기ToolStripMenuItem1.Click
        Dim file_path As String

        file_path = Label17.Text

        If file_path <> "" Then
            Shell("explorer.exe " & file_path, AppWinStyle.NormalFocus)
        End If

    End Sub
    Private Sub 폴더위치열기ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 폴더위치열기ToolStripMenuItem2.Click
        Dim file_path As String
        Dim file_len As Integer

        file_path = GetINIValue("custom_match_info", "Result_Form_Dir", ini_dir)

        file_len = InStrRev(file_path, "\")
        file_path = Strings.Left(file_path, file_len - 1)
        If file_path <> "" Then
            Shell("explorer.exe " & file_path, AppWinStyle.NormalFocus)
        End If

    End Sub
    Private Sub 원본성적서열기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 원본성적서열기ToolStripMenuItem.Click

        Dim file_path As String

        file_path = GetINIValue("custom_match_info", "Result_Form_Dir", ini_dir)

        If file_path <> "N/A" Then
            Shell("explorer.exe " & file_path, AppWinStyle.NormalFocus)
        End If

    End Sub
    Sub Input_user_info()
        'On Error GoTo missing
        '===================================================== 기본폼1 기입
        'XL.visible = True
        XL.sheets("기본폼1").select
        If Product_Name <> "" Then
            XL.Sheets("기본폼1").Range("D4").value = Product_Name
        End If
        If Machine_Name <> "" Then
            XL.Sheets("기본폼1").Range("D5").value = Machine_Name
        End If
        If Request_Dept <> "" Then
            XL.Sheets("기본폼1").Range("D6").value = Request_Dept
        End If
        If Request_Date <> "" Then
            XL.Sheets("기본폼1").Range("D7").value = Request_Date
        End If
        If Drawing_Num <> "" Then
            XL.Sheets("기본폼1").Range("H4").value = Drawing_Num
        End If
        If Program_Name <> "" Then
            XL.Sheets("기본폼1").Range("H5").value = Program_Name
        End If
        If Player_Name <> "" Then
            XL.Sheets("기본폼1").Range("H6").value = Player_Name
        End If
        If Measure_Date <> "" Then
            XL.Sheets("기본폼1").Range("H7").value = Measure_Date
        End If

        Dim Pic_L As Single
        Dim Pic_T As Single 'top
        Dim Pic_W As Single
        Dim Pic_H As Single
        Dim Pic_L2 As Single
        Dim Pic_T2 As Single 'top
        Dim Pic_W2 As Single
        Dim Pic_H2 As Single

        Pic_L = XL.sheets("기본폼1").Range("A1").Left
        Pic_T = XL.sheets("기본폼1").Range("A1").Top
        Pic_W = 180
        Pic_H = 45
        If select_pic_name <> "" Then

            ' With XL.ActiveSheet.Shapes.Range("Picture 1")
            '.Select
            '.Delete
            'End With

            XL.sheets("기본폼1").shapes.addpicture(fileName:=select_pic_name, Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L, Top:=Pic_T, Width:=Pic_W, Height:=Pic_H).select
        End If

        Pic_L2 = XL.sheets("기본폼1").Range("H1:I2").Left
        Pic_T2 = XL.sheets("기본폼1").Range("H1:I2").Top
        Pic_W2 = XL.sheets("기본폼1").Range("H1:I2").width - 6
        Pic_H2 = XL.sheets("기본폼1").Range("H1:I2").height - 20

        If select_pic_name_2 <> "" Then

            XL.sheets("기본폼1").shapes.addpicture(fileName:=select_pic_name_2, Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L2, Top:=Pic_T2, Width:=Pic_W2, Height:=Pic_H2).select

            With XL.selection.ShapeRange
                '.Width = 180.0
                '.Height = 45.0
                .IncrementTop(10)
                .IncrementLeft(3)
            End With

        End If
        '===================================================== 기본폼2 기입

        XL.sheets("기본폼2").select
        If Product_Name <> "" Then
            XL.Sheets("기본폼2").Range("D2").value = Product_Name
        End If
        ' If Machine_Name <> "" Then
        '  XL.Sheets("기본폼2").Range("D5").value = Machine_Name
        '  End If
        '  If Request_Dept <> "" Then
        '  XL.Sheets("기본폼2").Range("D6").value = Request_Dept
        '  End If
        If Request_Date <> "" Then
            XL.Sheets("기본폼2").Range("D3").value = Request_Date
        End If
        If Drawing_Num <> "" Then
            XL.Sheets("기본폼2").Range("H2").value = Drawing_Num
        End If
        ' If Program_Name <> "" Then
        ' XL.Sheets("기본폼2").Range("H5").value = Program_Name
        ' End If
        'If Player_Name <> "" Then
        ' XL.Sheets("기본폼2").Range("H6").value = Player_Name
        ' End If
        If Measure_Date <> "" Then
            XL.Sheets("기본폼2").Range("H3").value = Measure_Date
        End If

        '===================================================== 기본폼3 기입

        XL.sheets("기본폼3").select
        If Product_Name <> "" Then
            XL.Sheets("기본폼3").Range("D4").value = Product_Name
        End If
        If Machine_Name <> "" Then
            XL.Sheets("기본폼3").Range("D5").value = Machine_Name
        End If
        If Request_Dept <> "" Then
            XL.Sheets("기본폼3").Range("D6").value = Request_Dept
        End If
        If Request_Date <> "" Then
            XL.Sheets("기본폼3").Range("D7").value = Request_Date
        End If
        If Drawing_Num <> "" Then
            XL.Sheets("기본폼3").Range("H4").value = Drawing_Num
        End If
        If Program_Name <> "" Then
            XL.Sheets("기본폼3").Range("H5").value = Program_Name
        End If
        If Player_Name <> "" Then
            XL.Sheets("기본폼3").Range("H6").value = Player_Name
        End If
        If Measure_Date <> "" Then
            XL.Sheets("기본폼3").Range("H7").value = Measure_Date
        End If

        Pic_L = XL.sheets("기본폼3").Range("A1").Left
        Pic_T = XL.sheets("기본폼3").Range("A1").Top
        Pic_W = 180
        Pic_H = 45
        If select_pic_name <> "" Then
            'With XL.sheets("기본폼3").Shapes.Range("Picture 1")
            '.Select
            '.Delete
            'End With
            XL.sheets("기본폼3").shapes.addpicture(fileName:=select_pic_name, Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L, Top:=Pic_T, Width:=Pic_W, Height:=Pic_H).select

            With XL.Selection.ShapeRange
                '.Width = 180.0
                '.Height = 45.0
                .IncrementTop(2)
                .IncrementLeft(2)
            End With

        End If
        'XL.visible = True

        Pic_L2 = XL.sheets("기본폼3").Range("H1:I2").Left
        Pic_T2 = XL.sheets("기본폼3").Range("H1:I2").Top
        Pic_W2 = XL.sheets("기본폼3").Range("H1:I2").width - 6
        Pic_H2 = XL.sheets("기본폼3").Range("H1:I2").height - 20

        If select_pic_name_2 <> "" Then

            XL.sheets("기본폼3").shapes.addpicture(fileName:=select_pic_name_2, Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L2, Top:=Pic_T2, Width:=Pic_W2, Height:=Pic_H2).select

            With XL.selection.ShapeRange
                '.Width = 180.0
                '.Height = 45.0
                .IncrementTop(10)
                .IncrementLeft(3)
            End With

        End If

        Pic_L2 = XL.sheets("기본폼3").Range("A8:I20").Left
        Pic_T2 = XL.sheets("기본폼3").Range("A8:I20").Top
        Pic_W2 = XL.sheets("기본폼3").Range("A8:I20").width - 20
        Pic_H2 = XL.sheets("기본폼3").Range("A8:I20").height - 20

        If select_pic_name_3 <> "" Then

            XL.sheets("기본폼3").shapes.addpicture(fileName:=select_pic_name_3, Linktofile:=0, SaveWithDocument:=-1, Left:=Pic_L2, Top:=Pic_T2, Width:=Pic_W2, Height:=Pic_H2).select

            With XL.selection.ShapeRange
                '.Width = 180.0
                '.Height = 45.0
                .IncrementTop(10)
                .IncrementLeft(10)
            End With

        End If

        Exit Sub

missing:
        error_count = 1

        Resume Next

    End Sub

    Public Function Specialized_play(ByVal folder_dir As String, ByVal ini_Name As String) As Long
        On Error GoTo SPE

        ini_dir = ini_Name

        Input_property(Restore_str(ini_dir))

        '================================

        ReDim strData(10)

        Open_Dir = MRM_root_dir & "\MRM\Data\Resources\Result_Source.xlsx"   '원본 엑셀

        Dim fFindFile As New System.IO.FileInfo(Open_Dir)
        If fFindFile.Exists = False Then

            Call Origin_form()
        End If

        XL = CreateObject("Excel.application")

        CSV_Dir = GetINIValue("Matching_info", "CSV_file_Path", Restore_str(ini_dir))
        '=========================================================================
        '프로그래스바
        프로그레스바.Location = New Point(500, 500)
        프로그레스바.Show()
        '=========================================================================
        Dim type_value As String
        type_value = UCase(Strings.Right(CSV_Dir, 3))
        Select Case type_value
            Case "CSV"
                source_type = 1
            Case "ASC"
                source_type = 2
        End Select

        Select Case source_type

            Case 1
                Call Extension_type_1()         'csv    
            Case 2
                Call Extension_type_2()         'asc
        End Select

        Select Case strDate
            Case "True"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & DateString & "_" & Format(TimeOfDay, "HH-mm-ss") & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & DateString & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                End Select
            Case "False"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & "_" & Format(TimeOfDay, "HH-mm-ss''") & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", Restore_str(ini_dir)) & "\" & GetINIValue("Matching_info", "Save_File_Name", Restore_str(ini_dir)) & GetINIValue("Matching_info", "Save_Type", Restore_str(ini_dir))
                End Select

        End Select

        WPPS("Matching_Info", "Last_Paly_date", Now(), Restore_str(ini_dir))         '마지막 측정 시간 저장
        프로그레스바.Close()

        tempsave_dir = MRM_root_dir & "\MRM\Result\temp"
        Dim fFindFolder As New System.IO.DirectoryInfo(tempsave_dir)   '  폴더 존재 유무 확인용 선언

        Select Case Label22.Text
            Case "예"

                Select Case Label19.Text
                    Case ".xlsx"
                        With XL
                            .DisplayAlerts = False
                            .workbooks(1).SaveAS(filename:=Restore_str(Save_Dir))          '다른이름으로 저장 위치 지정
                            .Workbooks(1).close
                            .quit
                        End With

                    Case ".PDF"
                        With XL
                            .DisplayAlerts = False
                            .workbooks(1).ExportAsFixedFormat(Type:=0, Filename:=Restore_str(Save_Dir), Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False)
                            If fFindFolder.Exists = False Then
                                MkDir(MRM_root_dir & "\MRM\Result\temp")
                            End If
                            tempsave_dir = MRM_root_dir & "\MRM\Result\temp"
                            .workbooks(1).SaveAS(tempsave_dir & "\temp-auto.xlsx")
                            .Workbooks(1).close
                            .quit
                        End With
                End Select

                XL = Nothing
                '=============================================================================== 에러 구문 출력 
                If error_count <> 0 Then
                    Select Case error_count

                        Case 1          '링크 사진 경로 삭제 및 이동.
                            MsgBox("첨부한 그림의 경로가 변경되었거나 삭제되었습니다." & Environment.NewLine & "첨부 그림 경로나 파일 존재 유무를 다시 한번 확인해 주세요.",, "INTERAL ERROR OCCURRED - 01")

                    End Select
                End If
                '=============================================================================== 에러 구문 출력 
                If CheckBox1.Checked = True Then
                    Me.Close()
                End If

            Case "아니오"
                MsgBox("자동저장을 선택하지 않으셨습니다. 성적서가 화면에 켜집니다." & Environment.NewLine & Environment.NewLine & "성적서를 따로 저장해주세요.",, "성적서 자동 저장 취소")
                XL.DisplayAlerts = False
                If fFindFolder.Exists = False Then
                    MkDir(MRM_root_dir & "\MRM\Result\temp")
                End If
                XL.workbooks(1).SaveAS(tempsave_dir & "\temp.xlsx")
                XL.visible = True
                '=============================================================================== 에러 구문 출력 
                If error_count <> 0 Then
                    Select Case error_count

                        Case 1          '링크 사진 경로 삭제 및 이동.
                            MsgBox("첨부한 그림의 경로가 변경되었거나 삭제 되었습니다." & Environment.NewLine & "그림 경로나 그림 존재 유무를 다시 한번 확인해 주세요.",, "INTERAL ERROR OCCURRED - 01")

                    End Select
                End If
                '=============================================================================== 에러 구문 출력 
                If CheckBox1.Checked = True Then
                    Me.Close()
                End If
        End Select
        Specialized_play = 0
        Exit Function

SPE:
        Specialized_play = 1

    End Function

    Sub Input_property(ByVal ini_dir As String)                 '리스트 선택시 라벨 표시

        strDate = GetINIValue("Matching_info", "Check_Date", ini_dir)
        strTime = GetINIValue("Matching_info", "Check_Time", ini_dir)
        Label16.Text = GetINIValue("Matching_info", "CSV_file_Path", ini_dir)
        Label17.Text = GetINIValue("Matching_info", "Save_File_Path", ini_dir) '& "\" & GetINIValue("Matching_info", "Save_File_Name", MRM_root_dir & "\Data\Resources\Ini\" & ListBox1.SelectedItem.ToString & ".ini") & GetINIValue("Matching_info", "Save_Type", MRM_root_dir & "\Data\Resources\Ini\" & ListBox1.SelectedItem.ToString & ".ini")
        auto_save_check = GetINIValue("Matching_info", "auto_save", ini_dir)

        Select Case auto_save_check
            Case "True"
                Label22.Text = "예"
            Case "False"
                Label22.Text = "아니오"
        End Select

        Select Case strDate
            Case "True"

                Select Case strTime
                    Case "True"
                        Label18.Text = GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "   +현재 날짜 +현재 시간"
                    Case "False"
                        Label18.Text = GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "   +현재 날짜"
                End Select
            Case "False"
                Select Case strTime
                    Case "True"
                        Label18.Text = GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "   +현재 시간"
                    Case "False"
                        Label18.Text = GetINIValue("Matching_info", "Save_File_Name", ini_dir)
                End Select
        End Select

        Label19.Text = GetINIValue("Matching_info", "Save_Type", ini_dir)
        Label20.Text = GetINIValue("Matching_info", "Basic_form_seleted", ini_dir)   '
        iniPath = GetINIValue("Matching_info", "ini_dir", ini_dir)
        logo_path = GetINIValue("Matching_info", "logo_path", ini_dir)

        Select Case Label20.Text
            Case "기본폼1", "기본폼2", "기본폼3", "기본폼4", "기본폼5"
                Dim user_section As String
                Dim user_keyname() As String
                ReDim user_keyname(12)

                user_section = "user_info"
                user_keyname(0) = "Product_Name"
                user_keyname(1) = "Machine_Name"
                user_keyname(2) = "Request_Dept"
                user_keyname(3) = "Request_Date"
                user_keyname(4) = "Drawing_Num"
                user_keyname(5) = "Program_Name"
                user_keyname(6) = "Player_Name"
                user_keyname(7) = "Measure_Date"
                user_keyname(8) = "Check_date_1"
                user_keyname(9) = "Check_date_2"
                user_keyname(10) = "Select_pic_name"
                user_keyname(11) = "Select_pic_name_2"
                user_keyname(12) = "Select_pic_name_3"

                Product_Name = GetINIValue(user_section, user_keyname(0), ini_dir)
                Machine_Name = GetINIValue(user_section, user_keyname(1), ini_dir)
                Request_Dept = GetINIValue(user_section, user_keyname(2), ini_dir)

                Request_Date = GetINIValue(user_section, user_keyname(3), ini_dir)

                Drawing_Num = GetINIValue(user_section, user_keyname(4), ini_dir)
                Program_Name = GetINIValue(user_section, user_keyname(5), ini_dir)
                Player_Name = GetINIValue(user_section, user_keyname(6), ini_dir)

                Measure_Date = GetINIValue(user_section, user_keyname(7), ini_dir)

                select_pic_name = GetINIValue(user_section, user_keyname(10), ini_dir)

                select_pic_name_2 = GetINIValue(user_section, user_keyname(11), ini_dir)
                select_pic_name_3 = GetINIValue(user_section, user_keyname(12), ini_dir)

            Case "위치 지정"

                Dim custom_section As String
                Dim custom_Keyname(5) As String
                Dim Check_Keyname(10) As String
                Dim address_Keyname(10) As String

                Dim tab_selection As String
                Dim tab_count As Integer     ' 탭 개수 인식
                Dim i As Integer
                Dim Selected_tab As TabPage


                Dim tab_name() As String
                Dim Label_ad() As String
                Dim measure_ad() As String
                Dim component_ad() As String
                Dim design_ad() As String
                Dim error_ad() As String
                Dim UP_Tol_ad() As String
                Dim Low_Tol_ad() As String
                Dim judge_ad() As String
                Dim Line_count_ad() As Integer
                Dim origin_sheet_name() As String
                Dim input_direction() As String
                Dim add_Control() As control_structure
                Dim control_add_num As Integer
                Dim check_text(10) As String


                custom_section = "custom_match_info"
                Check_Keyname(0) = "label_check"
                Check_Keyname(1) = "component_check"
                Check_Keyname(2) = "measure_value_check"
                Check_Keyname(3) = "Design_value_check"
                Check_Keyname(4) = "UP_tol_check"
                Check_Keyname(5) = "Low_tol_check"
                Check_Keyname(6) = "error_check"
                Check_Keyname(7) = "judge_check"

                address_Keyname(0) = "label"
                address_Keyname(1) = "component"
                address_Keyname(2) = "measure_value"
                address_Keyname(3) = "Design_value"
                address_Keyname(4) = "UP_tol"
                address_Keyname(5) = "Low_tol"
                address_Keyname(6) = "error"
                address_Keyname(7) = "judge"

                custom_Keyname(0) = "Result_Form_Dir"
                custom_Keyname(1) = "line_count"
                custom_Keyname(2) = "input_direction"
                custom_Keyname(3) = "tab_count"
                custom_Keyname(4) = "tab_name"


                check_text(1) = "라벨명         : "
                check_text(2) = "요소             : "
                check_text(3) = "측정값         : "
                check_text(4) = "설계치         : "
                check_text(5) = "상한 공차    :"
                check_text(6) = "하한 공차    :"
                check_text(7) = "오차             :"
                check_text(8) = "판정             :"

                TabControl2.Controls.Clear()
                tab_count = GetINIValue(custom_section, custom_Keyname(3), ini_dir)

                ReDim Label_ad(tab_count)
                ReDim measure_ad(tab_count)
                ReDim component_ad(tab_count)
                ReDim design_ad(tab_count)
                ReDim error_ad(tab_count)
                ReDim UP_Tol_ad(tab_count)
                ReDim Low_Tol_ad(tab_count)
                ReDim judge_ad(tab_count)
                ReDim Line_count_ad(tab_count)
                ReDim tab_name(tab_count)
                ReDim origin_sheet_name(tab_count)
                ReDim input_direction(tab_count)
                ReDim add_Control(tab_count)



                For i = 1 To tab_count
                    tab_selection = "tab_" & i
                    tab_name(i) = GetINIValue(tab_selection, custom_Keyname(4), ini_dir)
                    TabControl2.TabPages.Add(tab_name(i))

                    ReDim add_Control(i).check_box(15)
                    ReDim add_Control(i).Text_box(15)
                    ReDim add_Control(i).label(15)
                Next i
                For i = 1 To tab_count                      ' 탭마다 컨트롤 생성 및 배치
                    tab_selection = "tab_" & i
                    For control_add_num = 1 To 15

                        add_Control(i).check_box(control_add_num) = New CheckBox
                        add_Control(i).Text_box(control_add_num) = New TextBox
                        add_Control(i).label(control_add_num) = New Label
                    Next control_add_num


                    TabControl2.SelectedIndex = i - 1
                    Selected_tab = TabControl2.SelectedTab
                    '======================================================================컨트롤 위치 고정용 
                    For control_add_num = 1 To 8
                        'TabControl2.SelectedTab = TabControl2.TabPages(tab_name(i))
                        Selected_tab.BackColor = SystemColors.Window
                        ' TabControl2.SelectedTab.Controls.Add(add_Control(i).check_box(control_add_num))
                        Selected_tab.Controls.Add(add_Control(i).check_box(control_add_num))
                        add_Control(i).check_box(control_add_num).Enabled = True
                        add_Control(i).check_box(control_add_num).Top = 25 * control_add_num
                        add_Control(i).check_box(control_add_num).Left = 20
                        add_Control(i).check_box(control_add_num).Height = 20
                        add_Control(i).check_box(control_add_num).Width = 100
                        add_Control(i).check_box(control_add_num).Checked = GetINIValue(tab_selection, Check_Keyname(control_add_num - 1), ini_dir)
                        add_Control(i).check_box(control_add_num).Text = check_text(control_add_num)
                        '체크박스 크기 106,22
                        '첫 체크박스 위치 18,32 두번쨰 18,55 
                        Selected_tab.Controls.Add(add_Control(i).label(control_add_num))
                        add_Control(i).label(control_add_num).Enabled = True
                        add_Control(i).label(control_add_num).Top = 25 * control_add_num
                        add_Control(i).label(control_add_num).Left = 140
                        add_Control(i).label(control_add_num).Height = 20
                        add_Control(i).label(control_add_num).Width = 30
                        add_Control(i).label(control_add_num).Text = GetINIValue(tab_selection, address_Keyname(control_add_num - 1), ini_dir)

                        '라벨 크기 32,18
                        '라벨 위치 144,32
                    Next control_add_num

                    Selected_tab.Controls.Add(add_Control(i).label(9))  '페이지 줄수 변수
                    add_Control(i).label(9).Enabled = True
                    add_Control(i).label(9).Top = 225
                    add_Control(i).label(9).Left = 140
                    add_Control(i).label(9).Height = 20
                    add_Control(i).label(9).Width = 150
                    add_Control(i).label(9).Text = GetINIValue(tab_selection, custom_Keyname(1), ini_dir)

                    Selected_tab.Controls.Add(add_Control(i).label(10))     '입력방향 변수
                    add_Control(i).label(10).Enabled = True
                    add_Control(i).label(10).Top = 250
                    add_Control(i).label(10).Left = 140
                    add_Control(i).label(10).Height = 20
                    add_Control(i).label(10).Width = 150
                    add_Control(i).label(10).Text = GetINIValue(tab_selection, custom_Keyname(2), ini_dir)

                    Selected_tab.Controls.Add(add_Control(i).label(11))     '사용 성적서 경로 변수
                    add_Control(i).label(11).Enabled = True
                    add_Control(i).label(11).Top = 300
                    add_Control(i).label(11).Left = 20
                    add_Control(i).label(11).Height = 20
                    add_Control(i).label(11).Width = 150
                    add_Control(i).label(11).AutoSize = True
                    add_Control(i).label(11).MaximumSize = New Size(180, 50)
                    add_Control(i).label(11).ContextMenuStrip() = ContextMenuStrip3
                    add_Control(i).label(11).Text = GetINIValue(custom_section, custom_Keyname(0), ini_dir)

                    Selected_tab.Controls.Add(add_Control(i).label(12))          '셀주소
                    add_Control(i).label(12).Enabled = True
                    add_Control(i).label(12).Top = 5
                    add_Control(i).label(12).Left = 140
                    add_Control(i).label(12).Height = 20
                    add_Control(i).label(12).Width = 150
                    add_Control(i).label(12).Text = "셀 주소"

                    Selected_tab.Controls.Add(add_Control(i).label(13))          '페이지줄수
                    add_Control(i).label(13).Enabled = True
                    add_Control(i).label(13).Top = 225
                    add_Control(i).label(13).Left = 20
                    add_Control(i).label(13).Height = 20
                    add_Control(i).label(13).Width = 150
                    add_Control(i).label(13).Text = "페이지 줄 수     :"

                    Selected_tab.Controls.Add(add_Control(i).label(14))          '입력방향  
                    add_Control(i).label(14).Enabled = True
                    add_Control(i).label(14).Top = 250
                    add_Control(i).label(14).Left = 20
                    add_Control(i).label(14).Height = 20
                    add_Control(i).label(14).Width = 150
                    add_Control(i).label(14).Text = "입력 방향          :"

                    Selected_tab.Controls.Add(add_Control(i).label(15))          '사용 성적서 경로
                    add_Control(i).label(15).Enabled = True
                    add_Control(i).label(15).Top = 275
                    add_Control(i).label(15).Left = 20
                    add_Control(i).label(15).Height = 20
                    add_Control(i).label(15).Width = 150

                    add_Control(i).label(15).Text = "사용 성적서 경로 :"

                    '======================================================================컨트롤 위치 고정용 


                Next i

        End Select

        Me.Refresh()
    End Sub



    ' 리스트박스에서 엔터 눌렀을때 키코드에 대응해서 이벤트 발생 ↓↓↓↓
    Private Sub KEY_down_EVENT(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListBox1.KeyDown

        If e.KeyCode = 13 Then      '13 = Enter
            Button1_Click(sender, New System.EventArgs())
        End If

        'If e.KeyCode = 116 Then      '116 = F5
        ' ListBox1.Refresh()
        ' End If

        If e.KeyCode = 46 Then          '   46 = delete
            Button5_Click(sender, New System.EventArgs())       '삭제 버튼
        End If

        'If e.KeyCode = 27 Then      '27 = ESC
        'ListBox1.Refresh()
        'End If
    End Sub

    Sub Get_list()

        Dim counting As Integer
        Dim Tempcounting As Integer
        Dim ListTemp As String
        counting = 0
        ListTemp = Dir(MRM_root_dir & "\MRM\data\resources\ini\*.ini")
        ReDim Matching_List(500)
        Matching_List(counting) = Dir(MRM_root_dir & "\MRM\data\resources\ini\*.ini")
        Do
            Tempcounting = InStr(ListTemp, ".ini")
            If Tempcounting = 0 Then Exit Do
            Matching_List(counting) = ListTemp.Substring(0, Tempcounting - 1)
            ListBox1.Items.Add(Matching_List(counting))
            counting = counting + 1
            ListTemp = Dir()
        Loop Until ListTemp = ""
        counting = counting - 1
        ReDim Preserve Matching_List(counting)

        List_check = 1

    End Sub
    Sub Measure_data_basic_form(form_txt As String)

        Dim input_string As String

        Dim cRow As Integer
        Dim i As Integer
        Dim dump_num As Integer
        Dim error_arry() As String
        Dim error_txt As String

        Dim Data_sheet As String
        Dim form_txt_1 As String
        Dim form_txt_2 As String
        Dim sheets_switch As Integer
        Dim sheets_del As Integer

        sheets_switch = 0

        Select Case form_txt

            Case "기본폼1"        '말머리 전부 포함               1사용
                form_txt_1 = "기본폼1"
                form_txt_2 = form_txt_1
                sheets_switch = 0

            Case "기본폼2"        '말머리 전부 없음               2 사용
                form_txt_1 = "기본폼2"
                form_txt_2 = form_txt_1
                sheets_switch = 1

            Case "기본폼3"        '그림 전부 삽입                 3 사용
                form_txt_1 = "기본폼3"
                form_txt_2 = form_txt_1
                sheets_switch = 2

            Case "기본폼4"        '첫 페이지 말머리 있음 2 페이지 부터 말머리 없음            1,2사용
                form_txt_1 = "기본폼1"
                form_txt_2 = "기본폼2"
                sheets_switch = 0
                sheets_del = 1

            Case "기본폼5"         '첫페이지 그림 삽입, 2 페이지부터 그림, 말머리 없음         2,3 사용
                form_txt_1 = "기본폼3"
                form_txt_2 = "기본폼2"
                sheets_switch = 2
                sheets_del = 1

            Case Else

                form_txt_1 = "기본폼1"
                form_txt_2 = "기본폼1"

        End Select

        Data_sheet = "Data Sheet"

        error_arry = Split("TP (3D),원형,동심도,진직도,PA,VT,VG,런아웃,대칭,평면도,TP (2D)", ",")

        '=========================================================================

        XL.Workbooks.open(Open_Dir)       '성적서 오픈
        XL.DisplayAlerts = False


        XL.Sheets.add(before:=XL.Sheets("기본폼1")) 'csv파일 가져올 워크시트 추가
        XL.activesheet.name = Data_sheet

        FileOpen(3, CSV_Dir, OpenMode.Input)            '배열 측정 정리용 구문

        'XL.visible = True

        Do Until EOF(3)
            cRow = cRow + 1
            input_string = LineInput(3)
            Dim input_arry() As String = Split(input_string, ",", -1)
            If input_arry(2).IndexOf("]") <> -1 Then
                Dim input_arry2() As String = Split(input_string, ",", -1)
                On Error Resume Next
                input_arry(0) = input_arry2(0)
                input_arry(1) = Mid(input_arry2(1), 2) & "," & Replace(input_arry2(2), """", " ")
                input_arry(2) = input_arry2(3)
                input_arry(3) = input_arry2(4)
                input_arry(4) = input_arry2(5)
                input_arry(5) = input_arry2(6)
                input_arry(6) = input_arry2(7)
                input_arry(7) = input_arry2(8)
                input_arry(8) = input_arry2(9)
                input_arry(9) = input_arry2(10)


            End If
            Select Case dump_num
                Case 0      'dump
                    For i = 0 To UBound(input_arry)
                        XL.sheets(Data_sheet).cells(cRow, i + 1).value = input_arry(i)
                        XL.sheets(Data_sheet).cells(cRow, i + 1).value = ""
                    Next

                    dump_num = 1
                    cRow = 0

                Case 1
                    For i = 0 To UBound(input_arry)
                        XL.sheets(Data_sheet).cells(cRow, i + 1).value = input_arry(i)
                    Next
            End Select

        Loop
        On Error GoTo 0         ' on error turn off
        FileClose(3)
        'XL.visible = True
        '================================================================기본 유저 정보 입력

        input_user_info()
        Sheet_del()     '시트  삭제

        XL.Sheets(form_txt_1).Copy(After:=XL.Sheets(Data_sheet))        '데이터 입력용 워크시트 복사
        XL.activesheet.name = "DATA-1"                                         '워크시트 선택

        Cell_Address = 1                    '워크시트1 셀 위치
        Sheet_Count = 1                     '데이터 입력 워크시트 카운트
        Cell_Count2 = XL.Sheets(Data_sheet).Rows.Count       '워크시트 행 위치 검색
        Cell_Count = XL.Sheets(Data_sheet).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
        Select Case form_txt                    '페이지수 계산

            Case "기본폼1"
                Quetient = Cell_Count / 28              '페이지수 계산  폼1 : 28, 폼2 : 32, 폼3 : 15
            Case "기본폼2"        '말머리 전부 없음               2 사용
                Quetient = Cell_Count / 32
            Case "기본폼3"        '그림 전부 삽입                 3 사용
                Quetient = Cell_Count / 15
            Case "기본폼4"        '첫 페이지 말머리 있음 2 페이지 부터 말머리 없음            1,2사용
                Quetient = ((Cell_Count - 28) / 32) + 1
            Case "기본폼5"         '첫페이지 그림 삽입, 2 페이지부터 그림, 말머리 없음         2,3 사용
                Quetient = ((Cell_Count - 15) / 32) + 1
            Case Else

        End Select

        TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

        프로그레스바.ProgressBar1.Maximum = Cell_Count

        Dim count As Integer

        If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

        Do      '데이터 입력
            Select Case sheets_switch
                Case 0      '기본폼1
                    Line_Count = 9
                Case 1      '기본폼2
                    Line_Count = 5
                Case 2      '기본폼 3
                    Line_Count = 22
            End Select

            Do Until Line_Count > 36
                XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = XL.Sheets(Data_sheet).Range("B" & Cell_Address).value & "(" & XL.Sheets(Data_sheet).Range("C" & Cell_Address).value & ")"   '라벨명
                XL.Sheets("DATA-" & Sheet_Count).Range("B" & Line_Count).value = XL.Sheets(Data_sheet).Range("D" & Cell_Address).value     '구성요소
                XL.Sheets("DATA-" & Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(Data_sheet).Range("E" & Cell_Address).value     '측정값
                XL.Sheets("DATA-" & Sheet_Count).Range("D" & Line_Count).value = XL.Sheets(Data_sheet).Range("F" & Cell_Address).value     '설계치
                error_txt = XL.Sheets(Data_sheet).Range("D" & Cell_Address).value

                If error_arry.Contains(error_txt) Then                                                                                      '기하공차 오차값을 측정값에 입력
                    XL.Sheets("DATA-" & Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(Data_sheet).Range("G" & Cell_Address).value
                Else
                    XL.Sheets("DATA-" & Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(Data_sheet).Range("G" & Cell_Address).value
                End If

                ' XL.Sheets(Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(1).Range("G" & Cell_Address).value     '오차   TP (3D), 동심도, 진직도, PA, VT, VG,런아웃, 대칭
                XL.Sheets("DATA-" & Sheet_Count).Range("F" & Line_Count).value = XL.Sheets(Data_sheet).Range("H" & Cell_Address).value     '상한
                XL.Sheets("DATA-" & Sheet_Count).Range("G" & Line_Count).value = XL.Sheets(Data_sheet).Range("I" & Cell_Address).value     '하한
                XL.Sheets("DATA-" & Sheet_Count).Range("H" & Line_Count).value = XL.Sheets(Data_sheet).Range("J" & Cell_Address).value     '판정 // 통과/실패

                '====================================================================================================================================
                If 프로그레스바.ProgressBar1.Value = 프로그레스바.ProgressBar1.Maximum Then
                Else
                    프로그레스바.ProgressBar1.Value += 1
                End If
                '====================================================================================================================================

                If XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = "()" Then
                    XL.Sheets("DATA-" & Sheet_Count).Range("A" & Line_Count).value = ""
                    Exit Do
                End If

                Line_Count = Line_Count + 1
                Cell_Address = Cell_Address + 1
                count += 1

            Loop

            Select Case sheets_switch
                Case 0, 2           '기본폼1, 기본폼3
                    XL.sheets("DATA-" & Sheet_Count).range("I3").value = Sheet_Count & "/" & TotalSheet  '페이지 번호 입력

                Case 1      '기본폼2
                    XL.sheets("DATA-" & Sheet_Count).range("I1").value = Sheet_Count & "/" & TotalSheet  '페이지 번호 입력
            End Select

            If Sheet_Count = TotalSheet Then        '같은 페이지일때 점프로 나가기
                Exit Do
            End If

            If Sheet_Count = 0 Then
                XL.Sheets(form_txt_1).Copy(After:=XL.Sheets("DATA-" & Sheet_Count))
                XL.activesheet.name = "DATA-" & (Sheet_Count + 1)
            Else
                XL.Sheets(form_txt_2).Copy(After:=XL.Sheets("DATA-" & Sheet_Count))
                XL.activesheet.name = "DATA-" & (Sheet_Count + 1)
                Select Case form_txt
                    Case "기본폼4"
                        sheets_switch = 1
                    Case "기본폼5"
                        sheets_switch = 1
                End Select

            End If

            Sheet_Count = Sheet_Count + 1

        Loop Until Sheet_Count > (TotalSheet)

        XL.Sheets(form_txt_1).delete
        If sheets_del = 1 Then
            XL.Sheets(form_txt_2).delete
        End If
        XL.Sheets(Data_sheet).delete

    End Sub
    Sub Sheet_del()

        '=========================================================================    기본 폼 선택에 따라 시트 삭제
        Select Case Label20.Text

            Case "기본폼1"        '말머리 전부 포함               1사용
                XL.sheets("기본폼2").delete
                XL.sheets("기본폼3").delete

            Case "기본폼2"        '말머리 전부 없음               2 사용
                XL.sheets("기본폼1").delete
                XL.sheets("기본폼3").delete

            Case "기본폼3"        '그림 전부 삽입                 3 사용
                XL.sheets("기본폼1").delete
                XL.sheets("기본폼2").delete

            Case "기본폼4"        '첫 페이지 말머리 있음 2 페이지 부터 말머리 없음            1,2사용
                XL.sheets("기본폼3").delete

            Case "기본폼5"         '첫페이지 그림 삽입, 2 페이지부터 그림, 말머리 없음         2,3 사용
                XL.sheets("기본폼1").delete

        End Select
        '=========================================================================

    End Sub
    Function Restore_str(str As String) As String
        Dim origin_str As String
        origin_str = str
        Return origin_str
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim sfd_For_manual As New SaveFileDialog()
        Dim manual As Byte() = My.Resources.MRM_사용_설명서
        With sfd_For_manual
            .InitialDirectory = "C:\"
            .Filter = "PDF|.PDF"
            .FilterIndex = 1
            .Title = "Save Manual"
            .FileName = "MRM 사용 설명서"
            .RestoreDirectory = True
            .CheckFileExists = False
            .CheckPathExists = True
        End With

        If sfd_For_manual.ShowDialog() = Windows.Forms.DialogResult.OK Then
            My.Computer.FileSystem.WriteAllBytes(sfd_For_manual.FileName, manual, False)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        LinkLabel1.LinkVisited = True
        System.Diagnostics.Process.Start("http://www.mitutoyokorea.com/")
    End Sub

    Function MID_CHECK_EH(ByVal serial_num As String) As String 'Error Handling

        On Error GoTo Out

        ' If serial_num = MID_CHECK.Class1.MID_Check() Then
        'Return MID_CHECK.Class1.MID_Check

        If MID_CHECK.Class1.MID_Check(serial_num) = serial_num Then
            Return MID_CHECK.Class1.MID_Check(serial_num)
        Else
            'SplashScreen1.Close()
            MsgBox("        >>>>>   파일 복사 감지   <<<<<    " & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다." & Environment.NewLine & "ERROR CODE : M-003", 48, "MRM Activation")   'M-003 : MID 불러오기 실패
            Me.Close()
            End
        End If

        Exit Function
Out:
        MsgBox("        >>>>>   파일 복사 감지   <<<<<    " & Environment.NewLine & "프로그램의 문제해결, 문의사항은 Mitutoyo Korea에 문의 부탁드립니다." & Environment.NewLine & "ERROR CODE : M-002", 48, "MRM Activation")            'M-002 : MID 불러오기 오류 발생
        Me.Close()
        End
    End Function

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim Change_name As String
        Dim Origin_name As String
        Dim Select_name As String
        Dim Folder_Name As String
        Dim Folder_Name2 As String
        Dim Select_ini_dir As String
        Dim Form_Ans As Integer
        Dim destination_dir As String
        Dim Change_PRG_name As String
        Dim Current_PRG_DIR As String


        If ListBox1.SelectedItem IsNot Nothing Then
        Else
            MsgBox("성적서 매칭 리스트를 선택해주세요",, "매칭 리스트 선택 오류")
            Exit Sub
        End If

        Select_name = ListBox1.SelectedItem.ToString()
        Select_ini_dir = MRM_root_dir & "\MRM\Data\Resources\ini\" & Select_name & ".ini"

        Form8.ShowDialog()
        Form_Ans = Form8.Ans
        Change_name = Form8.Change_Name
        Origin_name = System.Diagnostics.Process.GetCurrentProcess().ProcessName        '실행중인 프로세스 이름 읽어오기, 파일명 읽기
        Current_PRG_DIR = System.Reflection.Assembly.GetExecutingAssembly.Location      '실행된 파일 위치 읽어오기


        Folder_Name2 = MRM_root_dir & "\MRM\전용프로그램"
        Dim Folder_Exists3 As New System.IO.DirectoryInfo(Folder_Name2)
        If Folder_Exists3.Exists = False Then MkDir(Folder_Name2)

        Select Case Form_Ans

            Case 0

                Folder_Name = MRM_root_dir & "\MRM\Data\Resources\ini\" & Change_name

                destination_dir = Folder_Name & "\" & Change_name & ".ini"
                Change_PRG_name = MRM_root_dir & "\MRM\전용프로그램\" & Change_name & ".exe"


                Dim Folder_Exists As New System.IO.DirectoryInfo(Folder_Name)
                If Folder_Exists.Exists = False Then
                    MkDir(Folder_Name)
                Else
                    MsgBox("동일한 이름의 전용 프로그램이 존재합니다." & Environment.NewLine & "교체 하시려면 전용 버튼을 눌러 기존 전용 프로그램을 삭제 후 재생성해주세요.", vbCritical, "중복 이름 프로그램 존재")
                    Exit Sub
                End If

                Dim Folder_Exists2 As New System.IO.DirectoryInfo(Change_PRG_name)
                If Folder_Exists2.Exists = False Then

                Else
                    MsgBox("동일한 이름의 전용 프로그램이 존재합니다." & Environment.NewLine & "교체 하시려면 전용 버튼을 눌러 기존 전용 프로그램을 삭제 후 재생성해주세요.", vbCritical, "중복 이름 프로그램 존재")
                    Exit Sub
                End If
                'ini파일 복사후 이름변경

                FileCopy(Select_ini_dir, destination_dir)

                '프로그램 복사 후 이름 변경 

                FileCopy(Current_PRG_DIR, Change_PRG_name)

            Case 1
                MsgBox("전용 프로그램 생성 취소", , "전용프로그램 생성 취소")

        End Select
    End Sub

    Sub add_str()
        Dim add_str_section As String
        Dim add_str_keyname(10) As String
        Dim add_ans_value As String
        Dim dialog_value As Integer
        Dim i As Integer

        add_str_section = "add_str_"
        add_str_keyname(0) = "Description"
        add_str_keyname(1) = "value"
        add_str_keyname(2) = "loction"
        add_str_keyname(3) = "combo"
        add_str_keyname(4) = "use_check"
        add_str_keyname(5) = "apply_tab"
        add_str_keyname(6) = "input_type"

        For i = 1 To 3
            add_str_value(0) = GetINIValue(add_str_section & i, add_str_keyname(0), Restore_str(ini_dir))
            add_str_value(1) = GetINIValue(add_str_section & i, add_str_keyname(1), Restore_str(ini_dir))
            add_str_value(2) = GetINIValue(add_str_section & i, add_str_keyname(2), Restore_str(ini_dir))
            add_str_value(3) = GetINIValue(add_str_section & i, add_str_keyname(3), Restore_str(ini_dir))
            add_str_value(4) = GetINIValue(add_str_section & i, add_str_keyname(4), Restore_str(ini_dir))
            add_str_value(5) = GetINIValue(add_str_section & i, add_str_keyname(5), Restore_str(ini_dir))
            add_str_value(6) = GetINIValue(add_str_section & i, add_str_keyname(6), Restore_str(ini_dir))

            Select Case add_str_value(3)

                Case "자동기입"
                    If add_str_value(4) = True Then

                        Select Case add_str_value(6)
                            Case "텍스트"
                                XL.sheets(add_str_value(5)).select
                                XL.Sheets(add_str_value(5)).Range(add_str_value(2)).value = add_str_value(1)

                            Case "날짜"
                                XL.sheets(add_str_value(5)).select
                                XL.Sheets(add_str_value(5)).Range(add_str_value(2)).value = Format(Now(), "yyyy/MM/dd")
                            Case "시간"
                                XL.sheets(add_str_value(5)).select
                                XL.Sheets(add_str_value(5)).Range(add_str_value(2)).value = Format(Now(), "hh:mm:ss")
                            Case "날짜 + 시간"
                                XL.sheets(add_str_value(5)).select
                                XL.Sheets(add_str_value(5)).Range(add_str_value(2)).value = Now()
                        End Select

                    End If
                Case "매번 생성시"

                    If add_str_value(4) = True Then
                        dialog_value = Add_str_dialog.ShowDialog()
                        If dialog_value = 1 Then                ' 1>> OK 버튼
                            add_ans_value = Add_str_dialog.ans_textbox
                            XL.sheets(add_str_value(5)).select
                            XL.Sheets(add_str_value(5)).Range(add_str_value(2)).value = add_ans_value
                        End If
                    End If
            End Select

        Next i

    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Dim counting As Integer
        Dim Tempcounting As Integer
        Dim ListTemp As String

        ListBox1.Items.Clear()

        counting = 0
        ListTemp = Dir(MRM_root_dir & "\MRM\data\resources\ini\*.ini")
        ReDim Matching_List(500)
        Matching_List(counting) = Dir(MRM_root_dir & "\MRM\data\resources\ini\*.ini")
        Do
            Tempcounting = InStr(ListTemp, ".ini")
            If Tempcounting = 0 Then Exit Do
            Matching_List(counting) = ListTemp.Substring(0, Tempcounting - 1)
            ListBox1.Items.Add(Matching_List(counting))
            counting = counting + 1
            ListTemp = Dir()
        Loop Until ListTemp = ""
        counting = counting - 1
        ReDim Preserve Matching_List(counting)

        ListBox1.Refresh()

        List_check = 1                   '일반 프로그램 리스트

        Button5.BackColor = Color.Silver
        Button8.BackColor = Color.White

        Button2.Enabled = True
        PictureBox1.Enabled = True
        Button1.Enabled = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim counting As Integer
        Dim Tempcounting As Integer
        Dim ListTemp As String
        counting = 0

        ListBox1.Items.Clear()

        ListTemp = Dir(MRM_root_dir & "\MRM\전용프로그램\*.exe")
        ReDim Matching_List(500)
        Matching_List(counting) = Dir(MRM_root_dir & "\MRM\전용프로그램\*.exe")
        Do
            Tempcounting = InStr(ListTemp, ".exe")
            If Tempcounting = 0 Then Exit Do
            Matching_List(counting) = ListTemp.Substring(0, Tempcounting - 1)
            ListBox1.Items.Add(Matching_List(counting))
            counting = counting + 1
            ListTemp = Dir()
        Loop Until ListTemp = ""
        counting = counting - 1
        ReDim Preserve Matching_List(counting)

        ListBox1.Refresh()

        List_check = 2                  '전용 프로그램 리스트
        Button5.BackColor = Color.White
        Button8.BackColor = Color.Silver

        Button2.Enabled = False
        PictureBox1.Enabled = False
        Button1.Enabled = False
    End Sub
End Class