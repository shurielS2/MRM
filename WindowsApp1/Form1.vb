'=============================================================================================================
'2020-07-02 ver 1.00.1  기본성적서 이외에 위치지정 성적서 폼 라디오 버튼 및 속성 추가
'2020-07-14 ver 1.00.2  form3 -리스트 수정시 form1 리스트 재로드 수정
'2020-07-14 ver 1.00.2  위치지정 성적서의 베이스가 될 성적서 파일 지정 추가
'2020-07-15 ver 1.00.3  성적서 save folder 위치 지정시 지정 폴더가 없으면 생성 기능 추가
'2020-07-15 ver 1.00.3  기본 성적서폼 선택시 위치지정 속성 공백 지정해서 저장하는 기능 추가
'2020-07-15 ver 1.01.0  위치지정 성적서 생성 루틴 추가 
'2020-07-15 ver 1.01.1  줄 수 지정 숫자만 입력하는 이벤트 추가
'2020-07-15 ver 1.01.2  Form1 셀 주소 지정 영문과 숫자 분리 (ad_NUM, ad_str 함수)
'2020-07-15 ver 1.01.3  수정 로드시 체크박스 상태에 따라서 텍스트 박스 비활성화
'2020-07-15 ver 1.01.4  자동저장 안할시에 임시파일로 백업본 저장
'2020-07-16 ver 1.01.5  기본 성적서 없을때 자동 생성시 오류 workbook >workbooks 수정
'2020-07-16 ver 1.01.6  위치지정 loop문 line_count 스텍 누락 추가
'2020-07-16 ver 1.01.7  성적서 생성, 수정 시 공백 확인 누락된것 추가 (csv 파일, 저장 유형)
'2020-08-13 Ver 1.01.8  프로그래스바 추가 
'2020-08-13 Ver 1.01.9  성적서 자동저장 체크박스 추가. -기본으로는 자동저장
'2020-08-13 Ver 1.01.9  A9:I36 범위 shrinktofit = true 설정
'2020-08-18 Ver 1.02.0  파일 경로 오류 수정 (False <> false 문제)
'2020-08-18 Ver 1.02.0  상위 폴더 경로 (MRM) 추가
'2020-08-19 Ver 1.02.1  복사방지 추가 & 액티브 키(프로그램) 제작
'2020-08-20 Ver 1.03.0  3차원 측정기용 ASC 파일 매칭 추가
'2020-08-20 Ver 1.03.1  fileopen - input 구현 방식으로 변경
'2020-08-20 Ver 1.03.2  중복실행 방지 구문 추가
'=============================================================================================================
Public Class Form1


    Declare Function GPPS Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WPPS Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

    Dim Logic_value As Integer

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







    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load            '로드시 처음 행동

        Me.Location = New Point(100, 100)

        '==================================================== 머신 활성화 체크
        If Lib_Serial_Check() = False Then
            Me.Close()
        End If

        '==================================================== 중복실행 방지

        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            MsgBox("프로그램이 이미 실행중입니다!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, Me.Text)
            End
        End If

        '====================================================리스트 파일 로드
        Dim counting As Integer
        Dim Tempcounting As Integer
        Dim ListTemp As String
        counting = 0
        ListTemp = Dir(CurDir() & "\MRM\data\resources\ini\*.ini")
        ReDim Matching_List(500)
        Matching_List(counting) = Dir(CurDir() & "\MRM\data\resources\ini\*.ini")
        Do
            Tempcounting = InStr(ListTemp, ".ini")
            If Tempcounting = 0 Then Exit Sub
            Matching_List(counting) = ListTemp.Substring(0, Tempcounting - 1)
            ListBox1.Items.Add(Matching_List(counting))
            counting = counting + 1
            ListTemp = Dir()
        Loop Until ListTemp = ""
        counting = counting - 1
        ReDim Preserve Matching_List(counting)
        '====================================================
        '개요 select case 문으로 레지스트리 값 읽어 오는 판단에서 분기점
        '읽어 올때 레지스트리가 존재하면 정상 구동, 없으면 레지스트리 등록 절차로 이동
        '레지스트리 읽어 올 때 장비 id 랑 일치 여부 확인 복사 방지구문 구성
        '레지스트리 등록시 장비 id 등록 
        ' 결과예상 - 첫플레이시에만 장비아이디 등록 
        ' 이후 플레이시에는 레지스트리가 존재하기때문에 등록 절차없이 플레이
        '예상 문제점 컴퓨터를 옮겼는데 예전 컴퓨터에서 여전히 플레이 가능한 문제.
        'MsgBox(Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber))


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click           '성적서 생성 버튼
        If ListBox1.SelectedItem IsNot Nothing Then
        Else
            MsgBox("성적서 매칭 리스트를 선택해주세요",, "매칭 리스트 선택 오류")
            Exit Sub
        End If

        ini_Name = ListBox1.SelectedItem.ToString()
        ini_dir = CurDir() & "\MRM\Data\Resources\ini\" & ini_Name & ".ini"



        '================================


        ReDim strData(10)

        Open_Dir = CurDir() & "\MRM\Data\Resources\Result_Source.xlsx"   '원본 엑셀

        Call orign_form()

        CSV_Dir = Label16.Text
        '=========================================================================
        '프로그래스바
        Form6.Location = New Point(500, 500)
        Form6.Show()
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
                Call extension_type_3()     'csv    
            Case 2
                Call extension_type_2()     'asc
        End Select


        Select Case strDate
            Case "True"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", ini_dir) & "\" & GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "_" & DateString & "_" & Format(TimeOfDay, "HH-mm-ss") & GetINIValue("Matching_info", "Save_Type", ini_dir)

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", ini_dir) & "\" & GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "_" & DateString & GetINIValue("Matching_info", "Save_Type", ini_dir)

                End Select
            Case "False"
                Select Case strTime
                    Case "True"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", ini_dir) & "\" & GetINIValue("Matching_info", "Save_File_Name", ini_dir) & "_" & Format(TimeOfDay, "HH-mm-ss''") & GetINIValue("Matching_info", "Save_Type", ini_dir)

                    Case "False"
                        Save_Dir = GetINIValue("Matching_info", "Save_File_Path", ini_dir) & "\" & GetINIValue("Matching_info", "Save_File_Name", ini_dir) & GetINIValue("Matching_info", "Save_Type", ini_dir)
                End Select

        End Select


        WPPS("Matching_Info", "Last_Paly_date", Now(), ini_dir)         '마지막 측정 시간 저장
        Form6.Close()
        Select Case Label22.Text
            Case "True"

                Select Case Label19.Text
                    Case ".xlsx"
                        With XL
                            .DisplayAlerts = False
                            .workbooks(1).SaveAS(filename:=Save_Dir)          '다른이름으로 저장 위치 지정
                            .Workbooks(1).close
                            .quit
                        End With

                    Case ".PDF"
                        With XL
                            .DisplayAlerts = False

                            .workbooks(1).printout(copies:=1, Collate:=True, PrToFilename:=Label18.Text, ignorePrintAreas:=False)
                            .Workbooks(1).close
                            .quit
                        End With
                End Select

                XL = Nothing
                Me.Close()
            Case "False"
                MsgBox("자동저장을 선택하지 않으셨습니다. 성적서가 화면에 켜집니다. 성적서를 따로 저장해주세요.",, "성적서 자동 생성 취소")
                XL.DisplayAlerts = False
                tempsave_dir = CurDir() & "\MRM\Result\temp"
                Dim fFindFolder As New System.IO.DirectoryInfo(tempsave_dir)           '폴더 존재 유무 확인
                If fFindFolder.Exists = False Then
                    MkDir(CurDir() & "\MRM\Result\temp")
                End If
                XL.workbooks(1).SaveAS(tempsave_dir & "\temp.xlsx")
                XL.visible = True
                Me.Close()
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
        ini_dir = CurDir() & "\MRM\Data\Resources\ini\" & ini_Name & ".ini"

        If ListBox1.SelectedItem = Nothing Then

        Else
            If Logic_value <> 2 Then
                strDate = GetINIValue("Matching_info", "Check_Date", ini_dir)
                strTime = GetINIValue("Matching_info", "Check_Time", ini_dir)
                Label16.Text = GetINIValue("Matching_info", "CSV_file_Path", ini_dir)
                Label17.Text = GetINIValue("Matching_info", "Save_File_Path", ini_dir) '& "\" & GetINIValue("Matching_info", "Save_File_Name", CurDir() & "\Data\Resources\Ini\" & ListBox1.SelectedItem.ToString & ".ini") & GetINIValue("Matching_info", "Save_Type", CurDir() & "\Data\Resources\Ini\" & ListBox1.SelectedItem.ToString & ".ini")
                Label22.Text = GetINIValue("Matching_info", "auto_save", ini_dir)
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
                Label20.Text = GetINIValue("Matching_info", "Result_Form", ini_dir)
                iniPath = GetINIValue("Matching_info", "ini_dir", ini_dir)
                logo_path = GetINIValue("Matching_info", "logo_path", ini_dir)


            End If
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ListBox1.SelectedItem IsNot Nothing Then
            ListBox1.Items.Remove(ListBox1.SelectedItem)
            Kill(iniPath)
            Logic_value = 1
        Else
            MsgBox("삭제할 목록을 선택해주세요")
            Exit Sub

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedItem IsNot Nothing Then
            Form3.ShowDialog()


        Else
            MsgBox("수정할 목록을 선택해주세요")
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

        Lib_Serial_Check = True
        Dim reg_path As String
        Dim registry_value As Long

        reg_path = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\MRM"
        registry_value = My.Computer.Registry.GetValue(reg_path, "Active_Mechine_ID", Nothing)
        Select Case Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber)

            Case registry_value  '내 노트북 

            Case Else
                Lib_Serial_Check = False
                MsgBox("측정 프로그램 활성화를 위해서 아래 하드웨어 ID 를 기록하여" & Chr(13) & "Mitutoyo Korea 로 문의 바랍니다." &
                Chr(13) & Chr(13) & "HardWare ID : " & Math.Abs(CreateObject("Scripting.FileSystemObject").GetDrive("C:").SerialNumber), 16, "복사 방지 오류")
                '    MkDir("C:\")

        End Select

    End Function
    Sub orign_form()
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
        Dim fFindFile As New System.IO.FileInfo(Open_Dir)
        If fFindFile.Exists = False Then

            '=========================================================================
            cXL = CreateObject("Excel.application")
            'XL.Visible = True
            cXL.Workbooks.add

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
                .FormulaR1C1 = "검사성적서   "
            End With
            With cXL.Range("G1")
                .merge
                .Font.Name = "맑은 고딕"
                .Font.Size = 11
                .Font.Bold = True
                .Value = "기안"
            End With
            With cXL.Range("H1")
                .merge
                .Font.Name = "맑은 고딕"
                .Font.Size = 11
                .Font.Bold = True
                .Value = "검토"
            End With
            With cXL.Range("I1")
                .merge
                .Font.Name = "맑은 고딕"
                .Font.Size = 11
                .Font.Bold = True
                .Value = "승인"
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
            Dim picture_dir As Object
            My.Resources._1920px_Mitutoyo_company_logo.Save(CurDir() & "\MRM\data\Resources\default_logo.png")
            picture_dir = CurDir() & "\MRM\data\Resources\default_logo.png"
            cXL.sheets(1).pictures.insert(picture_dir).select

            With cXL.Selection.ShapeRange
                .Width = 180.0
                .Height = 45.0
                .IncrementTop(0.75)
                .IncrementTop(0.75)
                .IncrementLeft(0.75)
                .IncrementLeft(0.75)

            End With
            'XL.visible = True

            cXL.workbooks(1).SaveAS(CurDir() & "\MRM\data\Resources\Result_Source.xlsx")
            cXL.DisplayAlerts = False
            cXL.workbooks(1).close
            cXL.quit
            cXL = Nothing
            'Kill(picture_dir)
        End If

        '=========================================================================

    End Sub

    Sub extension_type_1()          'CSV 파일 읽어오기
        '=========================================================================
        '기본 혹은 위치지정 선택 
        Select Case Label20.Text
        '=========================================================================

            Case "기본"
                '=========================================================================
                XL = CreateObject("Excel.application")
                gXL = CreateObject("Excel.application")

                XL.Workbooks.open(Open_Dir)       '성적서 오픈
                gXL.Workbooks.open(CSV_Dir)         'csv파일 오픈

                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
                gXL.sheets(1).Range("A2:K1000").copy    'csv 값 복사
                XL.Sheets(1).Range("A1").PasteSpecial(-4163)    'csv값 붙여넣기 
                Clipboard.Clear()

                With gXL            'csv파일 종료
                    .DisplayAlerts = False
                    .workbooks(1).close()
                    .quit()
                End With

                gXL = Nothing
                ' XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / 28              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                Dim count As Integer

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 9
                    Do Until Line_Count > 36

                        XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명
                        XL.Sheets(Sheet_Count).Range("B" & Line_Count).value = XL.Sheets(1).Range("D" & Cell_Address).value     '구성요소
                        XL.Sheets(Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(1).Range("E" & Cell_Address).value     '측정값
                        XL.Sheets(Sheet_Count).Range("D" & Line_Count).value = XL.Sheets(1).Range("F" & Cell_Address).value     '설계치
                        XL.Sheets(Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(1).Range("G" & Cell_Address).value     '오차
                        XL.Sheets(Sheet_Count).Range("F" & Line_Count).value = XL.Sheets(1).Range("H" & Cell_Address).value     '상한
                        XL.Sheets(Sheet_Count).Range("G" & Line_Count).value = XL.Sheets(1).Range("I" & Cell_Address).value     '하한
                        XL.Sheets(Sheet_Count).Range("H" & Line_Count).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        If XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = "()" Then
                            XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = ""
                            Exit Do
                        End If
                        Line_Count = Line_Count + 1
                        Cell_Address = Cell_Address + 1
                        count += 1

                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================

                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False

             '=========================================================================
            Case "위치 지정"
                '=========================================================================
                Dim Label_ad(1) As String
                Dim measure_ad(1) As String
                Dim component_ad(1) As String
                Dim design_ad(1) As String
                Dim error_ad(1) As String
                Dim UP_Tol_ad(1) As String
                Dim Low_Tol_ad(1) As String
                Dim judge_ad(1) As String
                Dim Line_count_ad As Integer
                Dim Result_Form_dir As String
                Dim select_check_value() As String
                ReDim select_check_value(9)

                Label_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(0) = Ad_Str(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(0) = Ad_Str(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(0) = Ad_Str(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(0) = Ad_Str(GetINIValue("custom_match_info", "judge", ini_dir))
                Line_count_ad = GetINIValue("custom_match_info", "Line_count", ini_dir)
                Result_Form_dir = GetINIValue("custom_match_info", "Result_Form_dir", ini_dir)


                Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                select_check_value(0) = GetINIValue("check", "label", ini_dir)
                select_check_value(1) = GetINIValue("check", "Measure_value", ini_dir)
                select_check_value(2) = GetINIValue("check", "Design_value", ini_dir)
                select_check_value(3) = GetINIValue("check", "error", ini_dir)
                select_check_value(4) = GetINIValue("check", "UP_tol", ini_dir)
                select_check_value(5) = GetINIValue("check", "Low_tol", ini_dir)
                select_check_value(6) = GetINIValue("check", "judge", ini_dir)
                select_check_value(7) = GetINIValue("check", "component", ini_dir)



                XL = CreateObject("Excel.application")
                gXL = CreateObject("Excel.application")

                XL.Workbooks.open(Result_Form_dir)       '성적서 오픈
                gXL.Workbooks.open(CSV_Dir)         'csv파일 오픈

                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
                gXL.sheets(1).Range("A2:K1000").copy    'csv 값 복사
                XL.Sheets(1).Range("A1").PasteSpecial(-4163)    'csv값 붙여넣기 
                With gXL            'csv파일 종료
                    .DisplayAlerts = False
                    .workbooks(1).close
                    .quit
                End With
                Clipboard.Clear()
                gXL = Nothing
                'XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / Line_count_ad              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 1

                    Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                    component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                    measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                    design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                    error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                    UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                    Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                    judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                    Do Until Line_Count > Line_count_ad

                        If select_check_value(0) = True Then
                            XL.Sheets(Sheet_Count).Range(Label_ad(0) & Label_ad(1)).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명
                        End If

                        If select_check_value(7) = True Then
                            XL.Sheets(Sheet_Count).Range(component_ad(0) & component_ad(1)).value = XL.Sheets(1).Range("D" & Cell_Address).value     '구성요소
                        End If

                        If select_check_value(1) = True Then
                            XL.Sheets(Sheet_Count).Range(measure_ad(0) & measure_ad(1)).value = XL.Sheets(1).Range("E" & Cell_Address).value     '측정값
                        End If

                        If select_check_value(2) = True Then
                            XL.Sheets(Sheet_Count).Range(design_ad(0) & design_ad(1)).value = XL.Sheets(1).Range("F" & Cell_Address).value     '설계치
                        End If

                        If select_check_value(3) = True Then
                            XL.Sheets(Sheet_Count).Range(error_ad(0) & error_ad(1)).value = XL.Sheets(1).Range("G" & Cell_Address).value     '오차
                        End If

                        If select_check_value(4) = True Then
                            XL.Sheets(Sheet_Count).Range(UP_Tol_ad(0) & UP_Tol_ad(1)).value = XL.Sheets(1).Range("H" & Cell_Address).value     '상한
                        End If

                        If select_check_value(5) = True Then
                            XL.Sheets(Sheet_Count).Range(Low_Tol_ad(0) & Low_Tol_ad(1)).value = XL.Sheets(1).Range("I" & Cell_Address).value     '하한
                        End If

                        If select_check_value(6) = True Then
                            XL.Sheets(Sheet_Count).Range(judge_ad(0) & judge_ad(1)).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        End If

                        '====================================================================================================================================
                        '====================================================================================================================================
                        '라인 끝 빈공간 용
                        XL.Sheets(Sheet_Count).Range("BA1").value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명

                        If XL.Sheets(Sheet_Count).Range("BA1").value = "()" Then
                            XL.Sheets(Sheet_Count).Range("BA1").value = ""
                            Exit Do
                        End If
                        XL.Sheets(Sheet_Count).Range("BA1").delete
                        '====================================================================================================================================
                        '====================================================================================================================================
                        '셀주소 하나씩 내리기
                        Label_ad(1) = Label_ad(1) + 1
                        component_ad(1) = component_ad(1) + 1
                        measure_ad(1) = measure_ad(1) + 1
                        design_ad(1) = design_ad(1) + 1
                        error_ad(1) = error_ad(1) + 1
                        UP_Tol_ad(1) = UP_Tol_ad(1) + 1
                        Low_Tol_ad(1) = Low_Tol_ad(1) + 1
                        judge_ad(1) = judge_ad(1) + 1
                        '====================================================================================================================================
                        Cell_Address = Cell_Address + 1
                        Line_Count = Line_Count + 1
                        '====================================================================================================================================
                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================
                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False


                '=========================================================================
                '기본혹은 위치지정 끝
        End Select
        '=========================================================================
    End Sub

    Sub extension_type_2()          'ASC 파일 읽어오기

        Dim asc_URL As String
        Dim input_string As String

        Dim cRow As Integer
        Dim i As Integer
        asc_URL = "TEXT;" & Open_Dir

        Select Case Label20.Text
        '=========================================================================

            Case "기본"
                '=========================================================================
                XL = CreateObject("Excel.application")

                XL.Workbooks.open(Open_Dir)       '성적서 오픈
                ' XL.visible = True

                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가


                FileOpen(1, CSV_Dir, OpenMode.Input)


                Do Until EOF(1)
                    cRow = cRow + 1
                    input_string = LineInput(1)
                    Dim input_arry() As String = Split(input_string, ";")
                    For i = 0 To UBound(input_arry)
                        XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                    Next
                Loop


                ' XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / 28              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                Dim count As Integer

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 9
                    Do Until Line_Count > 36

                        XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("A" & Cell_Address).value & ")"   '라벨명
                        XL.Sheets(Sheet_Count).Range("B" & Line_Count).value = XL.Sheets(1).Range("C" & Cell_Address).value     '구성요소
                        XL.Sheets(Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(1).Range("G" & Cell_Address).value     '측정값
                        XL.Sheets(Sheet_Count).Range("D" & Line_Count).value = XL.Sheets(1).Range("D" & Cell_Address).value     '설계치
                        XL.Sheets(Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(1).Range("H" & Cell_Address).value     '오차
                        XL.Sheets(Sheet_Count).Range("F" & Line_Count).value = XL.Sheets(1).Range("E" & Cell_Address).value     '상한
                        XL.Sheets(Sheet_Count).Range("G" & Line_Count).value = XL.Sheets(1).Range("F" & Cell_Address).value     '하한
                        XL.Sheets(Sheet_Count).Range("H" & Line_Count).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        If XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = "()" Then
                            XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = ""
                            Exit Do
                        End If
                        Line_Count = Line_Count + 1
                        Cell_Address = Cell_Address + 1
                        count += 1

                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================

                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False

             '=========================================================================
            Case "위치 지정"
                '=========================================================================
                Dim Label_ad(1) As String
                Dim measure_ad(1) As String
                Dim component_ad(1) As String
                Dim design_ad(1) As String
                Dim error_ad(1) As String
                Dim UP_Tol_ad(1) As String
                Dim Low_Tol_ad(1) As String
                Dim judge_ad(1) As String
                Dim Line_count_ad As Integer
                Dim Result_Form_dir As String
                Dim select_check_value() As String
                ReDim select_check_value(9)

                Label_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(0) = Ad_Str(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(0) = Ad_Str(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(0) = Ad_Str(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(0) = Ad_Str(GetINIValue("custom_match_info", "judge", ini_dir))
                Line_count_ad = GetINIValue("custom_match_info", "Line_count", ini_dir)
                Result_Form_dir = GetINIValue("custom_match_info", "Result_Form_dir", ini_dir)


                Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                select_check_value(0) = GetINIValue("check", "label", ini_dir)
                select_check_value(1) = GetINIValue("check", "Measure_value", ini_dir)
                select_check_value(2) = GetINIValue("check", "Design_value", ini_dir)
                select_check_value(3) = GetINIValue("check", "error", ini_dir)
                select_check_value(4) = GetINIValue("check", "UP_tol", ini_dir)
                select_check_value(5) = GetINIValue("check", "Low_tol", ini_dir)
                select_check_value(6) = GetINIValue("check", "judge", ini_dir)
                select_check_value(7) = GetINIValue("check", "component", ini_dir)



                XL = CreateObject("Excel.application")
                gXL = CreateObject("Excel.application")

                XL.Workbooks.open(Result_Form_dir)       '성적서 오픈
                gXL.Workbooks.open(CSV_Dir)         'csv파일 오픈

                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
                FileOpen(1, CSV_Dir, OpenMode.Input)


                Do Until EOF(1)
                    cRow = cRow + 1
                    input_string = LineInput(1)
                    Dim input_arry() As String = Split(input_string, ";")
                    For i = 0 To UBound(input_arry)
                        XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                    Next
                Loop

                'XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / Line_count_ad              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 1

                    Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                    component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                    measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                    design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                    error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                    UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                    Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                    judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                    Do Until Line_Count > Line_count_ad

                        If select_check_value(0) = True Then
                            XL.Sheets(Sheet_Count).Range(Label_ad(0) & Label_ad(1)).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("A" & Cell_Address).value & ")"   '라벨명
                        End If

                        If select_check_value(7) = True Then
                            XL.Sheets(Sheet_Count).Range(component_ad(0) & component_ad(1)).value = XL.Sheets(1).Range("C" & Cell_Address).value     '구성요소
                        End If

                        If select_check_value(1) = True Then
                            XL.Sheets(Sheet_Count).Range(measure_ad(0) & measure_ad(1)).value = XL.Sheets(1).Range("G" & Cell_Address).value     '측정값
                        End If

                        If select_check_value(2) = True Then
                            XL.Sheets(Sheet_Count).Range(design_ad(0) & design_ad(1)).value = XL.Sheets(1).Range("D" & Cell_Address).value     '설계치
                        End If

                        If select_check_value(3) = True Then
                            XL.Sheets(Sheet_Count).Range(error_ad(0) & error_ad(1)).value = XL.Sheets(1).Range("H" & Cell_Address).value     '오차
                        End If

                        If select_check_value(4) = True Then
                            XL.Sheets(Sheet_Count).Range(UP_Tol_ad(0) & UP_Tol_ad(1)).value = XL.Sheets(1).Range("E" & Cell_Address).value     '상한
                        End If

                        If select_check_value(5) = True Then
                            XL.Sheets(Sheet_Count).Range(Low_Tol_ad(0) & Low_Tol_ad(1)).value = XL.Sheets(1).Range("F" & Cell_Address).value     '하한
                        End If

                        If select_check_value(6) = True Then
                            XL.Sheets(Sheet_Count).Range(judge_ad(0) & judge_ad(1)).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        End If

                        '====================================================================================================================================
                        '====================================================================================================================================
                        '라인 끝 빈공간 용
                        XL.Sheets(Sheet_Count).Range("BA1").value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("A" & Cell_Address).value & ")"   '라벨명

                        If XL.Sheets(Sheet_Count).Range("BA1").value = "()" Then
                            XL.Sheets(Sheet_Count).Range("BA1").value = ""
                            Exit Do
                        End If
                        XL.Sheets(Sheet_Count).Range("BA1").delete
                        '====================================================================================================================================
                        '====================================================================================================================================
                        '셀주소 하나씩 내리기
                        Label_ad(1) = Label_ad(1) + 1
                        component_ad(1) = component_ad(1) + 1
                        measure_ad(1) = measure_ad(1) + 1
                        design_ad(1) = design_ad(1) + 1
                        error_ad(1) = error_ad(1) + 1
                        UP_Tol_ad(1) = UP_Tol_ad(1) + 1
                        Low_Tol_ad(1) = Low_Tol_ad(1) + 1
                        judge_ad(1) = judge_ad(1) + 1
                        '====================================================================================================================================
                        Cell_Address = Cell_Address + 1
                        Line_Count = Line_Count + 1
                        '====================================================================================================================================
                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================
                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False


                '=========================================================================
                '기본혹은 위치지정 끝
        End Select
        '=========================================================================
    End Sub

    Sub extension_type_3()          'CSV 파일 읽어오기 fileopen - intput으로 구현 실험용

        Dim asc_URL As String
        Dim input_string As String

        Dim cRow As Integer
        Dim i As Integer
        Dim dump_num As Integer

        asc_URL = "TEXT;" & Open_Dir

        '=========================================================================
        '기본 혹은 위치지정 선택 
        Select Case Label20.Text
        '=========================================================================

            Case "기본"
                '=========================================================================
                XL = CreateObject("Excel.application")


                XL.Workbooks.open(Open_Dir)       '성적서 오픈


                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가


                FileOpen(1, CSV_Dir, OpenMode.Input)


                Do Until EOF(1)
                    cRow = cRow + 1
                    input_string = LineInput(1)
                    Dim input_arry() As String = Split(input_string, ",", 10)

                    Select Case dump_num
                        Case 0      'dump
                            For i = 0 To UBound(input_arry)
                                XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                            Next
                            dump_num = 1
                            cRow = 0
                        Case 1
                            For i = 0 To UBound(input_arry)
                                XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                            Next
                    End Select



                Loop


                'XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / 28              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                Dim count As Integer

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 9
                    Do Until Line_Count > 36

                        XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명
                        XL.Sheets(Sheet_Count).Range("B" & Line_Count).value = XL.Sheets(1).Range("D" & Cell_Address).value     '구성요소
                        XL.Sheets(Sheet_Count).Range("C" & Line_Count).value = XL.Sheets(1).Range("E" & Cell_Address).value     '측정값
                        XL.Sheets(Sheet_Count).Range("D" & Line_Count).value = XL.Sheets(1).Range("F" & Cell_Address).value     '설계치
                        XL.Sheets(Sheet_Count).Range("E" & Line_Count).value = XL.Sheets(1).Range("G" & Cell_Address).value     '오차
                        XL.Sheets(Sheet_Count).Range("F" & Line_Count).value = XL.Sheets(1).Range("H" & Cell_Address).value     '상한
                        XL.Sheets(Sheet_Count).Range("G" & Line_Count).value = XL.Sheets(1).Range("I" & Cell_Address).value     '하한
                        XL.Sheets(Sheet_Count).Range("H" & Line_Count).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        If XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = "()" Then
                            XL.Sheets(Sheet_Count).Range("A" & Line_Count).value = ""
                            Exit Do
                        End If
                        Line_Count = Line_Count + 1
                        Cell_Address = Cell_Address + 1
                        count += 1

                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================

                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False

             '=========================================================================
            Case "위치 지정"
                '=========================================================================
                Dim Label_ad(1) As String
                Dim measure_ad(1) As String
                Dim component_ad(1) As String
                Dim design_ad(1) As String
                Dim error_ad(1) As String
                Dim UP_Tol_ad(1) As String
                Dim Low_Tol_ad(1) As String
                Dim judge_ad(1) As String
                Dim Line_count_ad As Integer
                Dim Result_Form_dir As String
                Dim select_check_value() As String
                ReDim select_check_value(9)

                Label_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(0) = Ad_Str(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(0) = Ad_Str(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(0) = Ad_Str(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(0) = Ad_Str(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(0) = Ad_Str(GetINIValue("custom_match_info", "judge", ini_dir))
                Line_count_ad = GetINIValue("custom_match_info", "Line_count", ini_dir)
                Result_Form_dir = GetINIValue("custom_match_info", "Result_Form_dir", ini_dir)


                Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                select_check_value(0) = GetINIValue("check", "label", ini_dir)
                select_check_value(1) = GetINIValue("check", "Measure_value", ini_dir)
                select_check_value(2) = GetINIValue("check", "Design_value", ini_dir)
                select_check_value(3) = GetINIValue("check", "error", ini_dir)
                select_check_value(4) = GetINIValue("check", "UP_tol", ini_dir)
                select_check_value(5) = GetINIValue("check", "Low_tol", ini_dir)
                select_check_value(6) = GetINIValue("check", "judge", ini_dir)
                select_check_value(7) = GetINIValue("check", "component", ini_dir)



                XL = CreateObject("Excel.application")
                gXL = CreateObject("Excel.application")

                XL.Workbooks.open(Result_Form_dir)       '성적서 오픈
                gXL.Workbooks.open(CSV_Dir)         'csv파일 오픈

                XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가

                FileOpen(1, CSV_Dir, OpenMode.Input)


                Do Until EOF(1)
                    cRow = cRow + 1
                    input_string = LineInput(1)
                    Dim input_arry() As String = Split(input_string, ",", 10)

                    Select Case dump_num
                        Case 0      'dump
                            For i = 0 To UBound(input_arry)
                                XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                            Next
                            dump_num = 1
                            cRow = 0
                        Case 1
                            For i = 0 To UBound(input_arry)
                                XL.sheets(1).cells(cRow, i + 1).value = input_arry(i)
                            Next
                    End Select



                Loop

                'XL.visible = True
                XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
                XL.sheets(3).select                                         '워크시트 선택
                Cell_Address = 1                    '워크시트1 셀 위치
                Sheet_Count = 3                     '데이터 입력 워크시트 카운트
                Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
                Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
                Quetient = Cell_Count / Line_count_ad              '페이지수 계산
                TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

                Form6.ProgressBar1.Maximum = Cell_Count

                If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

                Do      '데이터 입력

                    Line_Count = 1

                    Label_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Label", ini_dir))
                    component_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "component", ini_dir))
                    measure_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "measure_value", ini_dir))
                    design_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Design_value", ini_dir))
                    error_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "error", ini_dir))
                    UP_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "UP_Tol", ini_dir))
                    Low_Tol_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "Low_Tol", ini_dir))
                    judge_ad(1) = Ad_NUM(GetINIValue("custom_match_info", "judge", ini_dir))

                    Do Until Line_Count > Line_count_ad

                        If select_check_value(0) = True Then
                            XL.Sheets(Sheet_Count).Range(Label_ad(0) & Label_ad(1)).value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명
                        End If

                        If select_check_value(7) = True Then
                            XL.Sheets(Sheet_Count).Range(component_ad(0) & component_ad(1)).value = XL.Sheets(1).Range("D" & Cell_Address).value     '구성요소
                        End If

                        If select_check_value(1) = True Then
                            XL.Sheets(Sheet_Count).Range(measure_ad(0) & measure_ad(1)).value = XL.Sheets(1).Range("E" & Cell_Address).value     '측정값
                        End If

                        If select_check_value(2) = True Then
                            XL.Sheets(Sheet_Count).Range(design_ad(0) & design_ad(1)).value = XL.Sheets(1).Range("F" & Cell_Address).value     '설계치
                        End If

                        If select_check_value(3) = True Then
                            XL.Sheets(Sheet_Count).Range(error_ad(0) & error_ad(1)).value = XL.Sheets(1).Range("G" & Cell_Address).value     '오차
                        End If

                        If select_check_value(4) = True Then
                            XL.Sheets(Sheet_Count).Range(UP_Tol_ad(0) & UP_Tol_ad(1)).value = XL.Sheets(1).Range("H" & Cell_Address).value     '상한
                        End If

                        If select_check_value(5) = True Then
                            XL.Sheets(Sheet_Count).Range(Low_Tol_ad(0) & Low_Tol_ad(1)).value = XL.Sheets(1).Range("I" & Cell_Address).value     '하한
                        End If

                        If select_check_value(6) = True Then
                            XL.Sheets(Sheet_Count).Range(judge_ad(0) & judge_ad(1)).value = XL.Sheets(1).Range("J" & Cell_Address).value     '판정 // 통과/실패
                        End If

                        '====================================================================================================================================
                        '====================================================================================================================================
                        '라인 끝 빈공간 용
                        XL.Sheets(Sheet_Count).Range("BA1").value = XL.Sheets(1).Range("B" & Cell_Address).value & "(" & XL.Sheets(1).Range("C" & Cell_Address).value & ")"   '라벨명

                        If XL.Sheets(Sheet_Count).Range("BA1").value = "()" Then
                            XL.Sheets(Sheet_Count).Range("BA1").value = ""
                            Exit Do
                        End If
                        XL.Sheets(Sheet_Count).Range("BA1").delete
                        '====================================================================================================================================
                        '====================================================================================================================================
                        '셀주소 하나씩 내리기
                        Label_ad(1) = Label_ad(1) + 1
                        component_ad(1) = component_ad(1) + 1
                        measure_ad(1) = measure_ad(1) + 1
                        design_ad(1) = design_ad(1) + 1
                        error_ad(1) = error_ad(1) + 1
                        UP_Tol_ad(1) = UP_Tol_ad(1) + 1
                        Low_Tol_ad(1) = Low_Tol_ad(1) + 1
                        judge_ad(1) = judge_ad(1) + 1
                        '====================================================================================================================================
                        Cell_Address = Cell_Address + 1
                        Line_Count = Line_Count + 1
                        '====================================================================================================================================
                        '====================================================================================================================================
                        Form6.ProgressBar1.Value += 1
                        '====================================================================================================================================
                    Loop
                    If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                        Exit Do
                    End If
                    XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))

                    Sheet_Count = Sheet_Count + 1

                Loop Until Sheet_Count > (TotalSheet + 2)

                XL.Sheets(1).visible = False
                XL.Sheets(2).visible = False


                '=========================================================================
                '기본혹은 위치지정 끝
        End Select
        '=========================================================================
    End Sub
End Class