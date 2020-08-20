Public Class Form1
    Dim Cur_Dir As String
    Dim Savs_Dir As String
    Dim CSV_Dir As String
    Dim Result_Orign As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        '================================

        Dim Cell_Address As Integer
        Dim XL As Object
        Dim gXL As Object
        Dim Sheet_Count As Integer
        Dim Line_Count As Integer
        Dim Quetient As Double

        Dim Cell_Count As Integer
        Dim Cell_Count2 As Integer
        Dim TotalSheet As Integer
        Dim strData() As String
        ReDim strData(10)
        '=========================================================================
        XL = CreateObject("Excel.application")
        gXL = CreateObject("Excel.application")

        XL.Workbooks.open("C:\datasavefile\QVB\dialog test\excel test\For_VB\Result_VB.xlsx")       '성적서 오픈
        gXL.Workbooks.open("C:\datasavefile\QVB\dialog test\excel test\data\easy test.csv")         'csv파일 오픈

        XL.Sheets.add(before:=XL.Sheets(1)) 'csv파일 가져올 워크시트 추가
        gXL.sheets(1).Range("A2:K1000").copy    'csv 값 복사
        XL.Sheets(1).Range("A1").PasteSpecial(-4163)    'csv값 붙여넣기 
        With gXL            'csv파일 종료
            .DisplayAlerts = False
            .workbooks(1).close
            .quit
        End With
        gXL = Nothing
        'XL.visible = True
        XL.Sheets("sheet1").Copy(After:=XL.Sheets("sheet1"))        '데이터 입력용 워크시트 복사
        XL.sheets(3).select                                         '워크시트 선택
        Cell_Address = 1                    '워크시트1 셀 위치
        Sheet_Count = 3                     '데이터 입력 워크시트 카운트
        Cell_Count2 = XL.Sheets(1).Rows.Count       '워크시트 행 위치 검색
        Cell_Count = XL.Sheets(1).Cells(Cell_Count2, 1).End(-4162).Row      '워크시트 행 위치 검색  -4162 : xlUp
        Quetient = Cell_Count / 28              '페이지수 계산
        TotalSheet = CInt(Quetient)             '페이지수 계산 (반올림)

        If Quetient - TotalSheet > 0 Then TotalSheet = TotalSheet + 1       '반올림값 보정 0.5 이하 +1페이지

        Do      '데이터 입력

            Line_Count = 9
            Do Until Line_Count > 34

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
            Loop
            If Sheet_Count = TotalSheet + 2 Then        '같은 페이지일때 점프로 나가기
                GoTo jump
            End If
            XL.Sheets("sheet1").Copy(After:=XL.Sheets(Sheet_Count))
jump:
            Sheet_Count = Sheet_Count + 1

        Loop Until Sheet_Count > (TotalSheet + 2)





        Select Case MsgBox("성적서 출력을 완료했습니다. 자동저장하시겠습니까?", 4)
            Case 6
                With XL
                    .DisplayAlerts = False
                    .workbooks(1).SaveAS("C:\datasavefile\QVB\dialog test\excel test\For_VB\222.xlsx")          '다른이름으로 저장 위치 지정
                    .Workbooks(1).close
                    .quit
                End With
            Case 7
                MsgBox("자동저장을 취소하셨습니다. 성적서를 저장해주세요.")
                XL.visible = True
        End Select

    End Sub

End Class
