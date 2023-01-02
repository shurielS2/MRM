Public Class Class1


    Public Shared Function MID_Check(input_SN As String) As String
        Dim user_SN() As String
        Dim i As Integer
        ReDim user_SN(5000)

        user_SN(0) = "1429864061"                'my com
        user_SN(1) = "1243865260"               '(주)21세기  QV
        user_SN(2) = "1557817102"               '(주)21세기  CMM
        user_SN(3) = "1645236319"               'RPM STRATO
        user_SN(4) = "1316229804"               'RPM CRT-AS574
        user_SN(5) = "768389765"                 'RPM CRT-AS7106
        user_SN(6) = "1939151463"               '성진세미텍 QV-X606P1S-D
        user_SN(7) = "796538077"               '구미 H-테크 - CRT-AV-7106   
        user_SN(8) = "1290459658"             '구미 이레테크 - CRT-AS9106
        user_SN(9) = "837651529"              '부천 대신산업 - CRT-AV9106  
        user_SN(10) = "497432293"              '부천 대신산업 - CRT-AS9168
        user_SN(11) = "786381889"               'M3 CRT-AV9166 
        user_SN(12) = "1729431577"               'M3-CRT-AV574  
        user_SN(13) = "401177235"            'M3-Busan QV-L404
        user_SN(14) = "1605543453"            'M3-Busan QV-X302
        user_SN(15) = "1498838607"               'M3-QV-Hybrid     
        user_SN(16) = "2005489994"               'M3-QV-WLI
        user_SN(17) = "1887186651"              'M3-QV-Active 
        user_SN(18) = "1523059546"             '부산 M3 - CRT-AV9168  
        user_SN(19) = "1848064147"              ' MKC 군포 5층 다목적실 
        user_SN(20) = "2024587627"             'TS-김동우 개인컴
        user_SN(21) = "1546861664"              'TS-문승 개인컴
        user_SN(22) = "2076100518"              'TS-이덕형 개인컴
        user_SN(23) = "1639927654"              'TS-임재준 개인컴
        user_SN(24) = "106421636"               'TS-장인환 개인컴
        user_SN(25) = "1546861664"              'TS-정병수 개인컴
        user_SN(26) = "617998793"               'TS-채수군 개인컴
        user_SN(27) = "1860574094"              'TS-김건우 개인컴
        user_SN(28) = "1801828475"              'TS-최성준 개인컴
        user_SN(29) = "658655528"               'TS-김재영 개인컴
        user_SN(30) = "326924221"               'TS-박건진 개인컴
        user_SN(31) = "1825693199"              'TS-이광형 개인컴
        user_SN(32) = "53500892"                'TS-정정훈 개인컴
        user_SN(33) = "1920010456"              'TS-최성준 개인컴
        user_SN(34) = "1424896520"              'TS-하늘 개인컴
        user_SN(35) = "444279140"               'TS-황인혁 개인컴
        user_SN(36) = "2013105385"               '구미 지오엠 CRT-AS9106
        user_SN(37) = "562938361"                '케이에스티이(KSTE)  QV-L404
        user_SN(38) = "1796488296"                '선경화학 QV-L404Z1L-D
        user_SN(39) = "1075522989"                'M3 QV-X302 Gen.E
        user_SN(40) = "891781689"                 'MKC 5층 강습용 CRT-AV544
        user_SN(41) = "733756473"                 'MKC 5층 강습용 QV-Active
        user_SN(42) = "821902057"                'MKC 5층 강습용 QV-Apex
        user_SN(43) = "19296819"                  'MKC M3 QV-Active
        user_SN(44) = "1991707591"                   'MKC 5층 강습용 CRT-AV544
        user_SN(45) = "497475135"               '의신정밀
        user_SN(46) = "1157476282"               '부산 M3 Kogame
        user_SN(47) = "1805180202"               '대구 CRT-AV544
        user_SN(48) = "551642947"                 'RPS 대전 QV-X606P1L-E
        user_SN(49) = "150806381"                 '부산 송주찬 사원 노트북

        If user_SN.Contains(input_SN) Then
            For i = 0 To 5000

                If user_SN.GetValue(i) = input_SN Then

                    Return user_SN.GetValue(i)
                    Exit For
                End If
            Next i

        Else
            Return False
        End If




        Return False
    End Function


    Public Shared Sub reg_del()

        On Error Resume Next
        Dim reg_del As Microsoft.Win32.RegistryKey

        Dim software_chk_QV As String
        Dim QV_install_chk As String
        Dim QV_version As String

        software_chk_QV = "HKEY_LOCAL_MACHINE\SOFTWARE\MEI\QVPak"               '폴더 존재 유무 확인용 
        QV_version = My.Computer.Registry.GetValue(software_chk_QV, "Current Version", Nothing)     'QV 버전 확인
        QV_install_chk = "SOFTWARE\MEI\QVPak\" & QV_version & "\QVClientMenu Config"                 '각 버전 폴더 진입   HKEY_LOCAL_MACHINE\
        reg_del = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(QV_install_chk, True)
        'reg_del.SetValue("MenuName12", "")
        ' reg_del.SetValue("CommandLine12", "")
        reg_del.DeleteValue("MenuName12", True)
        reg_del.DeleteValue("CommandLine12", True)

    End Sub


End Class
