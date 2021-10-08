Public Class Class1


    Public Shared Function MID_Check(input_SN As String) As String
        Dim user_SN() As String
        Dim i As Integer
        ReDim user_SN(5000)

        user_SN(0) = "931330708"                'my com
        user_SN(1) = "2005489994"               'M3-QV-WLI
        user_SN(2) = "2108496212"               'M3-QV-Hybrid
        user_SN(3) = "1198022579"               'M3-CRT-AV574
        user_SN(4) = "1243865260"               '(주)21세기  QV
        user_SN(5) = "1557817102"               '(주)21세기  CMM
        user_SN(6) = "1645236319"               'RPM STRATO
        user_SN(7) = "1316229804"               'RPM CRT-AS574
        user_SN(8) = "768389765"                 'RPM CRT-AS7106
        user_SN(9) = "1939151463"               '성진세미텍 QV-X606P1S-D
        user_SN(10) = "1523059546"             '부산 M3 - CRT-AV6168
        user_SN(11) = "796538077"               '구미 H-테크 - CRT-AV-7106
        user_SN(12) = "1290459658"             '구미 이레테크 - CRT-AS9106

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
        QV_install_chk = "SOFTWARE\MEI\QVPak\" & QV_version & "\QVClientMenu Config"                     '각 버전 폴더 진입   HKEY_LOCAL_MACHINE\
        reg_del = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(QV_install_chk, True)
        'reg_del.SetValue("MenuName12", "")
        ' reg_del.SetValue("CommandLine12", "")
        reg_del.DeleteValue("MenuName12", True)
        reg_del.DeleteValue("CommandLine12", True)

    End Sub


End Class
