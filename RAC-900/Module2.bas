Attribute VB_Name = "Module2"
Public pcdll(40000) As Byte
Public Readbuffer(612) As Byte             'Receive Buffer
Public Com, id As Long   'ComPort , NodeId
Public Com_check As Integer
Public ch_real As Integer
Public curdate As String
Public curtime As String
Public formatedcurdate As String
Public gStatus As String

Function FN_opencom(ByVal pscom As String, ByVal psbaud As String)
'Get COM PORT
Dim lbaud As Long

lbaud = Val(psbaud) 'set bit per sec
Com_check = TSMOpenComm(pscom, lbaud, 110, 8, 1, 2048, 2048, pcdll(1))
If Com_check >= 0 Then
   Com = Com_check
End If

FN_opencom = Com_check
End Function
Function FN_closecom(ByVal port As Integer)
FN_com_close = TSMCloseComm(pcdll(1))
End Function
Function FN_date()
    Dim yy, mm, dd, ww, i As Integer
    Dim Set_date As String
    Dim y(20) As Byte
 
    yy = Year(Now()) Mod 100
    mm = Month(Now())
    dd = Day(Now())
    ww = Weekday(Now())
 
    Set_date = Format(yy, "00") & Format(mm, "00") & Format(dd, "00") & _
               Format(ww, "00")
   
    For i = 1 To Len(Set_date)
        y(i) = Mid(Set_date, i, 1)
    Next i
 
    Call TSM_STDAT23(pcdll(1), id, y(1), 0)
End Function
Function FN_get_date()
    Dim dd(512) As Byte
    Call TSM_GTDAT22(pcdll(1), id, dd(1), 0)
    sdate = ""
    For i = 1 To 6
        sdate = sdate & Chr(dd(i))
    Next
    FN_get_date = "20" & sdate
End Function
Function FN_get_time()
    Dim dd(512) As Byte
    Call TSM_GTTIM24(pcdll(1), id, dd(1), 0)
    stime = ""
    For i = 1 To 6
        stime = stime & Chr(dd(i))
    Next
    FN_get_time = stime
End Function

Function FN_idactive(id)
    Dim iret
    FN_idactive = TSM_ENQND05(pcdll(1), id)
End Function
Function fn_read_parm(block, addr)
    Dim dd(512) As Byte
    dd(1) = Str(block)
    dd(2) = Str(addr)
    Call TSM_GTPBY91(pcdll(1), id, dd(1), 0)
    fn_read_parm = Val(dd(1))
End Function

Function fn_write_parm(block, addr, value)
    Dim dd(512) As Byte
    dd(1) = Str(block)
    dd(2) = Str(addr)
    dd(3) = Str(value)
    Call TSM_STPBY92(pcdll(1), id, dd(1), 0)
End Function
Function FN_time()
    Dim hh, mm, ss, i As Integer
    Dim Set_time As String
    Dim y(20) As Byte
 
    hh = Hour(Now())
    mm = Minute(Now())
    ss = Second(Now())
 
    Set_time = Format(hh, "00") & Format(mm, "00") & Format(ss, "00")
    For i = 1 To Len(Set_time)
        y(i) = Mid(Set_time, i, 1)
    Next i
  
    Call TSM_STTIM25(pcdll(1), id, y(1), 0)
End Function

Function FN_save(p_id)    'Save Data
    'Declare variable
    Dim Length, num, i, check, irecno, ierr, iretry, count As Integer
    Dim Str_tmp, s, s1, Tmp, sout As String
    Dim scard, stime, sdate, sshift, scard10, sshift1, srec1, skey As String
    Dim sAppPath, sINIfile, outpath1, out650s, errorPath As String
    Dim outpath As String * 255
    Dim out650 As String * 255
    Dim sCheck650 As String
    curtime = Format$(Time, "hhmmss")
    formatedcurdate = Format$(Now, "ddMMyyyy")
    check = 0
    
    'Check Application path is valid or not
    sAppPath = App.Path
    If Right(sAppPath, 1) <> "\" Then
      sAppPath = sAppPath & "\"
    End If
    
    Dim logPth As String
    logPth = sAppPath & "log.txt"

    'Make accessable file path for access
    'RTA600 file
    sINIfile = sAppPath & "RTA600.INI"
    'Get key value form file
    iret = GetPrivateProfileString("Output", "opath", "D:\DATA\", outpath, 255, sINIfile)
    outpath1 = Left$(outpath, iret) 'Tream the value
    
    'initial variable
    s = ""
    Str_tmp = ""
    Tmp = ""
    id = p_id
    ierr = 5
    irecno = 1
    count = 0
'Receive Data
    Do While ierr > 0
      DoEvents
      
      Length = TSM_RSTND09(pcdll(1), id, Readbuffer(1))
      MsgBox "Check: " & Chr(Readbuffer(4)) & Chr(Readbuffer(5)), vbInformation
        For i = 4 To Length
           s = s & Chr(Readbuffer(i))
        Next i
        Dim aDate, aTime As String
        aDate = FN_get_date()
        aTime = FN_get_time()
        
        Open logPth For Append As #3
        Print #3, s & " -- " & aDate & " -- " & aTime
        Close #3
        MsgBox "Buffer: " & s & aDate & " -- " & aTime, vbInformation
      
        'MsgBox "data:" & s, vbInformation
        'Decode receive data
        Do While InStr(s, "#") > 0
        '---------------------
           srec1 = Left(s, InStr(s, "#") - 1)
           cu = 1
           Do While InStr(srec1, ":") > 0
            Select Case cu
                Case 1
                    'card no
                    scard = Left(srec1, InStr(srec1, ":") - 1)
                Case 2
                    'date
                    sdateTime = Left(srec1, InStr(srec1, ":") - 1)
                    If CInt(sdateTime) > 60001 And CInt(sdateTime) < 50001 Then
                    stime = sdateTime
                    sdate = Format$(Now, "yyyyMMdd")
                    Else
                    stime = sdateTime
                    sdate = Format$(DateAdd("d", -1, Now), "yyyyMMdd")
                    End If
             End Select
             srec1 = Mid(srec1, InStr(srec1, ":") + 1)
             cu = cu + 1
            Loop
           
           'Formated output String
              sout = Format(id, "000") & ":" & scard & ":" & sdate & ":" & Left(stime, 6) & ":11"
              soutfile = outpath1 & formatedcurdate & "-RTA.txt"
              Open soutfile For Append As #1
              Print #1, sout
              RSTATUS.sout.Caption = sout
              irecno = irecno + 1
              count = count + 1
              Close #1
           
           'Process next package
           s = Mid(s, InStr(s, "#") + 1)
        Loop
        Call TSM_ACKGN06(pcdll(1), id)
    Loop
    
End Function

