Attribute VB_Name = "Module2"
Public pcdll(20000) As Byte
Public Readbuffer(512) As Byte             'Receive Buffer
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
      'Compare receive data is correct?
      Length = TSM_RSTND09(pcdll(1), id, Readbuffer(1))
      s = ""
      s1 = ""
      'Check length
      If Length > 3 Then
         For i = 4 To Length
            s1 = s1 & Chr(Readbuffer(i))
         Next i
      End If
    
      Length = TSM_RSTND09(pcdll(1), id, Readbuffer(1))
      If Length > 3 Then
         For i = 4 To Length
            s = s & Chr(Readbuffer(i))
         Next i
        Open logPth For Append As #3
        Print #3, s
        Close #3
         'MsgBox "Buffer: " & s, vbInformation
      End If
      'Compare resule is difference, iERR - 1
      If s <> s1 Then
         Length = -999
      Else
         ierr = 5
      End If
      
      If Length > 3 Then
         
         MsgBox "data:" & s, vbInformation
         'Decode receive data
         Do While InStr(s, "#") > 0
            srec1 = Left(s, InStr(s, "#") - 1)
            cu = 1
            Do While InStr(srec1, ":") > 0
                Select Case cu
                    Case 1
                        'card no
                        scard = Left(srec1, InStr(srec1, ":") - 1)
                    Case 2
                        'date
                        sdate = Left(srec1, InStr(srec1, ":") - 1)
                        If Left(sdate, 2) > "80" Then
                            sdate = "19" & Mid(sdate, 1, 6)
                        Else
                            sdate = "20" & Mid(sdate, 1, 6)
                        End If
                    Case 3
                        'time
                         stime = Left(srec1, InStr(srec1, ":") - 1)
                    Case 4
                        'shift
                         sshift = Left(srec1, InStr(srec1, ":") - 1)
                    Case 5
                         'keypad
                         skey = Left(srec1, InStr(srec1, ":") - 1)
                    End Select
                srec1 = Mid(srec1, InStr(srec1, ":") + 1)
                cu = cu + 1
            Loop
            If Len(srec1) > 0 Then
                skey = srec1
            End If
            'Export to file
            'Fix Len to 10 bytes
            scard8 = ""
            If Len(scard) > 10 Then
                scard10 = Right(scard, 10)
            ElseIf Len(scard) < 10 Then
                scard10 = Space(10 - Len(scard)) & scard
            Else
                scard10 = scard
            End If
            

            'Intial Stage
            sCheck650 = "OK"
            errorPath = sAppPath & "error.log"
            
            'Check Time Format
            If Len(stime) <> 6 Or Mid(stime, 1, 2) > "23" Or Mid(stime, 3, 2) > "59" Or Mid(stime, 5, 2) > "59" Then
                Open errorPath For Append As #2
                Print #2, s
                Print #2, "Error Time Format!"
                Close #2
                sCheck650 = "Err1"
                Open logPth For Append As #3
                Print #3, "Error Time Format!" & id
                Close #3
                'End
            End If
            
            'Check Date format
            If Mid(sdate, 5, 2) > "12" Or Mid(stime, 7, 2) > "31" Then
                Open errorPath For Append As #2
                Print #2, s
                Print #2, "Error Date Format!"
                Close #2
                sCheck650 = "Err2"
                Open logPth For Append As #3
                Print #3, "Error Date Format!" & id
                Close #3
            End If
            
            'Check time & date format is valid or not
            If sCheck650 = "OK" Then
            check = 1
            'Formated output String
               sout = Format(id, "000") & ":" & scard10 & ":" & sdate & ":" & Left(stime, 6) & ":11"
               soutfile = outpath1 & formatedcurdate & "-RTA.txt"
               Open soutfile For Append As #1
               Print #1, sout
               RSTATUS.sout.Caption = sout
               irecno = irecno + 1
               count = count + 1
               Close #1
            End If
            
            'Process next package
            s = Mid(s, InStr(s, "#") + 1)
         Loop
         
         Call TSM_ACKGN06(pcdll(1), id)
      Else
        If Length = 3 Then
            If check = 0 Then
                RSTATUS.sout.Caption = "NID:" & id & " -- No Data!"
                Open logPth For Append As #3
                Print #3, "NID:" & id & " -- No Data!"
                Close #3
            Else
                RSTATUS.sout.Caption = "NID:" & id & " -- Data Received!"
                Open logPth For Append As #3
                Print #3, "NID:" & id & " -- Data " & count & " Received!"
                Close #3
            End If
            Exit Do
        Else
            ierr = ierr - 1
        End If
      End If
      
    Loop
    
End Function

