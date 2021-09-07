VERSION 5.00
Begin VB.Form RSTATUS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Vistasoft IT Bangladesh Ltd."
   ClientHeight    =   1455
   ClientLeft      =   7485
   ClientTop       =   2310
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RSTATUS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   6090
   Begin VB.Label sout 
      Alignment       =   2  'Center
      Caption         =   "Wait For Data !"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Node?"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5625
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "RSTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim pp(10) As String
Dim scom, sids, sauto, sid, sbaud, para, ss As String
Dim ireturn, iret, iids As Integer
Dim iTimeout As Long
Dim sINI As String * 255
Dim sAppPath, sINIfile As String
para = Command$
curdate = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
If UCase(Command$) = "/S" Then
   Exit Sub
End If

RSTATUS.Show
RSTATUS.msg.Caption = "Starting..."

sAppPath = App.Path
If Right(sAppPath, 1) <> "\" Then
    sAppPath = sAppPath & "\"
End If

sINIfile = sAppPath & "RTA600.INI"
iret = GetPrivateProfileString("Main", "NID01", "", sINI, 255, sINIfile)

Dim logPth As String
logPth = sAppPath & "log.txt"
      
'default 3 nodes
If iret = 0 Then
MsgBox "Please Check RTA600.INI file exist or not?", vbInformation
lblShowMsg.Text = lblShowMsg.Text & "Please Check RTA600.INI file exist or not?\n"

End If


'max. 50 nodes
For iids = 1 To 50
ss = "NID" & Format(iids, "00")
RSTATUS.msg.Caption = "Receive " & ss

'Get Device infp strng from 'RTA600.INI' file
iret = GetPrivateProfileString("Main", ss, "", sINI, 255, sINIfile)

'Stop program when not set INI file
If iret = 0 Or sINI = "" Then
    RSTATUS.msg.Caption = "Receive finish!!! [Check D:\DATA\]"
    Sleep 1000
    End
End If

'Assing Device info string in para variable
para = sINI
'MsgBox para, vbInformation

'Discribe the Device info string
If Len(para) > 0 And InStr(para, ",") > 0 Then
   sid = Left(para, InStr(para, ",") - 1) 'node id
   para = Mid(para, InStr(para, ",") + 1) 'Move cursor
   scom = Left(para, InStr(para, ",") - 1) 'Com port
   para = Mid(para, InStr(para, ",") + 1) 'Move cursor
   sbaud = Left(para, InStr(para, ",") - 1) 'Buffer sise
   para = Mid(para, InStr(para, ",") + 1) 'Move cursor
   iTimeout = Val(para) 'Timeout
        
   'Open COM PORT
   ireturn = FN_opencom(scom, sbaud)
   If ireturn >= 0 Then
      ireturn = TSMSetRespondPeriod(pcdll(1), iTimeout)
      ireturn = TSMSetTimeout(pcdll(1), 100)
      RSTATUS.msg.Caption = "Open COM Port Success, " & scom
        Open logPth For Append As #1
        Print #1, "Open COM Port Success, " & scom
        Close #1
      'MsgBox "Open COM Success, Node" & sid, vbInformation
   Else
      RSTATUS.msg.Caption = "Open COM Port Failure, " & scom
      RSTATUS.sout.Caption = "Failed to open port !!"
        Open logPth For Append As #1
        Print #1, "Open COM Port Failure, " & scom
        Close #1
      'MsgBox "Open COM Failure, Node" & sid, vbInformation
      GoTo nextid
   End If
   
   'Receive Data..
   iid = Val(sid)
   'Check Id is valid or not
   If iid > 0 Then
      DoEvents
      ireturn = FN_idactive(iid)
      If ireturn > 0 Then
         ireturn = fn_write_parm(2, 101, 3)
         ireturn = fn_write_parm(2, 102, 1)
         ireturn = fn_write_parm(2, 103, 2)
         ireturn = fn_write_parm(2, 104, 5)
         ireturn = fn_write_parm(2, 105, 9)  'KeyIn
         ireturn = fn_write_parm(2, 106, 0)
         
         'Dim v_hh, v_mm, v_ss, v_i As Integer
         'Dim Set_time As String
         'Dim y(20) As Byte
         
         'v_hh = Hour(Now())
         'v_mm = Minute(Now())
         'v_ss = Second(Now())
        
         'Set_time = Format(v_hh, "00") & Format(v_mm, "00") & Format(v_ss, "00")
         'For v_i = 1 To Len(Set_time)
         'y(i) = Mid(Set_time, v_i, 1)
         'Next v_i
         
         'Call TSM_STTIM25(pcdll(1), id, y(1), 0)
         'RSTATUS.msg.Caption = "NODE: " & iid & " :: Set Time & Date..."
         
         'Sleep 500
         RSTATUS.msg.Caption = "NODE: " & iid & " :: Data Processing..."
         FN_save (iid)
         RSTATUS.msg.Caption = "NODE: " & iid & " :: Data Received!!"
            Open logPth For Append As #1
            Print #1, "NODE: " & iid & " :: Data Received!!"
            Close #1
      Else
         RSTATUS.msg.Caption = "Not in online!, NODE: " & iid
            Open logPth For Append As #1
            Print #1, "Not in online!, NODE: " & iid
            Close #1
         'MsgBox "Not online, NODE: " & iid
      End If
   End If
   'Close COM port
   FN_closecom (1)
Else
   If Len(para) > 0 Then
      RSTATUS.msg.Caption = "RTA600.INI Data isn't valid!!!"
        Open logPth For Append As #1
        Print #1, "RTA600.INI Data isn't valid!!!"
        Close #1
   End If
End If

nextid:

Next

End Sub
