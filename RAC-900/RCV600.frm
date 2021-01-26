VERSION 5.00
Begin VB.Form RCV600 
   AutoRedraw      =   -1  'True
   Caption         =   "RTA600DataCollect"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   ScaleHeight     =   6555
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Date/Time"
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox ofile 
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox ids 
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Text            =   "1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox timeouts 
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Text            =   "50"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox coms 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Text            =   "COM1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Receive"
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text_RCV 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3240
      Width           =   8655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label lab_msg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK!"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   5640
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Example: RTA600 1,COM1,9600,100 (Node ID:1,COM1,BaudRate:9600, Timeout 100)"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   7095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Output:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Timeout:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "NodeID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Command line: RTA650 id,com,baud,timeout"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   7095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Receive Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Comm:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "RCV600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim iTimeout, iret As Long

    scom = coms.Text
    sids = ids.Text
    iTimeout = Val(timeouts.Text)
    'Open COM PORT
    ireturn = FN_opencom(scom, "9600")
    If ireturn >= 0 Then
    ireturn = TSMSetRespondPeriod(pcdll(1), iTimeout)
    
    Else
        MsgBox ("COM Port open fail,Return code:(" & ireturn & ")")
    End If
    
    If Val(sids) > 0 Then
            iid = Val(sids)
            DoEvents
            ireturn = FN_idactive(iid)
            If ireturn > 0 Then
                ireturn = fn_write_parm(2, 101, 3)  'CardNo
                ireturn = fn_write_parm(2, 102, 1)  'Date
                ireturn = fn_write_parm(2, 103, 2)  'Time
                ireturn = fn_write_parm(2, 104, 5)  'shift
                ireturn = fn_write_parm(2, 105, 9)  'KeyIn
                ireturn = fn_write_parm(2, 106, 0)
                Sleep 500
                FN_save (iid)
                lab_msg.Caption = "ID:" & iid & "Received!!"
            Else
                lab_msg.Caption = "Not Online, ID:" & iid
            End If
    End If
    FN_closecom (1)
    MsgBox "Finish!!"
End Sub

Private Sub Command2_Click()
    Dim iTimeout As Long
    scom = coms.Text
    id = Val(ids.Text)
    iTimeout = Val(timeouts.Text)
    'Open COM PORT
    ireturn = FN_opencom(scom, "9600")
    If ireturn >= 0 Then
    ireturn = TSMSetRespondPeriod(pcdll(1), iTimeout)
    
    Else
        MsgBox ("Open COM Fail,Return Code:(" & ireturn & ")")
    End If
    
iret = 0
sret = FN_get_date() & " " & FN_get_time()

'iret = fn_read_parm(2, 73)
    FN_closecom (1)
MsgBox (sret)

End Sub

Private Sub Command3_Click()


End Sub

Private Sub Form_Load()
Dim pp(10) As String
Dim scom, sids, sauto, sid, sbaud, para, ss As String
Dim ireturn, iret, iids As Integer
Dim iTimeout As Long
Dim sINI As String * 255
Dim sAppPath, sINIfile As String
para = Command$
curdate = Year(Date) & Format(Month(Date), "00") & Format(Day(Date), "00")
ofile.Text = curdate & ".TXT"
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

'default 3 nodes
If iret = 0 Then
    iret = WritePrivateProfileString("Main", "NID01", "1,COM1,9600,50", sINIfile)
    iret = WritePrivateProfileString("Main", "NID02", "2,COM1,9600,50", sINIfile)
    iret = WritePrivateProfileString("Main", "NID03", "3,COM1,9600,50", sINIfile)
    iret = WritePrivateProfileString("Output", "opath", "C:\CARDATA\BACKUP\", sINIfile)
    iret = WritePrivateProfileString("Output", "ofile650", "C:\CARDATA\RTA600.TXT", sINIfile)
    MsgBox ("Setting Parameter(" & sINIfile & ")")
    End
End If

'max. 50 nodes
For iids = 1 To 50
ss = "NID" & Format(iids, "00")
RSTATUS.msg.Caption = "Receive" & ss
iret = GetPrivateProfileString("Main", ss, "", sINI, 255, sINIfile)
'Stop program when not set INI file
If iret = 0 Or sINI = "" Then
    RSTATUS.msg.Caption = "Receive finish!!!"
    Sleep 1000
    End
End If
para = sINI
If Len(para) > 0 And InStr(para, ",") > 0 Then
   sid = Left(para, InStr(para, ",") - 1)
   para = Mid(para, InStr(para, ",") + 1)
   scom = Left(para, InStr(para, ",") - 1)
   para = Mid(para, InStr(para, ",") + 1)
   sbaud = Left(para, InStr(para, ",") - 1)
   para = Mid(para, InStr(para, ",") + 1)
   iTimeout = Val(para)
   coms.Text = scom
   timeouts.Text = iTimeout
   ids.Text = sbauds
   'ofile.Text = curdate & ".TXT"
    
        
   'Open COM PORT
   ireturn = FN_opencom(scom, sbaud)
   If ireturn >= 0 Then
      ireturn = TSMSetRespondPeriod(pcdll(1), iTimeout)
      ireturn = TSMSetTimeout(pcdll(1), 100)
   Else
      
      RSTATUS.msg.Caption = "Open COM Failure,Return Code(" & ireturn & ")"
      GoTo nextid
   End If
   
   'Receive Data..
   iid = Val(sid)
   MsgBox "Now we will move to the next step " & iid, vbInformation
   
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
         Sleep 500
         FN_save (iid)
         RSTATUS.msg.Caption = "ID:" & iid & "Received!!"
         lab_msg.Caption = "ID:" & iid & "Received!!"
      Else
         RSTATUS.msg.Caption = "Not online!, ID:" & iid
         lab_msg.Caption = "Not online, ID:" & iid
      End If
   End If
   FN_closecom (1)
Else
   If Len(para) > 0 Then
      RSTATUS.msg.Caption = "RTA600.INI Parameter error!!!"
      
   End If
End If

nextid:

Next

End Sub

Private Sub Timer1_Timer()

End Sub
