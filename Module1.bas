Attribute VB_Name = "Module1"
'Open Com Port
Public Declare Function TSMOpenComm Lib "TSMCOM32.DLL" (ByVal port As String, ByVal p1 As Long, ByVal p2 As Byte, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long, ByVal p6 As Long, ByRef pcdll As Byte) As Integer

'Close Com Port
Public Declare Function TSMCloseComm Lib "TSMCOM32.DLL" (ByRef pcdll As Byte) As Integer

Public Declare Function TSMSetRespondPeriod Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal iPeriod As Integer) As Integer
'Set TimeOut
Public Declare Function TSMSetTimeout Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal iTimeout As Integer) As Integer
'Clean Data
Public Declare Function TSM_CLMSP04 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long) As Integer
'Console inquiry working node for link status
Public Declare Function TSM_ENQND05 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long) As Integer
'Affirmative acknowledgments
Public Declare Function TSM_ACKGN06 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long) As Integer
'Request working node to transmit scanning datum.
Public Declare Function TSM_RSTND09 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcReadBuffer As Byte) As Integer
'Get Date
Public Declare Function TSM_GTDAT22 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcReadBuffer As Byte, ByVal nConfirm As Integer) As Integer
'Set Date
Public Declare Function TSM_STDAT23 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nConfirm As Integer) As Integer
'Get Time
Public Declare Function TSM_GTTIM24 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcReadBuffer As Byte, ByVal nConfirm As Integer) As Integer
'Set Time
Public Declare Function TSM_STTIM25 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nConfirm As Integer) As Integer
Public Declare Function TSM_54 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pnMode As Integer, nConfirm As Integer) As Integer
Public Declare Function TSM_55 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByVal nMode As Integer, ByVal nConfirm As Integer) As Integer

Public Declare Function TSM_LDMES62 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte) As Integer

Public Declare Function TSM_STOPS76 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nSendLen As Integer) As Integer
'Function int                               TSM_76( ref char pcdlls[20000], int nNodeId,      int nBlock,          REF char pcBadge[150]  ,int nBadgeLen ) Library "TSMCOM32.DLL"
Public Declare Function TSM_76 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByVal nBlock As Long, ByRef pcBadge As Byte, ByVal nbadgelen As Long) As Integer
Public Declare Function TSM_STTPS78 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nSendLen As Integer) As Integer

Public Declare Function TSM_DESIP7c Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nSendLen As Integer) As Integer
'Function int  TSM_90                             (ref char pcdlls[20000], int nNodeId,        REF char pcBadge[150], int nBadgeLen,         REF char pcPinCode[150],   int nPinCodeLen,         REF char pcName[150], int nNameLenn ) Library "TSMCOM32.DLL"
Public Declare Function TSM_90 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcBdage As Byte, ByVal nbadgelen As Long, ByRef pcpincode As Byte, ByVal npincodelen As Long, ByRef pcname As Byte, ByVal nnamelen As Long) As Integer
Public Declare Function TSM_a3 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcBadge As Byte, ByVal nbadgelen As Long, ByRef pcpincode As Byte, ByVal npincodelen As Long, ByRef pcname As Byte, ByVal nnamelen As Long) As Integer
'Function int  TSM_GTPBY91( ref char pcdlls[20000],                         int nNodeId,           REF char pcReadBuffer[150],  int nConfirm ) Library "TSMCOM32.DLL"
Public Declare Function TSM_GTPBY91 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, ByRef pcReadBuffer As Byte, ByVal nConfirm As Integer) As Integer
Public Declare Function TSM_STPBY92 Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nNodeId As Long, pcSend As Byte, ByVal nConfirm As Integer) As Integer

'Read Com
Public Declare Function TSMReadCommChar Lib "TSMCOM32.DLL" (ByRef pcdll As Byte) As Integer
'Write Com
Public Declare Function TSMWriteCommChar Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal ch As Long) As Integer

Public Declare Function TSMGetInBuf Lib "TSMCOM32.DLL" (ByRef pcdll As Byte) As Integer

Public Declare Function TSMDelay Lib "TSMCOM32.DLL" (ByVal uDelay As Integer)

Public Declare Function TSMCleanInBuf Lib "TSMCOM32.DLL" (ByRef pcdll As Byte, ByVal nCleanBytes As Integer)

'Public Declare Function TSMVersion Lib "TSMCOM32.DLL" ()

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Read INI Files

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Write INI
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'
Public Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long



