VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toDay, previosDay As String
Dim a As Integer

Private Sub Form_Load()
a = showDate()

End Sub

Function showDate()
toDay = Format$(Now, "yyyyMMdd:hhmmss")
previosDay = Format$(DateAdd("d", -1, Now), "yyyyMMdd:hhmmss")
lbl1.Caption = "Today- " & toDay
lbl2.Caption = "Pev -" & previosDay
End Function
