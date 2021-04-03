VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   BorderStyle     =   0  '沒有框線
   Caption         =   "任務精通"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   FillColor       =   &H00FF8080&
   ForeColor       =   &H8000000D&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2655
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox Check16 
         BackColor       =   &H00808080&
         Caption         =   "第七關"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00808080&
         Caption         =   "第六關"
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00808080&
         Caption         =   "第五關"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00808080&
         Caption         =   "第四關"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00808080&
         Caption         =   "第三關"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00808080&
         Caption         =   "第二關"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00808080&
         Caption         =   "第一關"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '透明
         Caption         =   "暴走精通"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3495
      Begin VB.CheckBox Check9 
         BackColor       =   &H00808080&
         Caption         =   "第九關"
         Height          =   300
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00808080&
         Caption         =   "第八關"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808080&
         Caption         =   "第七關"
         Height          =   225
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00808080&
         Caption         =   "第六關"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
         Caption         =   "第五關"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "第四關"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "第三關"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "第二關"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "第一關"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "Dr.Hell"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "暴走精通"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  '透明
      Caption         =   "任務精通"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "Dr.Hell"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   3840
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Any, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Const SW_SHOW = 5
Const SW_HIDE = 0
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Tr_hwnd As Long         '遊戲的HWND值
Private Tr_pid As Long          '遊戲的的PID值
Private Tr_hproc As Long        '遊戲的HPROC值
Private EAX As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const MF_BYPOSITION = &H400

Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_STYLE = (-16)


Private Sub Check10_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1117613072), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1197766425), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1156029103), 4, 0

End Sub

Private Sub Check11_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1183712019), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1183553520), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1162901479), 4, 0

End Sub

Private Sub Check12_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1183579114), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1172971594), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1171811348), 4, 0
End Sub

Private Sub Check13_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1108588507), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1190806813), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1154318680), 4, 0
End Sub

Private Sub Check14_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1160668120), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1197367685), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1154416590), 4, 0

End Sub

Private Sub Check15_Click()
Dim EAX As Long 'ECC5EC
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1138300810), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(1191520368), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1154336954), 4, 0
End Sub

Private Sub Check16_Click()

Dim EAX As Long 'ECC5EC
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(1167149902), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(&HC5C11133), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1184110318), 4, 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Tr_hwnd = FindWindow(vbNullString, "Tales Runner")
GetWindowThreadProcessId Tr_hwnd, Tr_pid
Tr_hproc = OpenProcess(&H1F0FFF, False, Tr_pid)

If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
Public Sub Delay(DelayTime As Single)
Dim ST As Single
  ST = Timer
  Do Until Timer - ST > DelayTime
    DoEvents
  Loop
End Sub
Private Sub Check1_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HECC5EC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check1.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HEC8844, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check1.Value = 0
End If

End Sub

Private Sub Check2_Click()
On Error Resume Next 'ED05FC
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HECC5EC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HECC5EC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HECC5EC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check2.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HEC8844, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check2.Value = 0
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check3.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HECF320, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check3.Value = 0
End If
End Sub

Private Sub Check5_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check5.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HECF320, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check5.Value = 0
End If
End Sub

Private Sub Check6_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check6.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HECF320, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check6.Value = 0
End If
End Sub

Private Sub Check7_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check7.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HECF320, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check7.Value = 0
End If
End Sub

Private Sub Check8_Click()
On Error Resume Next
Dim time As Long
Dim EAX As Long
Dim V As Long
Dim C As Long
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 4  '時間碼
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Delay 0.1
Check8.Value = 0

Dim F As Long
Dim X As Long
ReadProcessMemory Tr_hproc, ByVal &HECF320, ByVal Name, Len(C), 2 '現在地圖
If F = 1202 Then
ReadProcessMemory Tr_hproc, ByVal &HF40B30, ByVal Name, Len(time), 0 '時間碼


ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(8000), 2, 0 '距離碼
Delay 0.1
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0 '距離碼
Check8.Value = 0
End If
End Sub

Private Sub Check9_Click()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A4, VarPtr(&HC445E1D7), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4A8, VarPtr(&HC3382C8D), 4, 0

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H34, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H4AC, VarPtr(1172737532), 4, 0
End Sub


Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Frame1.Visible = True
Frame2.Visible = False

End Sub

Private Sub Label4_Click()
Frame1.Visible = False
Frame2.Visible = True
End Sub
