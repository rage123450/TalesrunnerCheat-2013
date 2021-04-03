VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ForeColor       =   &H8000000D&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1260
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   2760
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   840
      Top             =   2280
   End
   Begin VB.Label Label6 
      Height          =   2535
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "開啟程式"
      Enabled         =   0   'False
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
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "正在驗測程式版本"
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
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00808080&
      Caption         =   "W Cheat x2.0"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "程式版本:"
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
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Close 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "版本認證"
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   4560
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_STYLE = (-16)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
Private Sub Delay(DelayTime As Single)
Dim ST As Single
  ST = Timer
  Do Until Timer - ST > DelayTime
    DoEvents
  Loop
End Sub
Private Sub Close_Click()
End
End Sub

Private Sub Form_Load()
Dim TempLng As Long
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 1, 230, LWA_ALPHA
Timer1.Enabled = True
Timer2.Enabled = True
End Sub


Private Sub Label5_Click()
Form1.Show
Unload Form2
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Inet1.OpenURL("https://dl.dropbox.com/s/kmp8fs6c5yk1mrq/ver.txt?token_hash=AAG_NWCImdQSRJwipq7KghA16EiNbapznZJuuVDhAUhJUw&dl=1") ' 在免費空間裡上傳個txt檔,裡面輸入版本號碼
If Label6.Caption = Label3.Caption Then '偵測是否與本地的版本相同
Label5.Enabled = True

Else
MsgBox "錯誤版本，請上官網下載最新", vbCritical, "403錯誤"
End
End If
End Sub

Private Sub Timer2_Timer()
If Label5.Enabled = True Then
Label4.Caption = "正在驗測程式版本" & ""
Delay 0.25
Label4.Caption = "正在驗測程式版本" & ""
Delay 0.25
Label4.Caption = "正在驗測程式版本" & ""
Else
Label4.Caption = "正在驗測程式版本" & "."
Delay 0.25
Label4.Caption = "正在驗測程式版本" & ".."
Delay 0.25
Label4.Caption = "正在驗測程式版本" & "..."
End If
End Sub

Private Sub Style_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub FrMcaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
