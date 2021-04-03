VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form4"
   ScaleHeight     =   1575
   ScaleWidth      =   2820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin SHDocVwCtl.WebBrowser webLogin 
      Height          =   135
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CheckBox chkMemory 
      BackColor       =   &H00404040&
      Caption         =   "記住帳號密碼"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox password 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      Height          =   270
      IMEMode         =   3  '暫止
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox username 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1695
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
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.Label cmdLogin 
      BackStyle       =   0  '透明
      Caption         =   "登入"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label labPass 
      BackStyle       =   0  '透明
      Caption         =   "密碼:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label labUser 
      BackStyle       =   0  '透明
      Caption         =   "帳號:"
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
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "登入器"
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
      Top             =   120
      Width           =   1215
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
Attribute VB_Name = "Form4"
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

Private Sub cmdLogin_Click()
On Error Resume Next
   webLogin.Document.getElementById("fdLoginUserid").Value = username.Text
   webLogin.Document.getElementById("fdLoginUserPass").Value = password.Text
   webLogin.Document.getElementById("loginForm").Click

Dim PostData() As Byte
Dim vHeaders As String
    PostData = "fdLoginUserid=&fdLoginUserPass=&szLoginUserid=" + username + "&szLoginUserPass=" + password + "&useNew=1&x=34&y=23"
    PostData = StrConv(PostData, vbFromUnicode)
vHeaders = "Content-Type: application/x-www-form-urlencoded" & _
vbCrLf
    webLogin.Navigate "http://weblogin.talesrunner.com.hk/weblogin.php", , , PostData, vHeaders
        If chkMemory.Value = 1 Then
            SaveSetting "vbPath", "txt", "username", username
            SaveSetting "vbPath", "txt", "password", password
        End If
        Me.Caption = "Mnighthk 原創作品"
End Sub

Private Sub Form_Load()
    webLogin.Navigate ("http://weblogin.talesrunner.com.hk/web_login.html")
    username = GetSetting("vbPath", "txt", "username")
    password = GetSetting("vbPath", "txt", "password")
End Sub

Private Sub Label2_Click()
End
End Sub

Private Sub webLogin_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If InStr(webLogin.Document.body.innerText, "帳號或密碼輸入錯誤") > 0 Then
    MsgBox "帳號或密碼輸入錯誤", 0 + 64, "提示"
    webLogin.Navigate ("http://weblogin.talesrunner.com.hk/web_login.html")
End If
If InStr(webLogin.Document.body.innerText, "啟用錯誤，可能是系統繁忙中，") > 0 Then
    MsgBox "啟用錯誤 ,可能是系統繁忙中,或閣下的帳號並未進行註冊", 0 + 64, "提示"
    webLogin.Navigate ("http://weblogin.talesrunner.com.hk/web_login.html")
End If
If InStr(webLogin.Document.body.innerText, "帳號被禁止使用,若有問題,請洽客戶服務員(TP)!!") > 0 Then
    MsgBox "帳號被禁止使用", 0 + 64, "提示"
    webLogin.Navigate ("http://weblogin.talesrunner.com.hk/web_login.html")
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub



