VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  '沒有框線
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ForeColor       =   &H8000000D&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   3360
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   3360
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame5"
      Height          =   2055
      Left            =   2400
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox Check31 
         BackColor       =   &H00808080&
         Caption         =   "洗頻"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2400
         TabIndex        =   69
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  '平面
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   68
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Timer Timer14 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2760
         Top             =   120
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00808080&
         Caption         =   "自動開始"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer Timer11 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1200
         Top             =   120
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00808080&
         Caption         =   "智能準備"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      Height          =   2055
      Left            =   2400
      TabIndex        =   61
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Timer Timer26 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Top             =   120
      End
      Begin VB.Timer Timer25 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4200
         Top             =   120
      End
      Begin VB.CheckBox Check29 
         BackColor       =   &H00808080&
         Caption         =   "秒終魔10S"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
      Begin VB.Timer Timer21 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1320
         Top             =   120
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00808080&
         Caption         =   "8心鎖時"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   240
         Width           =   1095
      End
   End
   Begin SHDocVwCtl.WebBrowser webLogin 
      Height          =   135
      Left            =   8040
      TabIndex        =   45
      Top             =   120
      Width           =   255
      ExtentX         =   450
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
      Location        =   "http:///"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   2055
      Left            =   2400
      TabIndex        =   17
      Top             =   480
      Width           =   5415
      Begin VB.CheckBox chkMemory 
         BackColor       =   &H00808080&
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
         Left            =   840
         TabIndex        =   43
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox password 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Height          =   270
         IMEMode         =   3  '暫止
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   42
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox username 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         Height          =   270
         Left            =   720
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Timer Timer19 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5040
         Top             =   960
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00808080&
         Caption         =   "快速入場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Timer Timer16 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3840
         Top             =   960
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00808080&
         Caption         =   "陸地遊水"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2520
         Top             =   960
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00808080&
         Caption         =   "陸地滑雪"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00808080&
         Caption         =   "遊戲置頂"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00808080&
         Caption         =   "隱藏遊戲"
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3840
         Top             =   480
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00808080&
         Caption         =   "無敵狀態"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   600
         Width           =   1095
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2520
         Top             =   480
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00808080&
         Caption         =   "隨時離場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   600
         Width           =   1095
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1200
         Top             =   480
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00808080&
         Caption         =   "快速收田"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5040
         Top             =   0
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00808080&
         Caption         =   "永不禁言"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3840
         Top             =   0
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00808080&
         Caption         =   "無限字數"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2520
         Top             =   0
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00808080&
         Caption         =   "無限怒氣"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1200
         Top             =   0
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "無限藍氣"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   120
         Width           =   1095
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
         Left            =   3000
         TabIndex        =   44
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label labPass 
         BackStyle       =   0  '透明
         Caption         =   "密碼:"
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
         TabIndex        =   41
         Top             =   1440
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
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   2055
      Left            =   2400
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox Check24 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Timer Timer22 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4080
         Top             =   600
      End
      Begin VB.CheckBox Check22 
         BackColor       =   &H00808080&
         Caption         =   "持續憤怒"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   55
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer Timer20 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2640
         Top             =   720
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00808080&
         Caption         =   "附身場主"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   54
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer Timer18 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1080
         Top             =   720
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00808080&
         Caption         =   "排除障礙"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.Timer Timer17 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4320
         Top             =   120
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00808080&
         Caption         =   "踩板無限沖"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3120
         TabIndex        =   52
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00808080&
         Caption         =   "改頻"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  '平面
         Height          =   270
         Left            =   1200
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
      Begin VB.Timer Timer13 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   480
         Top             =   120
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00808080&
         Caption         =   "個人教學任務改場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   2400
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox Check30 
         BackColor       =   &H00808080&
         Caption         =   "三王完場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2760
         TabIndex        =   67
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer Timer24 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1080
         Top             =   600
      End
      Begin VB.CheckBox Check28 
         BackColor       =   &H00808080&
         Caption         =   "小紅帽完場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   60
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check27 
         BackColor       =   &H00808080&
         Caption         =   "四心完場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check26 
         BackColor       =   &H00808080&
         Caption         =   "雪二完場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4080
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check25 
         BackColor       =   &H00808080&
         Caption         =   "雪一完場"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer Timer15 
         Left            =   2280
         Top             =   120
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00808080&
         Caption         =   "連續多圈"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.Timer Timer12 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1200
         Top             =   120
      End
      Begin VB.CheckBox Check23 
         BackColor       =   &H00808080&
         Caption         =   "一鍵一圈"
         BeginProperty Font 
            Name            =   "微軟正黑體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2160
   End
   Begin VB.Label Label20 
      BackStyle       =   0  '透明
      Caption         =   "掛機功能"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6480
      TabIndex        =   46
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00808080&
      X1              =   7320
      X2              =   7320
      Y1              =   360
      Y2              =   120
   End
   Begin VB.Label Label19b 
      BackColor       =   &H8000000D&
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "八心功能"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5520
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label18b 
      BackColor       =   &H8000000D&
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  '透明
      Caption         =   "完場功能"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label17b 
      BackColor       =   &H8000000D&
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  '透明
      Caption         =   "進階功能"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label16b 
      BackColor       =   &H8000000D&
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   3480
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  '透明
      Caption         =   "普通功能"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label14b 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00808080&
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "Windows 7 64位元的玩家 ,在HShield出現時按 ""注入外掛"""
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   4815
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
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
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
      Left            =   960
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
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
      Left            =   960
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
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
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
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
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
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
      Left            =   960
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   7920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000D&
      X1              =   0
      X2              =   8400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   3240
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   3120
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   -240
      X2              =   7920
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "注入外掛"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "房間密碼:"
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
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "玩家點數:"
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
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "玩家金錢:"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "玩家經驗: "
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
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "遊戲名稱:"
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
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   8040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   0
      Y2              =   2640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "W Cheat x3.0"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
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
Dim Hig As Boolean
Public Function WriteMemory(ByVal lpAddress As Long, ByVal lpBuffer As Long, ByVal lpSize As Long) As Long
WriteMemory = WriteProcessMemory(pHandle, ByVal lpAddress, ByVal lpBuffer, ByVal lpSize, False)
End Function
Public Function ReadLong(ByVal lpAddress As Long) As Long
    Dim Value As Long
    ReadProcessMemory pHandle, ByVal lpAddress, ByVal VarPtr(Value), ByVal 4, False
    ReadLong = Value
End Function
Public Sub Delay(DelayTime As Single)
Dim ST As Single
  ST = Timer
  Do Until Timer - ST > DelayTime
    DoEvents
  Loop
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
On Error Resume Next
Dim V As Long
Dim EAX As Long
V = InputBox("請輸入改頻數值", "溫馨提示")
ReadProcessMemory Tr_hproc, &HEE66A8, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &HC, VarPtr(V), 4, 0
End If
End Sub

Private Sub Check17_Click()
Timer17.Enabled = Check17.Value
End Sub

Private Sub Check18_Click()
Timer18.Enabled = Check18.Value
End Sub

Private Sub Check19_Click()
Timer19.Enabled = Check19.Value
End Sub

Private Sub Check21_Click()
Timer21.Enabled = Check21.Value
End Sub

Private Sub Check22_Click()
Timer22.Enabled = Check22.Value
End Sub

Private Sub Check23_Click()
Timer12.Enabled = Check23.Value
End Sub

Private Sub Check24_Click()
If Check24.Value = 1 Then
Form3.Show
MsgBox "如果部份功能沒有反應,請重新選取並按↑制", vbOKOnly, "提醒"
Else
Unload Form3
End If
End Sub

Private Sub Check25_Click()

Timer24.Enabled = Check25.Value
End Sub

Private Sub Check26_Click()
Timer24.Enabled = Check26.Value
End Sub

Private Sub Check27_Click()

Timer24.Enabled = Check27
End Sub

Private Sub Check28_Click()
Timer24.Enabled = Check28.Value
End Sub

Private Sub Check29_Click()
Timer26.Enabled = Check29.Value

End Sub

Private Sub Check30_Click()
Timer23.Enabled = Check30.Value
End Sub

Private Sub Check31_Click()
Timer27.Enabled = Check31.Value
End Sub

Private Sub Check7_Click()
Timer9.Enabled = Check7.Value
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
Private Sub Close_Click()
End
End Sub
Private Sub Check1_Click()
Timer3.Enabled = Check1.Value
End Sub
Private Sub Check10_Click()
Timer10.Enabled = Check10.Value
End Sub
Private Sub Check11_Click()
If Check11.Value = 1 Then
Timer13.Enabled = True
Else
Timer13.Enabled = False
End If
End Sub
Private Sub Check12_Click()
Timer11.Enabled = Check12.Value
End Sub

Private Sub Check13_Click()
Timer14.Enabled = Check13.Value
End Sub

Private Sub Check14_Click()
Timer15.Enabled = Check14.Value
End Sub

Private Sub Check2_Click()
Timer4.Enabled = Check2.Value
End Sub
Private Sub Check3_Click()
Timer5.Enabled = Check3.Value
End Sub
Private Sub Check4_Click()
Timer6.Enabled = Check4.Value
End Sub
Private Sub Check5_Click()
Timer7.Enabled = Check5.Value
End Sub
Private Sub Check6_Click()
Timer8.Enabled = Check6.Value
End Sub
Private Sub Check9_Click()
Dim hwn As Long
If Check9.Value = 1 Then
hwn = FindWindow(vbNullString, "Tales Runner") '搵跑Online窗口,可以將Tales Runner更換為其他名稱
SetWindowPos hwn, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE '置頂
Else
hwn = FindWindow(vbNullString, "Tales Runner")
SetWindowPos hwn, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE '取消置頂
End If
End Sub

Private Sub Form_Load()
 webLogin.Navigate ("http://weblogin.talesrunner.com.hk/web_login.html")
    username = GetSetting("vbPath", "txt", "username")
    password = GetSetting("vbPath", "txt", "password")

Call Crack_Hs
End Sub
Private Function SEND_MESSAGE(nMessage As String)

Dim hwnd As Long
hwnd = FindWindow(vbNullString, "Tales Runner")
  Dim data() As Byte, I As Long
  I = 0#
  data = StrConv(nMessage, vbFromUnicode)

  While I <= UBound(data)
      If data(I) < 128 Then
          PostMessage hwnd, &H102, data(I), 0&
          I = I + 1
      Else
          PostMessage hwnd, &H102, data(I), 0&
          PostMessage hwnd, &H102, data(I + 1), 0&
          I = I + 2
      End If
  Wend
  PostMessage hwnd, &H100, vbKeyReturn, 0&
End Function



Private Sub Label14_Click()
Label14b.BackStyle = 1
Label16b.BackStyle = 0
Label17b.BackStyle = 0
Label18b.BackStyle = 0
Label19b.BackStyle = 0
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label16_Click()
Label16b.BackStyle = 1
Label14b.BackStyle = 0
Label17b.BackStyle = 0
Label18b.BackStyle = 0
Label19b.BackStyle = 0
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label17_Click()
MsgBox "所有功能完場時間為45S"
Label16b.BackStyle = 0
Label14b.BackStyle = 0
Label17b.BackStyle = 1
Label18b.BackStyle = 0
Label19b.BackStyle = 0
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
End Sub

Private Sub Label18_Click()
Label16b.BackStyle = 0
Label14b.BackStyle = 0
Label17b.BackStyle = 0
Label18b.BackStyle = 1
Label19b.BackStyle = 0
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
End Sub

Private Sub Label20_Click()
Label16b.BackStyle = 0
Label14b.BackStyle = 0
Label17b.BackStyle = 0
Label18b.BackStyle = 0
Label19b.BackStyle = 1
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End Sub

Private Sub Label7_Click()
Timer1.Enabled = True
Tr_hwnd = FindWindow(vbNullString, "Tales Runner") '第一個是類名, 第二個標題
If Tr_hwnd = 0 Then MsgBox "找不到遊戲", vbCritical, "": Exit Sub '如果遊戲沒有打開，則退出
GetWindowThreadProcessId Tr_hwnd, Tr_pid '取得遊戲的PID值
Tr_hproc = OpenProcess(&H1F0FFF, False, Tr_pid) '以PID值打開進程
If Tr_hproc = 0 Then
MsgBox "注入外掛失敗", vbCritical, "提醒"
Else
MsgBox "成功注入外掛", vbOKOnly, "提醒"
End If
End Sub

Private Sub Timer1_Timer()
Dim Value As String * 12
ReadProcessMemory Tr_hproc, ByVal &HE9EB0C, ByVal Value, Len(Value), 0 '遊戲名稱
Label8.Caption = "" & Value

Dim Value1 As Long
ReadProcessMemory Tr_hproc, &HE9EB6C, ByVal VarPtr(Value1), 4, 0 '遊戲經驗
Label9.Caption = Value1

Dim Value2 As Long
ReadProcessMemory Tr_hproc, &HE9EB5C, ByVal VarPtr(Value2), 4, 0 '遊戲金錢
Label10.Caption = Value2

Dim Value3 As Long
ReadProcessMemory Tr_hproc, &HE9EB60, ByVal VarPtr(Value3), 4, 0 '遊戲點數
Label11.Caption = Value3
End Sub
Private Sub Timer10_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H14C, VarPtr(1), 4, 0 '陸地滑雪
End Sub
Private Sub Timer11_Timer()
Dim Value As Long
ReadProcessMemory Tr_hproc, &HECF2DC, ByVal VarPtr(Value), 4, 0 '智能準備
If Value = 1 Then
PostMessage Tr_hwnd, &H100, &H78, 0& '按下
PostMessage Tr_hwnd, &H101, &H78, 0& '彈上
End If
End Sub
Private Sub Timer12_Timer()
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
Check23.Value = 0

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
Check23.Value = 0
End If

End Sub

Private Sub Timer13_Timer()
On Error Resume Next
Dim V As Long
Dim EAX As Long
V = Text2.Text
ReadProcessMemory Tr_hproc, &HEC8954, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &HC, VarPtr(V), 4, 0
End Sub

Private Sub Timer14_Timer()
PostMessage Tr_hwnd, &H100, &H79, 0& '按下
PostMessage Tr_hwnd, &H101, &H79, 0& '彈上
End Sub

Private Sub Timer15_Timer()
Dim time As Long
Dim door As Long
ReadProcessMemory Tr_hproc, &HF40B30, ByVal VarPtr(time), 4, 0
ReadProcessMemory Tr_hproc, &HECC5EC, ByVal VarPtr(door), 4, 0
If door = 9 Then
If time >= 2700 And time <= 2900 Then

Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H28, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1000), 4, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(5000), 2, 0
Delay 0.1
WriteProcessMemory Tr_hproc, ByVal EAX + &H630, VarPtr(1), 2, 0
End If
End If
End Sub

Private Sub Timer16_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H124, VarPtr(1), 4, 0 '陸地滑雪
End Sub

Private Sub Timer17_Timer()
WriteProcessMemory Tr_hproc, &HEC94C2, ByVal VarPtr(-1), 4, 0
End Sub

Private Sub Timer18_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H106, VarPtr(16479), 4, 0
End Sub

Private Sub Timer19_Timer()
WriteProcessMemory Tr_hproc, &HEF4924, ByVal VarPtr(0), 4, 0

End Sub

Private Sub Timer2_Timer()
Label13.Caption = (Now)
Dim VALUE4 As String * 12
ReadProcessMemory Tr_hproc, ByVal &HE9EC5C, ByVal VALUE4, Len(VALUE4), 0
Label12.Caption = "" & VALUE4
End Sub

Private Sub Timer20_Timer()
WriteProcessMemory Tr_hproc, &HEC8750, ByVal VarPtr(0), 4, 0 '無限字數

End Sub

Private Sub Timer21_Timer()
WriteProcessMemory Tr_hproc, &HEE7194, ByVal VarPtr(0), 4, 0
End Sub

Private Sub Timer22_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H116, VarPtr(16285), 4, 0 '無限藍氣
End Sub

Private Sub Timer23_Timer()
Dim time As Long
Dim door As Long
ReadProcessMemory Tr_hproc, &HF44B38, ByVal VarPtr(time), 4, 0
ReadProcessMemory Tr_hproc, &HECF2C0, ByVal VarPtr(door), 4, 0
If door = 10 Then
If time >= 1 Then

Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &HCC, VarPtr(1), 4, 0
End If
End If
If door = 10 Then
If time >= 2700 And time <= 2900 Then
Dim EAX1 As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX1), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX1 + &H50, VarPtr(EAX1), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX1 + &H8F4, VarPtr(2100000000), 4, 0
End If
End If
End Sub

Private Sub Timer24_Timer()
Dim time As Long
Dim door As Long
ReadProcessMemory Tr_hproc, &HF44B38, ByVal VarPtr(time), 4, 0
ReadProcessMemory Tr_hproc, &HECF2C0, ByVal VarPtr(door), 4, 0

If door = 10 Then
If time >= 2700 And time <= 2760 Then
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
End If
End If

End Sub

Private Sub Timer25_Timer()
Dim time As Long
Dim door As Long
ReadProcessMemory Tr_hproc, &HF44B38, ByVal VarPtr(time), 4, 0
ReadProcessMemory Tr_hproc, &HECF2C0, ByVal VarPtr(door), 4, 0

If door = 10 Then
If time >= 1 And time <= 60 Then
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1187394726), 4, 0 '高度
End If
End If

If door = 10 Then
If time >= 150 And time <= 160 Then
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1187515121), 4, 0 '高度

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H610, VarPtr(1201656663), 4, 0 '前後

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H614, VarPtr(1108353613), 4, 0 '左右

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H8F4, VarPtr(1132000000), 4, 0 '速p
End If
End If

If door = 10 Then
If time >= 170 And time <= 180 Then
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1187505791), 4, 0 '高度

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H610, VarPtr(1206448486), 4, 0 '前後

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H614, VarPtr(1067710042), 4, 0 '左右

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H8F4, VarPtr(1132000000), 4, 0 '速p
End If
End If

If door = 10 Then
If time >= 190 And time <= 200 Then
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1187619135), 4, 0 '高度

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H610, VarPtr(1206833405), 4, 0 '前後

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H614, VarPtr(1108160687), 4, 0 '左右

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H8F4, VarPtr(1132000000), 4, 0 '速p
End If
End If

If door = 10 Then
If time >= 210 And time <= 220 Then
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1187986784), 4, 0 '高度

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H610, VarPtr(1207162694), 4, 0 '前後

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H614, VarPtr(1126119689), 4, 0 '左右

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H8F4, VarPtr(1132000000), 4, 0 '速p
End If
End If

If door = 10 Then
If time >= 500 And time <= 1200 Then
ReadProcessMemory Tr_hproc, &HEE7048, ByVal VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H7DF, VarPtr(30000), 4, 0 '扣王

ReadProcessMemory Tr_hproc, &HEE7048, ByVal VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H7C3, VarPtr(256), 4, 0 '定王1

ReadProcessMemory Tr_hproc, &HEE7048, ByVal VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H7C3, VarPtr(256), 4, 0 '定王2
End If
End If

If door = 10 Then
If time >= 500 And time <= 1200 Then

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H8F4, VarPtr(1162000000), 4, 0 '速p

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H618, VarPtr(1188355490), 4, 0 '高度

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H610, VarPtr(1207834544), 4, 0 '前後

ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H0, VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H614, VarPtr(&HC4301702), 4, 0 '左右
End If
End If
End Sub

Private Sub Timer26_Timer()
Timer25.Enabled = Timer26.Enabled
End Sub

Private Sub Timer27_Timer()
SEND_MESSAGE Text1.Text 'SEND TEXT內容
End Sub

Private Sub Timer3_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H7D0, VarPtr(4294967295#), 4, 0 '無限藍氣
End Sub
Private Sub Timer4_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &H7CC, VarPtr(4294967295#), 4, 0 '無限紅氣
End Sub
Private Sub Timer5_Timer()
WriteProcessMemory Tr_hproc, &HF30B58, ByVal VarPtr(999), 4, 0 '無限字數
End Sub
Private Sub Timer6_Timer()
WriteProcessMemory Tr_hproc, &HEE788C, ByVal VarPtr(1), 4, 0 '永不禁言
End Sub
Private Sub Timer7_Timer()
WriteProcessMemory Tr_hproc, &HECF2C0, ByVal VarPtr(1), 4, 0 '快速收田
PostMessage Tr_hwnd, &H100, &H1B, 0& '按下
PostMessage Tr_hwnd, &H101, &H1B, 0& '彈上
End Sub
Private Sub Timer8_Timer()
WriteProcessMemory Tr_hproc, &HEF48C4, ByVal VarPtr(1), 4, 0 '隨時離場
End Sub
Private Sub Timer9_Timer()
Dim EAX As Long
ReadProcessMemory Tr_hproc, &HED05FC, ByVal VarPtr(EAX), 4, 0
ReadProcessMemory Tr_hproc, ByVal EAX + &H50, VarPtr(EAX), 4, 0
WriteProcessMemory Tr_hproc, ByVal EAX + &HCC, VarPtr(1), 4, 0 '無敵模式
End Sub
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





