VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "오디오 재생기"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   10065
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Timer timSR 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Interval        =   90
      Left            =   10440
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   4695
      Begin VB.TextBox txtLyr 
         BackColor       =   &H8000000F&
         Height          =   1935
         Left            =   2160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   38
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblG 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   2000
      End
      Begin VB.Label lblSP 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   2000
      End
      Begin VB.Label lblLP 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   2000
      End
      Begin VB.Label lblYear 
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label lblAlbum 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label lblTrackNumber 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label lblArtist 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2000
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   5040
      TabIndex        =   29
      Top             =   2400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
   End
   Begin VB.Timer timSBManager 
      Left            =   10440
      Top             =   1800
   End
   Begin VB.Timer timVizManager 
      Interval        =   150
      Left            =   10440
      Top             =   4560
   End
   Begin VB.CommandButton CommandButton7 
      Caption         =   "60>"
      Height          =   375
      Left            =   8400
      TabIndex        =   27
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton6 
      Caption         =   "30>"
      Height          =   375
      Left            =   7920
      TabIndex        =   26
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton5 
      Caption         =   "10>"
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton4 
      Caption         =   "<10"
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "<30"
      Height          =   375
      Left            =   6480
      TabIndex        =   23
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "<60"
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "▲"
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "▼"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   5280
      Width           =   375
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   1815
      Left            =   9480
      TabIndex        =   12
      Top             =   3480
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   3201
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      TickFrequency   =   0
      Value           =   50
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   2190
      Left            =   2640
      TabIndex        =   11
      Top             =   480
      Width           =   2295
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5760
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "준비"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   10440
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "열기(&O)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   4815
   End
   Begin VB.CommandButton go 
      BackColor       =   &H00E0E0E0&
      Caption         =   "→"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   2250
      Hidden          =   -1  'True
      Left            =   120
      Pattern         =   "*.mp3;*.mid;*.wma;*.rmi;*.midi;*.mp1;*.mp2;*.mpg;*.mpeg;*.wav;*.wave;*.midi;*.rmi;*.wpl;*.aac;*.amr;*.m4a;*.snd"
      System          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   91
      Left            =   5040
      Picture         =   "main.frx":0442
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   90
      Left            =   5040
      Picture         =   "main.frx":18BE
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   89
      Left            =   5040
      Picture         =   "main.frx":2C72
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   88
      Left            =   5040
      Picture         =   "main.frx":4053
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   87
      Left            =   5040
      Picture         =   "main.frx":52B2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   86
      Left            =   5040
      Picture         =   "main.frx":639C
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   85
      Left            =   5040
      Picture         =   "main.frx":7498
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   84
      Left            =   5040
      Picture         =   "main.frx":8736
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   83
      Left            =   5040
      Picture         =   "main.frx":9ACC
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   82
      Left            =   5040
      Picture         =   "main.frx":AE91
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   81
      Left            =   5040
      Picture         =   "main.frx":C239
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   80
      Left            =   5040
      Picture         =   "main.frx":D6EF
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   79
      Left            =   5040
      Picture         =   "main.frx":EC3F
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   78
      Left            =   5040
      Picture         =   "main.frx":1009D
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   77
      Left            =   5040
      Picture         =   "main.frx":11507
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   76
      Left            =   5040
      Picture         =   "main.frx":12964
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   75
      Left            =   5040
      Picture         =   "main.frx":13DC3
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   74
      Left            =   5040
      Picture         =   "main.frx":1523A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   73
      Left            =   5040
      Picture         =   "main.frx":16967
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   72
      Left            =   5040
      Picture         =   "main.frx":1812B
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   71
      Left            =   5040
      Picture         =   "main.frx":1954C
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   70
      Left            =   5040
      Picture         =   "main.frx":1A96D
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   69
      Left            =   5040
      Picture         =   "main.frx":1BDFC
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   68
      Left            =   5040
      Picture         =   "main.frx":1D1AC
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   67
      Left            =   5040
      Picture         =   "main.frx":1E63B
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   66
      Left            =   5040
      Picture         =   "main.frx":1F9EB
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   65
      Left            =   5040
      Picture         =   "main.frx":20E8B
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   64
      Left            =   5040
      Picture         =   "main.frx":2221F
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   63
      Left            =   5040
      Picture         =   "main.frx":235BD
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   62
      Left            =   5040
      Picture         =   "main.frx":24948
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   61
      Left            =   5040
      Picture         =   "main.frx":25DD2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   60
      Left            =   5040
      Picture         =   "main.frx":27227
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   59
      Left            =   5040
      Picture         =   "main.frx":28675
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   58
      Left            =   5040
      Picture         =   "main.frx":29A9D
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   57
      Left            =   5040
      Picture         =   "main.frx":2AE93
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   56
      Left            =   5040
      Picture         =   "main.frx":2C27A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   55
      Left            =   5040
      Picture         =   "main.frx":2D65F
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   54
      Left            =   5040
      Picture         =   "main.frx":2EA4C
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   53
      Left            =   5040
      Picture         =   "main.frx":2FDF9
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   52
      Left            =   5040
      Picture         =   "main.frx":311AD
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   51
      Left            =   5040
      Picture         =   "main.frx":32545
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   50
      Left            =   5040
      Picture         =   "main.frx":3392D
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   49
      Left            =   5040
      Picture         =   "main.frx":34E46
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   48
      Left            =   5040
      Picture         =   "main.frx":3622A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   47
      Left            =   5040
      Picture         =   "main.frx":37632
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   46
      Left            =   5040
      Picture         =   "main.frx":38A67
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   45
      Left            =   5040
      Picture         =   "main.frx":39E64
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   44
      Left            =   5040
      Picture         =   "main.frx":3B1FC
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   43
      Left            =   5040
      Picture         =   "main.frx":3C5B7
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   42
      Left            =   5040
      Picture         =   "main.frx":3D842
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   41
      Left            =   5040
      Picture         =   "main.frx":3E6E3
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   40
      Left            =   5040
      Picture         =   "main.frx":3F73B
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   39
      Left            =   5040
      Picture         =   "main.frx":409F0
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   38
      Left            =   5040
      Picture         =   "main.frx":41D3A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   37
      Left            =   5040
      Picture         =   "main.frx":430BE
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   36
      Left            =   5040
      Picture         =   "main.frx":444C3
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   35
      Left            =   5040
      Picture         =   "main.frx":457E0
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   34
      Left            =   5040
      Picture         =   "main.frx":460C1
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   33
      Left            =   5040
      Picture         =   "main.frx":46CD1
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   32
      Left            =   5040
      Picture         =   "main.frx":47B8E
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   31
      Left            =   5040
      Picture         =   "main.frx":48BC4
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   30
      Left            =   5040
      Picture         =   "main.frx":49CF0
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   29
      Left            =   5040
      Picture         =   "main.frx":4B00A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   28
      Left            =   5040
      Picture         =   "main.frx":4C3BE
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   27
      Left            =   5040
      Picture         =   "main.frx":4D8FD
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   26
      Left            =   5040
      Picture         =   "main.frx":4ED86
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   25
      Left            =   5040
      Picture         =   "main.frx":50248
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   24
      Left            =   5040
      Picture         =   "main.frx":516B3
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   23
      Left            =   5040
      Picture         =   "main.frx":52A79
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   22
      Left            =   5040
      Picture         =   "main.frx":53E72
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   21
      Left            =   5040
      Picture         =   "main.frx":552A2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   20
      Left            =   5040
      Picture         =   "main.frx":5674C
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   19
      Left            =   5040
      Picture         =   "main.frx":57B5E
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   18
      Left            =   5040
      Picture         =   "main.frx":58EC2
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   17
      Left            =   5040
      Picture         =   "main.frx":5A2CC
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   16
      Left            =   5040
      Picture         =   "main.frx":5B6E6
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   15
      Left            =   5040
      Picture         =   "main.frx":5CB23
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   14
      Left            =   5040
      Picture         =   "main.frx":5DFCA
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   13
      Left            =   5040
      Picture         =   "main.frx":5F47B
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   12
      Left            =   5040
      Picture         =   "main.frx":608BB
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   11
      Left            =   5040
      Picture         =   "main.frx":61C94
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   10
      Left            =   5040
      Picture         =   "main.frx":63055
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   9
      Left            =   5040
      Picture         =   "main.frx":644E6
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   8
      Left            =   5040
      Picture         =   "main.frx":658F9
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   7
      Left            =   5040
      Picture         =   "main.frx":66D5A
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   6
      Left            =   5040
      Picture         =   "main.frx":681C3
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   5
      Left            =   5040
      Picture         =   "main.frx":69682
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   4
      Left            =   5040
      Picture         =   "main.frx":6A9D5
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   3
      Left            =   5040
      Picture         =   "main.frx":6B9AF
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   2
      Left            =   5040
      Picture         =   "main.frx":6C895
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   1
      Left            =   5040
      Picture         =   "main.frx":6D521
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblNext 
      BackStyle       =   0  '투명
      Caption         =   "▶"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   41
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  '투명
      Caption         =   "◀"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   40
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "◀"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "맨 앞으로"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '투명
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8880
      TabIndex        =   39
      Top             =   3045
      Width           =   255
   End
   Begin VB.Image imgVisBalls 
      Height          =   1725
      Index           =   0
      Left            =   5040
      Picture         =   "main.frx":6DE22
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin WMPLibCtl.WindowsMediaPlayer mplayer 
      Height          =   240
      Left            =   4800
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   300
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   529
      _cy             =   423
   End
   Begin VB.Label ToggleButton1 
      BackStyle       =   0  '투명
      Caption         =   "◀×"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   21
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label mutev 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   20
      Top             =   480
      Width           =   855
   End
   Begin VB.Label CommandButton8 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  '투명
      Caption         =   "▶"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Label CommandButton9 
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H008080FF&
      BackStyle       =   0  '투명
      Caption         =   "||"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   4020
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '투명
      Caption         =   "●"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "■"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "음량"
      Height          =   255
      Left            =   9600
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   10920
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label status 
      Caption         =   "1"
      Height          =   495
      Left            =   10800
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "0"
      Height          =   255
      Left            =   8880
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image imgVizBlank 
      Height          =   1725
      Left            =   5040
      Picture         =   "main.frx":6E473
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4935
   End
   Begin VB.Image imgControllerBackground 
      Height          =   2775
      Left            =   5040
      Picture         =   "main.frx":133AD5
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Menu file 
      Caption         =   "파일(&F)"
      Begin VB.Menu open 
         Caption         =   "열기(&O)"
      End
      Begin VB.Menu sprtor 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "종료(&X)"
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "재생(&P)"
      Begin VB.Menu mnuPlayRepeatSection 
         Caption         =   "구간 반복(&S)..."
      End
      Begin VB.Menu mnuPlayCancelSectionRepeat 
         Caption         =   "구간 반복 해제(&E)"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "옵션(&O)"
      Begin VB.Menu mnuOptionsLoop 
         Caption         =   "반복(&L)"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep43555 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAnims 
         Caption         =   "애니메이션(&A)"
         Begin VB.Menu mnuAnimsOption 
            Caption         =   "기본효과 &1"
            Index           =   1
         End
         Begin VB.Menu mnuAnimsOption 
            Caption         =   "기본효과 &2"
            Index           =   2
         End
         Begin VB.Menu mnuAnimsOption 
            Caption         =   "기본효과 &3"
            Index           =   3
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "도움말(&H)"
      Begin VB.Menu log 
         Caption         =   "변경 사항(&U)"
         Visible         =   0   'False
      End
      Begin VB.Menu about 
         Caption         =   "정보(&A)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vi As Integer
Dim vizidx As Integer

Public SRS As Double
Public SRE As Double

Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
    mplayer.URL = File1.Path & "\" & File1.FileName
    mplayer.currentMedia.getItemInfo ("Author")
    CommandButton8.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
    mplayer.settings.volume = mplayer.settings.volume - 5
    Slider2.Value = Slider2.Value + 5
End Sub

Private Sub Command3_Click()
On Error Resume Next
    mplayer.settings.volume = mplayer.settings.volume + 5
    Slider2.Value = Slider2.Value - 5
End Sub

Private Sub CommandButton10_Click()
    
End Sub



Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition - 60
    Slider1.Value = Slider1.Value - 60
End Sub

Private Sub CommandButton3_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition - 30
    Slider1.Value = Slider1.Value - 30
End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition - 10
    Slider1.Value = Slider1.Value - 10
End Sub

Private Sub CommandButton5_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition + 10
    Slider1.Value = Slider1.Value + 10
End Sub

Private Sub CommandButton6_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition + 30
    Slider1.Value = Slider1.Value + 30
End Sub

Private Sub CommandButton7_Click()
On Error Resume Next
    mplayer.Controls.currentPosition = mplayer.Controls.currentPosition + 60
    Slider1.Value = Slider1.Value + 60
End Sub

Private Sub CommandButton8_Click()
On Error Resume Next
    CommandButton9.Visible = True
    CommandButton8.Visible = False
    mplayer.Controls.play
    Timer1.Enabled = True
    'status.Caption = "1"
    'Do While mplayer.Controls.currentPosition < mplayer.currentMedia.duration And status.Caption = "1"
        'Slider1.Value = mplayer.Controls.currentPosition
        'Label1.Caption = mplayer.Controls.currentPosition
    'Loop
    StatusBar1.SimpleText = "재생 중 - " & mplayer.currentMedia.Name
End Sub

Private Sub CommandButton9_Click()
On Error Resume Next
    CommandButton9.Visible = False
    CommandButton8.Visible = True
    mplayer.Controls.pause
    'status.Caption = "0"
    Timer1.Enabled = False
    StatusBar1.SimpleText = "일시 중지"
    
End Sub

Private Sub Dir1_Change()
On Error Resume Next
    File1.Path = Dir1.Path
    Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    File1.Path = Drive1.Drive
    Dir1.Path = Drive1.Drive
End Sub

Private Sub exit_Click()

    End
End Sub

Private Sub Form_Load()
On Error Resume Next
    vizidx = CInt(GetSetting(App.Title, "Options", "Animation", 1))
    vi = 0
    mnuAnimsOption_Click vizidx
    File1.Path = GetSetting(App.Title, "Config", "Path", "C:\WINDOWS\MEDIA\")
    Text1.Text = GetSetting(App.Title, "Config", "Path", "C:\WINDOWS\MEDIA\")
    Slider1.Value = 0
    'status.Caption = "0"
    mplayer.settings.setMode "loop", True
    mplayer.settings.mute = False
    mutev.Caption = "1"
    ToggleButton1.Caption = "◀))"
    Dir1.Path = GetSetting(App.Title, "Config", "Path", "C:\WINDOWS\MEDIA\")
    mplayer.enableContextMenu = False
    
    SRS = -1
    SRE = -1
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 2 Then
Me.WindowState = 0
End If
    Me.Width = 10500
    Me.Height = 7125
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Config", "Path", Text1.Text
End Sub

Private Sub go_Click()
On Error Resume Next
    File1.Path = Text1.Text
    Dir1.Path = Text1.Text
    Drive1.Drive = Dir1.Path
End Sub

Private Sub hj_Click()

End Sub

Private Sub Label10_Click()
On Error Resume Next
    mplayer.Controls.stop
    Timer1.Enabled = False
    Label1.Caption = "0"
    Slider1.Value = 0
    'status.Caption = "0"
    StatusBar1.SimpleText = "정지"
    CommandButton9.Visible = False
    CommandButton8.Visible = True
End Sub

Private Sub Label11_Click()
On Error Resume Next
mplayer.settings.volume = 75
Slider2.Value = 75
End Sub

Private Sub Label5_Click()
On Error Resume Next
    mplayer.settings.volume = 0
    Slider2.Value = 0
End Sub

Private Sub Label6_Click()
On Error Resume Next
mplayer.settings.volume = 25
Slider2.Value = 25
End Sub

Private Sub Label7_Click()
On Error Resume Next
mplayer.settings.volume = 50
Slider2.Value = 50
End Sub

Private Sub Label8_Click()
On Error Resume Next
mplayer.Controls.currentPosition = 0
End Sub

Private Sub Label9_Click()
On Error Resume Next
Slider2.Value = 100
mplayer.settings.volume = 100
End Sub

Private Sub lblBack_Click()
    mplayer.Controls.Previous
End Sub

Private Sub lblNext_Click()
    mplayer.Controls.Next
End Sub

Private Sub log_Click()
    ulog.Show
End Sub

Private Sub mnuAnimsOption_Click(Index As Integer)
    SaveSetting App.Title, "Options", "Animation", Index
    vizidx = Index
    
    Select Case Index
        Case 1
            imgVisBalls(vi).Visible = False
            vi = 0
            imgVisBalls(vi).Visible = True
        Case 2
            imgVisBalls(vi).Visible = False
            vi = 36
            imgVisBalls(vi).Visible = True
        Case 3
            imgVisBalls(vi).Visible = False
            vi = 56
            imgVisBalls(vi).Visible = True
    End Select
End Sub

Private Sub mnuOptionsLoop_Click()
    mnuOptionsLoop.Checked = Not mnuOptionsLoop.Checked
    mplayer.settings.setMode "loop", mnuOptionsLoop.Checked
End Sub

Private Sub mnuPlayCancelSectionRepeat_Click()
    SRS = -1
    SRE = -1
    timSR.Enabled = False
End Sub

Private Sub mnuPlayRepeatSection_Click()
    frmSectionRepeat.Show '모달 일부러 안 함
End Sub

Private Sub mplayer_MediaChange(ByVal Item As Object)
    On Error Resume Next
    'status.Caption = "1"
    StatusBar1.SimpleText = "재생 중 - " & mplayer.currentMedia.Name
    Timer1.Enabled = True
    Label3.Caption = mplayer.currentMedia.duration
    Slider1.Max = mplayer.currentMedia.duration
    Label2.Caption = Fix(mplayer.currentMedia.duration * 100) / 100

    'Do While mplayer.Controls.currentPosition < mplayer.currentMedia.duration And status.Caption = "1"
        'Slider1.Value = mplayer.Controls.currentPosition
        'Label1.Caption = mplayer.Controls.currentPosition
    'Loop
    CommandButton9.Visible = True
    CommandButton8.Visible = False
    
    lblArtist.Caption = mplayer.currentMedia.getItemInfo("Author")
    lblAlbum.Caption = mplayer.currentMedia.getItemInfo("WM/AlbumTitle")
    lblTrackNumber.Caption = "#" & mplayer.currentMedia.getItemInfo("WM/TrackNumber")
    lblYear.Caption = mplayer.currentMedia.getItemInfo("WM/Year") & "년"
    lblG.Caption = mplayer.currentMedia.getItemInfo("WM/Genre")
    
    lblLP.Caption = mplayer.currentMedia.getItemInfo("WM/Writer") & " 작사"
    lblSP.Caption = mplayer.currentMedia.getItemInfo("WM/Composer") & " 작곡"
    
    txtLyr.Text = mplayer.currentMedia.getItemInfo("WM/Lyrics")
End Sub

Private Sub open_Click()
On Error Resume Next
    mplayer.URL = File1.Path & "\" & File1.FileName
End Sub



Private Sub shigakwa_Click()
    If shigakwa.Checked = True Then
        Image2.Visible = True
        mplayer.Height = 1
        shigakwa.Checked = False
    Else
        Image2.Visible = False
        mplayer.Height = 1695
        shigakwa.Checked = True
    End If
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next
    mplayer.Controls.currentPosition = Slider1.Value
End Sub



Private Sub Slider2_Change()
On Error Resume Next
    mplayer.settings.volume = (100 - Slider2.Value)
    timVizManager.Interval = 100 + Slider2.Value
End Sub

Private Sub Slider2_Scroll()
On Error Resume Next
    mplayer.settings.volume = (100 - Slider2.Value)
    timVizManager.Interval = 100 + Slider2.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Label1.Caption = Fix(mplayer.Controls.currentPosition * 100) / 100
    Slider1.Value = mplayer.Controls.currentPosition
End Sub

Private Sub timSR_Timer()
    If mplayer.Controls.currentPosition > SRE Then
        mplayer.Controls.currentPosition = SRS
    End If
End Sub

Private Sub timVizManager_Timer()
    On Error Resume Next
    If Timer1.Enabled = True Then
        imgVizBlank.Visible = True
        imgVisBalls(vi).Visible = False
        
        vi = vi + 1
        
        Select Case vizidx
            Case 1
                If vi > 35 Then vi = 0
            Case 2
                If vi > 55 Then vi = 36
            Case 3
                If vi > 91 Then vi = 56
        End Select
        
        imgVisBalls(vi).Visible = True
        'Debug.Print vi
    Else
        imgVisBalls(vi).Visible = False
        imgVizBlank.Visible = True
    End If
End Sub

Private Sub ToggleButton1_Click()
On Error Resume Next
    If mutev.Caption = "1" Then
        mplayer.settings.mute = False
        ToggleButton1.Caption = "◀))"
        mutev.Caption = "0"
    Else
        mplayer.settings.mute = True
        ToggleButton1.Caption = "◀×"
        mutev.Caption = "1"
    End If
End Sub

