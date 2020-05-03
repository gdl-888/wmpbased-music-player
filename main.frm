VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "음악 재생기"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10095
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton CommandButton7 
      Caption         =   "60>"
      Height          =   375
      Left            =   8400
      TabIndex        =   33
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton6 
      Caption         =   "30>"
      Height          =   375
      Left            =   7920
      TabIndex        =   32
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton5 
      Caption         =   "10>"
      Height          =   375
      Left            =   7440
      TabIndex        =   31
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton4 
      Caption         =   "<10"
      Height          =   375
      Left            =   6960
      TabIndex        =   30
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton3 
      Caption         =   "<30"
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "<60"
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "▲"
      Height          =   255
      Left            =   9600
      TabIndex        =   20
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "▼"
      Height          =   255
      Left            =   9600
      TabIndex        =   19
      Top             =   5280
      Width           =   375
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   1815
      Left            =   9480
      TabIndex        =   17
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   3201
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelStart        =   100
      TickStyle       =   1
      TickFrequency   =   0
      Value           =   100
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   4815
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5760
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "준비"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "열기(&O)"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   4815
   End
   Begin VB.CommandButton go 
      BackColor       =   &H00E0E0E0&
      Caption         =   "→"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   3690
      Left            =   120
      Pattern         =   "*.mp3;*.mid;*.wma;*.rmi;*.midi;*.mp1;*.mp2;*.mpg;*.mpeg;*.wav;*.wave;*.midi;*.rmi;*.wpl"
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   327682
      Max             =   1
   End
   Begin WMPLibCtl.WindowsMediaPlayer mplayer 
      Height          =   1800
      Left            =   5040
      TabIndex        =   34
      Top             =   120
      Width           =   4860
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
      _cx             =   8573
      _cy             =   3175
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
      Left            =   7440
      TabIndex        =   27
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label mutev 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   4080
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
      TabIndex        =   24
      Top             =   4080
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
      Left            =   7440
      TabIndex        =   23
      Top             =   4800
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
      TabIndex        =   22
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "↗"
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
      Left            =   5640
      TabIndex        =   21
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '투명
      Caption         =   "75% >"
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "100% >"
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "50% >"
      Height          =   255
      Left            =   9000
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "25% >"
      Height          =   255
      Left            =   9000
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   " 0% >"
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "음량"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   10920
      TabIndex        =   9
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label status 
      Caption         =   "1"
      Height          =   495
      Left            =   10800
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "0"
      Height          =   255
      Left            =   8880
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   5040
      Picture         =   "main.frx":030A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3465
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
         Caption         =   "비상문(&E)"
      End
   End
   Begin VB.Menu vogi 
      Caption         =   "보기(&V)"
      Begin VB.Menu shigakwa 
         Caption         =   "시각화(&Z)"
         Checked         =   -1  'True
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
Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
    mplayer.URL = File1.Path & "\" & File1.FileName
    CommandButton8.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
    mplayer.settings.volume = mplayer.settings.volume + 5
    Slider2.Value = Slider2.Value + 5
End Sub

Private Sub Command3_Click()
On Error Resume Next
    mplayer.settings.volume = mplayer.settings.volume - 5
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
    File1.Path = "C:\WINDOWS\MEDIA\"
    Text1.Text = "C:\WINDOWS\MEDIA\"
    Slider1.Value = 0
    'status.Caption = "0"
    mplayer.settings.setMode "loop", True
    mplayer.settings.mute = False
    mutev.Caption = "1"
    ToggleButton1.Caption = "◀))"
    Dir1.Path = "C:\WINDOWS\MEDIA\"
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 2 Then
Me.WindowState = 0
End If
    Me.Width = 10500
    Me.Height = 7125
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
mplayer.fullScreen = True
End Sub

Private Sub Label9_Click()
On Error Resume Next
Slider2.Value = 100
mplayer.settings.volume = 100
End Sub

Private Sub log_Click()
    ulog.Show
End Sub

Private Sub mplayer_MediaChange(ByVal Item As Object)
    On Error Resume Next
    'status.Caption = "1"
    StatusBar1.SimpleText = "재생 중 - " & mplayer.currentMedia.Name
    Timer1.Enabled = True
    Label3.Caption = mplayer.currentMedia.duration
    Slider1.Max = mplayer.currentMedia.duration
    Label2.Caption = mplayer.currentMedia.duration
    'Do While mplayer.Controls.currentPosition < mplayer.currentMedia.duration And status.Caption = "1"
        'Slider1.Value = mplayer.Controls.currentPosition
        'Label1.Caption = mplayer.Controls.currentPosition
    'Loop
    CommandButton9.Visible = True
    CommandButton8.Visible = False
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
    mplayer.settings.volume = Slider2.Value
End Sub

Private Sub Slider2_Scroll()
On Error Resume Next
    mplayer.settings.volume = Slider2.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Label1.Caption = mplayer.Controls.currentPosition
    Slider1.Value = mplayer.Controls.currentPosition
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

