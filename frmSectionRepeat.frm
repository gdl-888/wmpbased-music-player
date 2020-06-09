VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmSectionRepeat 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "구간 반복 설정"
   ClientHeight    =   1110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4725
   Icon            =   "frmSectionRepeat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFinishTIme 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   360
      Width           =   825
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   270
      Left            =   1201
      TabIndex        =   4
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtStartTime"
      BuddyDispid     =   196612
      OrigLeft        =   1440
      OrigTop         =   360
      OrigRight       =   1695
      OrigBottom      =   615
      Increment       =   5
      Max             =   2147483647
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtStartTime 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "취소"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin ComCtl2.UpDown UpDown2 
      Height          =   270
      Left            =   2626
      TabIndex        =   7
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtFinishTIme"
      BuddyDispid     =   196615
      OrigLeft        =   1440
      OrigTop         =   360
      OrigRight       =   1695
      OrigBottom      =   615
      Increment       =   5
      Max             =   2147483647
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "종료 시간(초):"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "시작 시간(초):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSectionRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Form1.SRS = txtStartTime.Text
    Form1.SRE = txtFinishTIme.Text
    Form1.timSR.Enabled = True
    
    Unload Me
End Sub
