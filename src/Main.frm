VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainFrm 
   Caption         =   "温度电阻检测仪 - 温州智润机电有限公司"
   ClientHeight    =   7290
   ClientLeft      =   1620
   ClientTop       =   1740
   ClientWidth     =   10680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10680
   Begin MSComctlLib.ImageList ToolbarIcon 
      Left            =   8520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":173C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1DD9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm Phone 
      Left            =   7560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
      InBufferSize    =   256
      OutBufferSize   =   256
      RThreshold      =   6
      RTSEnable       =   -1  'True
      BaudRate        =   115200
      InputMode       =   1
   End
   Begin VB.Timer Interval 
      Interval        =   1500
      Left            =   6720
      Top             =   120
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ToolbarIcon"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComDlg.CommonDialog ComDlg 
         Left            =   9500
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.PictureBox PBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6300
      Left            =   0
      ScaleHeight     =   6300
      ScaleWidth      =   10680
      TabIndex        =   1
      Top             =   600
      Width           =   10680
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6915
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MenuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileOpen 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuFileSaveAs 
         Caption         =   "另存(&A)"
      End
      Begin VB.Menu MenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileExport 
         Caption         =   "导出(&E)"
         Shortcut        =   ^P
      End
      Begin VB.Menu MenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileQuit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuSet 
      Caption         =   "设置(&S)"
      Begin VB.Menu MenuSetOption 
         Caption         =   "选项(&O)"
      End
   End
   Begin VB.Menu MenuRecord 
      Caption         =   "录制(&R)"
      Begin VB.Menu MenuRecordToggle 
         Caption         =   "开始(&S)"
      End
      Begin VB.Menu MenuRecordStop 
         Caption         =   "停止(&E)"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "关于(&A)"
      End
      Begin VB.Menu MenuHelpManual 
         Caption         =   "手册(&M)"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private box As XBox
Private fileName As String

' == 启动项 ==
' 加载配置文件
Private Sub Form_Load()

    fileName = ""

    InitConfig
    LoadConfig

    Set box = New XBox
    box.Init PBox

End Sub

' == 功能项 ==

' 定时发送消息
Private Sub Interval_Timer()

    If Phone.PortOpen Then Phone.Output = "Z"

End Sub

' 定时收取消息
Private Sub Phone_OnComm()

    Dim buffer() As Byte
    Dim msg As CommMessage

    If Phone.CommEvent = comEvReceive Then
        buffer = Phone.Input
        msg = ParseMessage(buffer)
        If msg.Resistance <> INVAILD_DATA Then
            box.PutPoint msg.Temperature, msg.Resistance
        End If
    End If

End Sub

' == 菜单项 ==
' 提供软件开放的所有功能

Private Sub MenuFileNew_Click()

    box.Reset
    fileName = ""

End Sub

Private Sub MenuFileOpen_Click()
    
    ComDlg.DialogTitle = "打开文件"
    ComDlg.Filter = "Text File(*.txt)|*.txt"
    ComDlg.fileName = ""
    ComDlg.ShowOpen

    If ComDlg.fileName <> "" Then
        fileName = ComDlg.fileName
        box.Load fileName
    End If

End Sub

Private Sub MenuFileSave_Click()

    If fileName = "" Then
        ComDlg.DialogTitle = "保存文件"
        ComDlg.Filter = "Text File(*.txt)|*.txt"
        ComDlg.fileName = ""
        ComDlg.ShowSave
        If ComDlg.fileName <> "" Then
            fileName = ComDlg.fileName
        End If
    End If

    If fileName <> "" Then
        box.Save fileName
    End If

End Sub

' = 文件项 =
Private Sub MenuFileQuit_Click()

    If Phone.PortOpen Then Phone.PortOpen = False
    End

End Sub

' = 设置项 =
Private Sub MenuSetOption_Click()

    ConfigFrm.Left = MainFrm.Left + (MainFrm.Width - ConfigFrm.Width) \ 3
    ConfigFrm.Top = MainFrm.Top + (MainFrm.Height - ConfigFrm.Height) \ 3
    InitConfigFrm
    ConfigFrm.TabBox.SelectedItem = ConfigFrm.TabBox.Tabs(1)
    ConfigFrm.Show

End Sub

Private Sub RecordStart()

    MenuRecordToggle.Caption = "暂停"
    Toolbar.Buttons(5).Image = 5

    If Not Phone.PortOpen Then
        Phone.CommPort = AppCfg.CommPort
        Phone.PortOpen = True
    End If

End Sub

Private Sub RecordStop()

    MenuRecordToggle.Caption = "开始"
    Toolbar.Buttons(5).Image = 4
    If Phone.PortOpen Then
        Phone.PortOpen = False
    End If

End Sub

' = 录制项 =
Private Sub MenuRecordToggle_Click()

    If Toolbar.Buttons(5).Image = 5 Then
        RecordStop
    Else
        RecordStart
    End If

End Sub

Private Sub MenuRecordStop_Click()

    RecordStop

End Sub

' = 帮助项 =
Private Sub MenuHelpAbout_Click()

    AboutFrm.Left = MainFrm.Left + (MainFrm.Width - AboutFrm.Width) \ 3
    AboutFrm.Top = MainFrm.Top + (MainFrm.Height - AboutFrm.Height) \ 3
    AboutFrm.Show

End Sub

Private Sub MenuHelpManual_Click()

    MsgBox "Not support yet"

End Sub

' == 界面项 ==
' 基本为调用菜单项来完成任务

Private Sub Form_Unload(Cancel As Integer)

    MenuFileQuit_Click

End Sub

' = 工具栏 =
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 1 Then
        MenuFileNew_Click
    ElseIf Button.Index = 2 Then
        MenuFileOpen_Click
    ElseIf Button.Index = 3 Then
        MenuFileSave_Click
    ElseIf Button.Index = 5 Then
        MenuRecordToggle_Click
    ElseIf Button.Index = 6 Then
        MenuRecordStop_Click
    End If

End Sub


' 当选项窗口开启时，不能切换到主窗口
Private Sub Form_Activate()

    If ConfigFrm.Visible Then
        ConfigFrm.SetFocus
    ElseIf AboutFrm.Visible Then
        AboutFrm.SetFocus
    End If

End Sub
