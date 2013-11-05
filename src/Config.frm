VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ConfigFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选项"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame TabPanel 
      BorderStyle     =   0  'None
      Height          =   3400
      Index           =   1
      Left            =   6400
      TabIndex        =   5
      Top             =   350
      Width           =   6290
      Begin VB.Frame CommFrame 
         Caption         =   "串口选项"
         Height          =   3100
         Left            =   200
         TabIndex        =   26
         Top             =   200
         Width           =   5890
         Begin VB.TextBox CommInterval 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            TabIndex        =   32
            Text            =   "1500"
            Top             =   750
            Width           =   800
         End
         Begin VB.TextBox CommPort 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            TabIndex        =   31
            Text            =   "1"
            Top             =   350
            Width           =   800
         End
         Begin VB.Label CommIntervalMsg 
            Caption         =   "要求为不小于1000的正整数，单位毫秒"
            Height          =   200
            Left            =   2200
            TabIndex        =   30
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label CommIntervalLabel 
            Caption         =   "采样间隔："
            Height          =   200
            Left            =   200
            TabIndex        =   29
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label CommPortMsg 
            Caption         =   "要求为介于1到16之间的正整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   28
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "采样端口："
            Height          =   200
            Left            =   200
            TabIndex        =   27
            Top             =   350
            Width           =   1000
         End
      End
   End
   Begin VB.Frame TabPanel 
      BorderStyle     =   0  'None
      Height          =   3400
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   350
      Width           =   6290
      Begin VB.Frame ResistanceRangeFrame 
         Caption         =   "电阻范围"
         Height          =   1500
         Left            =   200
         TabIndex        =   16
         Top             =   1800
         Width           =   5890
         Begin VB.TextBox ResiMin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            TabIndex        =   19
            Text            =   "0"
            Top             =   350
            Width           =   800
         End
         Begin VB.TextBox ResiMax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            TabIndex        =   18
            Text            =   "1500"
            Top             =   750
            Width           =   800
         End
         Begin VB.TextBox ResiStep 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            TabIndex        =   17
            Text            =   "100"
            Top             =   1150
            Width           =   800
         End
         Begin VB.Label ResiMinLabel 
            Caption         =   "最小阻值："
            Height          =   200
            Left            =   200
            TabIndex        =   25
            Top             =   350
            Width           =   1000
         End
         Begin VB.Label ResiMinMsg 
            Caption         =   "要求为一个自然数"
            Height          =   200
            Left            =   2200
            TabIndex        =   24
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label ResiMaxLabel 
            Caption         =   "最大阻值："
            Height          =   200
            Left            =   200
            TabIndex        =   23
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label ResiMaxMsg 
            Caption         =   "要求为一个正整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   22
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label ResiStepLabel 
            Caption         =   "阻值间隔："
            Height          =   195
            Left            =   200
            TabIndex        =   21
            Top             =   1155
            Width           =   1005
         End
         Begin VB.Label ResiStepMsg 
            Caption         =   "要求为一个正整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   20
            Top             =   1150
            Width           =   3290
         End
      End
      Begin VB.Frame TemperatureRangeFrame 
         Caption         =   "温度范围"
         Height          =   1500
         Left            =   200
         TabIndex        =   6
         Top             =   200
         Width           =   5890
         Begin VB.TextBox TempStep 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "10"
            Top             =   1150
            Width           =   800
         End
         Begin VB.TextBox TempMax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   14
            Text            =   "120"
            Top             =   750
            Width           =   800
         End
         Begin VB.TextBox TempMin 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   8
            Text            =   "0"
            Top             =   350
            Width           =   800
         End
         Begin VB.Label TempStepMsg 
            Caption         =   "要求为一个正整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   13
            Top             =   1150
            Width           =   3290
         End
         Begin VB.Label TempStepLabel 
            Caption         =   "温度间隔："
            Height          =   200
            Left            =   200
            TabIndex        =   12
            Top             =   1150
            Width           =   1000
         End
         Begin VB.Label TempMaxMsg 
            Caption         =   "要求为介于-40到140之间的整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   11
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label TempMaxLabel 
            Caption         =   "最大温度："
            Height          =   200
            Left            =   200
            TabIndex        =   10
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label TempMinMsg 
            Caption         =   "要求为介于-40到140之间的整数"
            Height          =   200
            Left            =   2200
            TabIndex        =   9
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label TempMinLabel 
            Caption         =   "最小温度："
            Height          =   200
            Left            =   200
            TabIndex        =   7
            Top             =   350
            Width           =   1000
         End
      End
   End
   Begin VB.CommandButton OptionApplyBtn 
      Caption         =   "应用"
      Height          =   350
      Left            =   4915
      TabIndex        =   3
      Top             =   4000
      Width           =   1200
   End
   Begin VB.CommandButton OptionCancelBtn 
      Caption         =   "取消"
      Height          =   350
      Left            =   3515
      TabIndex        =   2
      Top             =   4000
      Width           =   1200
   End
   Begin VB.CommandButton OptionSubmitBtn 
      Caption         =   "确定"
      Height          =   350
      Left            =   2115
      TabIndex        =   1
      Top             =   4000
      Width           =   1200
   End
   Begin MSComctlLib.TabStrip TabBox 
      Height          =   3900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6879
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "通用"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "高级"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ConfigFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const MinTemperature As Integer = -40
Private Const MaxTemperature As Integer = 140

' === 通用卡 ===

' == 校验项 ==

' 最小温度
Private Function CheckTempMin() As Boolean

    Dim i As Integer, from As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempMin.Text
    CheckTempMin = True
    
    length = Len(value)
    If length = 0 Then
        TempMinMsg.Caption = "最小温度不能为空"
        CheckTempMin = False
    Else
        from = IIf(Mid(value, 1, 1) = "-", 2, 1)
        For i = from To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Or length < from Then
            TempMinMsg.Caption = "非法符号，请输入-40到140之间的整数"
            CheckTempMin = False
        Else
            i = CInt(value)
            
            If i < MinTemperature Or MaxTemperature < i Then
                TempMinMsg.Caption = "范围有误，请输入-40到140之间的整数"
                CheckTempMin = False
            End If
        End If
    End If

    If CheckTempMin Then
        AppCfg.TempMin = CLng(TempMin.Text)
        TempMinMsg.Caption = ""
    Else
        TempMinMsg.ForeColor = vbRed
    End If

End Function

' 最大温度
Private Function CheckTempMax() As Boolean

    Dim i As Integer, from As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempMax.Text
    CheckTempMax = True
    
    length = Len(value)
    If length = 0 Then
        TempMaxMsg.Caption = "最大温度不能为空"
        CheckTempMax = False
    Else
        from = IIf(Mid(value, 1, 1) = "-", 2, 1)
        For i = from To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Or length < from Then
            TempMaxMsg.Caption = "非法符号，请输入-40到140之间的整数"
            CheckTempMax = False
        Else
            i = CInt(value)
            
            If i < MinTemperature Or MaxTemperature < i Then
                TempMaxMsg.Caption = "范围有误，请输入-40到140之间的整数"
                CheckTempMax = False
            ElseIf i < CInt(TempMin.Text) Then
                TempMaxMsg.Caption = "最大温度不可小于最小温度"
                CheckTempMax = False
            End If
        End If
    End If

    If CheckTempMax Then
        AppCfg.TempMax = CLng(TempMax.Text)
        TempMaxMsg.Caption = ""
    Else
        TempMaxMsg.ForeColor = vbRed
    End If

End Function

' 温度间隔
Private Function CheckTempStep() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempStep.Text
    CheckTempStep = True
    
    length = Len(value)
    If length = 0 Then
        TempStepMsg.Caption = "温度间隔不能为空"
        CheckTempStep = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            TempStepMsg.Caption = "非法符号，请输入一个正整数"
            CheckTempStep = False
        End If
    End If

    If CheckTempStep Then
        AppCfg.TempStep = CLng(TempStep.Text)
        TempStepMsg.Caption = ""
    Else
        TempStepMsg.ForeColor = vbRed
    End If

End Function

' 最小电阻
Private Function CheckResiMin() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiMin.Text
    CheckResiMin = True
    
    length = Len(value)
    If length = 0 Then
        ResiMinMsg.Caption = "最小电阻不能为空"
        CheckResiMin = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiMinMsg.Caption = "非法符号，请输入一个自然数"
            CheckResiMin = False
        ElseIf CLng(value) < 0 Then
            ResiMinMsg.Caption = "请输入一个自然数"
            CheckResiMin = False
        End If
    End If

    If CheckResiMin Then
        AppCfg.ResiMin = CLng(ResiMin.Text)
        ResiMinMsg.Caption = ""
    Else
        ResiMinMsg.ForeColor = vbRed
    End If

End Function

' 最大电阻
Private Function CheckResiMax() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiMax.Text
    CheckResiMax = True
    
    length = Len(value)
    If length = 0 Then
        ResiMaxMsg.Caption = "最大电阻不能为空"
        CheckResiMax = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiMaxMsg.Caption = "非法符号，请输入一个正整数"
            CheckResiMax = False
        ElseIf CLng(value) < CLng(ResiMin.Text) Then
            ResiMaxMsg.Caption = "最大电阻不能小于最小电阻"
            CheckResiMax = False
        End If
    End If

    If CheckResiMax Then
        AppCfg.ResiMax = CLng(ResiMax.Text)
        ResiMaxMsg.Caption = ""
    Else
        ResiMaxMsg.ForeColor = vbRed
    End If

End Function

' 阻值间隔
Private Function CheckResiStep() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiStep.Text
    CheckResiStep = True
    
    length = Len(value)
    If length = 0 Then
        ResiStepMsg.Caption = "阻值间隔不能为空"
        CheckResiStep = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiStepMsg.Caption = "非法符号，请输入一个正整数"
            CheckResiStep = False
        ElseIf CLng(value) < 1 Then
            ResiStepMsg.Caption = "请输入一个正整数"
            CheckResiStep = False
        End If
    End If

    If CheckResiStep Then
        AppCfg.ResiStep = CLng(ResiStep.Text)
        ResiStepMsg.Caption = ""
    Else
        ResiStepMsg.ForeColor = vbRed
    End If

End Function

' === 高级卡 ===

' 采样端口
Private Function CheckCommPort() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = CommPort.Text
    CheckCommPort = True
    
    length = Len(value)
    If length = 0 Then
        CommPortMsg.Caption = "采样端口不能为空"
        CheckCommPort = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            CommPortMsg.Caption = "非法符号，请输入1-16之间的一个正整数"
            CheckCommPort = False
        ElseIf CLng(value) < 1 Or CLng(value) > 16 Then
            CommPortMsg.Caption = "请输入一个1-16之间的正整数"
            CheckCommPort = False
        End If
    End If

    If CheckCommPort Then
        AppCfg.CommPort = CLng(CommPort.Text)
        CommPortMsg.Caption = ""
    Else
        CommPortMsg.ForeColor = vbRed
    End If

End Function

' 采样间隔
Private Function CheckCommInterval() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = CommInterval.Text
    CheckCommInterval = True
    
    length = Len(value)
    If length = 0 Then
        CommIntervalMsg.Caption = "采样间隔不能为空"
        CheckCommInterval = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            CommIntervalMsg.Caption = "非法符号，请输入一个不小于1000的正整数"
            CheckCommInterval = False
        ElseIf CLng(value) < 1000 Then
            CommIntervalMsg.Caption = "请输入一个不小于1000的正整数"
            CheckCommInterval = False
        End If
    End If

    If CheckCommInterval Then
        AppCfg.CommInterval = CLng(CommInterval.Text)
        CommIntervalMsg.Caption = ""
    Else
        CommIntervalMsg.ForeColor = vbRed
    End If

End Function

' === 共用卡 ===

' 汇总的校验函数
Function CheckAll() As Boolean

    CheckAll = CheckTempMin _
           And CheckTempMax _
           And CheckTempStep _
           And CheckResiMin _
           And CheckResiMax _
           And CheckResiStep _
           And CheckCommPort _
           And CheckCommInterval

End Function

' == 启动项 ==

Private Sub Form_Load()

    Dim i As Integer
    
    For i = 0 To 1
        TabPanel(i).Visible = False
        TabPanel(i).Left = 0
    Next

End Sub

' == 界面项 ==

Private Sub OptionApplyBtn_Click()

    If CheckAll Then
        SaveConfig
    End If

End Sub

Private Sub OptionCancelBtn_Click()

    Me.Hide

End Sub

Private Sub OptionSubmitBtn_Click()

    If CheckAll Then
        SaveConfig
        Me.Hide
    End If

End Sub

Private Sub TabBox_Click()

    TabPanel(TabBox.SelectedItem.Index - 1).Visible = True
    TabPanel(2 - TabBox.SelectedItem.Index).Visible = False

End Sub
