VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ConfigFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ��"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame TabPanel 
      BorderStyle     =   0  'None
      Height          =   3400
      Index           =   1
      Left            =   6400
      TabIndex        =   5
      Top             =   350
      Width           =   6290
      Begin VB.Frame CommFrame 
         Caption         =   "����ѡ��"
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
            Caption         =   "Ҫ��Ϊ��С��1000������������λ����"
            Height          =   200
            Left            =   2200
            TabIndex        =   30
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label CommIntervalLabel 
            Caption         =   "���������"
            Height          =   200
            Left            =   200
            TabIndex        =   29
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label CommPortMsg 
            Caption         =   "Ҫ��Ϊ����1��16֮���������"
            Height          =   200
            Left            =   2200
            TabIndex        =   28
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "�����˿ڣ�"
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
         Caption         =   "���跶Χ"
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
            Caption         =   "��С��ֵ��"
            Height          =   200
            Left            =   200
            TabIndex        =   25
            Top             =   350
            Width           =   1000
         End
         Begin VB.Label ResiMinMsg 
            Caption         =   "Ҫ��Ϊһ����Ȼ��"
            Height          =   200
            Left            =   2200
            TabIndex        =   24
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label ResiMaxLabel 
            Caption         =   "�����ֵ��"
            Height          =   200
            Left            =   200
            TabIndex        =   23
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label ResiMaxMsg 
            Caption         =   "Ҫ��Ϊһ��������"
            Height          =   200
            Left            =   2200
            TabIndex        =   22
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label ResiStepLabel 
            Caption         =   "��ֵ�����"
            Height          =   195
            Left            =   200
            TabIndex        =   21
            Top             =   1155
            Width           =   1005
         End
         Begin VB.Label ResiStepMsg 
            Caption         =   "Ҫ��Ϊһ��������"
            Height          =   200
            Left            =   2200
            TabIndex        =   20
            Top             =   1150
            Width           =   3290
         End
      End
      Begin VB.Frame TemperatureRangeFrame 
         Caption         =   "�¶ȷ�Χ"
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
            Caption         =   "Ҫ��Ϊһ��������"
            Height          =   200
            Left            =   2200
            TabIndex        =   13
            Top             =   1150
            Width           =   3290
         End
         Begin VB.Label TempStepLabel 
            Caption         =   "�¶ȼ����"
            Height          =   200
            Left            =   200
            TabIndex        =   12
            Top             =   1150
            Width           =   1000
         End
         Begin VB.Label TempMaxMsg 
            Caption         =   "Ҫ��Ϊ����-40��140֮�������"
            Height          =   200
            Left            =   2200
            TabIndex        =   11
            Top             =   750
            Width           =   3290
         End
         Begin VB.Label TempMaxLabel 
            Caption         =   "����¶ȣ�"
            Height          =   200
            Left            =   200
            TabIndex        =   10
            Top             =   750
            Width           =   1000
         End
         Begin VB.Label TempMinMsg 
            Caption         =   "Ҫ��Ϊ����-40��140֮�������"
            Height          =   200
            Left            =   2200
            TabIndex        =   9
            Top             =   350
            Width           =   3290
         End
         Begin VB.Label TempMinLabel 
            Caption         =   "��С�¶ȣ�"
            Height          =   200
            Left            =   200
            TabIndex        =   7
            Top             =   350
            Width           =   1000
         End
      End
   End
   Begin VB.CommandButton OptionApplyBtn 
      Caption         =   "Ӧ��"
      Height          =   350
      Left            =   4915
      TabIndex        =   3
      Top             =   4000
      Width           =   1200
   End
   Begin VB.CommandButton OptionCancelBtn 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   3515
      TabIndex        =   2
      Top             =   4000
      Width           =   1200
   End
   Begin VB.CommandButton OptionSubmitBtn 
      Caption         =   "ȷ��"
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
            Caption         =   "ͨ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�߼�"
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

' === ͨ�ÿ� ===

' == У���� ==

' ��С�¶�
Private Function CheckTempMin() As Boolean

    Dim i As Integer, from As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempMin.Text
    CheckTempMin = True
    
    length = Len(value)
    If length = 0 Then
        TempMinMsg.Caption = "��С�¶Ȳ���Ϊ��"
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
            TempMinMsg.Caption = "�Ƿ����ţ�������-40��140֮�������"
            CheckTempMin = False
        Else
            i = CInt(value)
            
            If i < MinTemperature Or MaxTemperature < i Then
                TempMinMsg.Caption = "��Χ����������-40��140֮�������"
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

' ����¶�
Private Function CheckTempMax() As Boolean

    Dim i As Integer, from As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempMax.Text
    CheckTempMax = True
    
    length = Len(value)
    If length = 0 Then
        TempMaxMsg.Caption = "����¶Ȳ���Ϊ��"
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
            TempMaxMsg.Caption = "�Ƿ����ţ�������-40��140֮�������"
            CheckTempMax = False
        Else
            i = CInt(value)
            
            If i < MinTemperature Or MaxTemperature < i Then
                TempMaxMsg.Caption = "��Χ����������-40��140֮�������"
                CheckTempMax = False
            ElseIf i < CInt(TempMin.Text) Then
                TempMaxMsg.Caption = "����¶Ȳ���С����С�¶�"
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

' �¶ȼ��
Private Function CheckTempStep() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = TempStep.Text
    CheckTempStep = True
    
    length = Len(value)
    If length = 0 Then
        TempStepMsg.Caption = "�¶ȼ������Ϊ��"
        CheckTempStep = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            TempStepMsg.Caption = "�Ƿ����ţ�������һ��������"
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

' ��С����
Private Function CheckResiMin() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiMin.Text
    CheckResiMin = True
    
    length = Len(value)
    If length = 0 Then
        ResiMinMsg.Caption = "��С���費��Ϊ��"
        CheckResiMin = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiMinMsg.Caption = "�Ƿ����ţ�������һ����Ȼ��"
            CheckResiMin = False
        ElseIf CLng(value) < 0 Then
            ResiMinMsg.Caption = "������һ����Ȼ��"
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

' ������
Private Function CheckResiMax() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiMax.Text
    CheckResiMax = True
    
    length = Len(value)
    If length = 0 Then
        ResiMaxMsg.Caption = "�����費��Ϊ��"
        CheckResiMax = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiMaxMsg.Caption = "�Ƿ����ţ�������һ��������"
            CheckResiMax = False
        ElseIf CLng(value) < CLng(ResiMin.Text) Then
            ResiMaxMsg.Caption = "�����費��С����С����"
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

' ��ֵ���
Private Function CheckResiStep() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = ResiStep.Text
    CheckResiStep = True
    
    length = Len(value)
    If length = 0 Then
        ResiStepMsg.Caption = "��ֵ�������Ϊ��"
        CheckResiStep = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            ResiStepMsg.Caption = "�Ƿ����ţ�������һ��������"
            CheckResiStep = False
        ElseIf CLng(value) < 1 Then
            ResiStepMsg.Caption = "������һ��������"
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

' === �߼��� ===

' �����˿�
Private Function CheckCommPort() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = CommPort.Text
    CheckCommPort = True
    
    length = Len(value)
    If length = 0 Then
        CommPortMsg.Caption = "�����˿ڲ���Ϊ��"
        CheckCommPort = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            CommPortMsg.Caption = "�Ƿ����ţ�������1-16֮���һ��������"
            CheckCommPort = False
        ElseIf CLng(value) < 1 Or CLng(value) > 16 Then
            CommPortMsg.Caption = "������һ��1-16֮���������"
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

' �������
Private Function CheckCommInterval() As Boolean

    Dim i As Integer, length As Integer
    Dim value As String, char As String
    
    value = CommInterval.Text
    CheckCommInterval = True
    
    length = Len(value)
    If length = 0 Then
        CommIntervalMsg.Caption = "�����������Ϊ��"
        CheckCommInterval = False
    Else
        For i = 1 To length
            char = Mid(value, i, 1)
            If "9" < char Or char < "0" Then
                Exit For
            End If
        Next
        
        If i <= length Then
            CommIntervalMsg.Caption = "�Ƿ����ţ�������һ����С��1000��������"
            CheckCommInterval = False
        ElseIf CLng(value) < 1000 Then
            CommIntervalMsg.Caption = "������һ����С��1000��������"
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

' === ���ÿ� ===

' ���ܵ�У�麯��
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

' == ������ ==

Private Sub Form_Load()

    Dim i As Integer
    
    For i = 0 To 1
        TabPanel(i).Visible = False
        TabPanel(i).Left = 0
    Next

End Sub

' == ������ ==

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
