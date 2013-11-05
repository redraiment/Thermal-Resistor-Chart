Attribute VB_Name = "Config"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Integer, _
    ByVal lpFileName As String _
) As Integer
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String _
) As Integer

Private ConfigFile As String

Public Type Config
    TempMin As Long
    TempMax As Long
    TempStep As Long
    
    ResiMin As Long
    ResiMax As Long
    ResiStep As Long
    
    CommPort As Long
    CommInterval As Long
End Type

Public AppCfg As Config

Public Sub InitConfig()

    ConfigFile = App.Path & "\config.ini"

End Sub

Public Sub LoadConfig()

    Dim value As String * 32

    GetPrivateProfileString "temp", "min", "0", value, 32, ConfigFile
    AppCfg.TempMin = CLng(value)

    GetPrivateProfileString "temp", "max", "120", value, 32, ConfigFile
    AppCfg.TempMax = CLng(value)

    GetPrivateProfileString "temp", "step", "10", value, 32, ConfigFile
    AppCfg.TempStep = CLng(value)


    GetPrivateProfileString "resi", "min", "0", value, 32, ConfigFile
    AppCfg.ResiMin = CLng(value)

    GetPrivateProfileString "resi", "max", "1500", value, 32, ConfigFile
    AppCfg.ResiMax = CLng(value)

    GetPrivateProfileString "resi", "step", "100", value, 32, ConfigFile
    AppCfg.ResiStep = CLng(value)


    GetPrivateProfileString "comm", "port", "1", value, 32, ConfigFile
    AppCfg.CommPort = CLng(value)

    GetPrivateProfileString "comm", "interval", "1000", value, 32, ConfigFile
    AppCfg.CommInterval = CLng(value)

End Sub

Public Sub SaveConfig()

    WritePrivateProfileString "temp", "min", AppCfg.TempMin, ConfigFile
    WritePrivateProfileString "temp", "max", AppCfg.TempMax, ConfigFile
    WritePrivateProfileString "temp", "step", AppCfg.TempStep, ConfigFile
    
    WritePrivateProfileString "resi", "min", AppCfg.ResiMin, ConfigFile
    WritePrivateProfileString "resi", "max", AppCfg.ResiMax, ConfigFile
    WritePrivateProfileString "resi", "step", AppCfg.ResiStep, ConfigFile

    WritePrivateProfileString "comm", "port", AppCfg.CommPort, ConfigFile
    WritePrivateProfileString "comm", "interval", AppCfg.CommInterval, ConfigFile

End Sub

Public Sub InitConfigFrm()

    ConfigFrm.TempMin.Text = AppCfg.TempMin
    ConfigFrm.TempMinMsg.Caption = "要求为介于-40到140之间的整数"
    ConfigFrm.TempMinMsg.ForeColor = vbBlack
    
    ConfigFrm.TempMax.Text = AppCfg.TempMax
    ConfigFrm.TempMaxMsg.Caption = "要求为介于-40到140之间的整数"
    ConfigFrm.TempMaxMsg.ForeColor = vbBlack
    
    ConfigFrm.TempStep.Text = AppCfg.TempStep
    ConfigFrm.TempStepMsg.Caption = "要求为一个正整数"
    ConfigFrm.TempStepMsg.ForeColor = vbBlack

    ConfigFrm.ResiMin.Text = AppCfg.ResiMin
    ConfigFrm.ResiMinMsg.Caption = "要求为一个自然数"
    ConfigFrm.ResiMinMsg.ForeColor = vbBlack
    
    ConfigFrm.ResiMax.Text = AppCfg.ResiMax
    ConfigFrm.ResiMaxMsg.Caption = "要求为一个正整数"
    ConfigFrm.ResiMaxMsg.ForeColor = vbBlack
    
    ConfigFrm.ResiStep.Text = AppCfg.ResiStep
    ConfigFrm.ResiStepMsg.Caption = "要求为一个正整数"
    ConfigFrm.ResiStepMsg.ForeColor = vbBlack

    ConfigFrm.CommPort.Text = AppCfg.CommPort
    ConfigFrm.CommPortMsg.Caption = "要求为介于1到16之间的正整数"
    ConfigFrm.CommPortMsg.ForeColor = vbBlack

    ConfigFrm.CommInterval.Text = AppCfg.CommInterval
    ConfigFrm.CommIntervalMsg.Caption = "要求为不小于1000的正整数，单位毫秒"
    ConfigFrm.CommIntervalMsg.ForeColor = vbBlack

    ConfigFrm.TabPanel(0).Visible = True
    ConfigFrm.TabPanel(1).Visible = False

End Sub

