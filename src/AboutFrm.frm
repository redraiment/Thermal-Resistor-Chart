VERSION 5.00
Begin VB.Form AboutFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 温度电阻检测仪"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "AboutFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3930
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton AboutConfirm 
      Caption         =   "确定"
      Height          =   350
      Left            =   1500
      TabIndex        =   1
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label AboutMsg 
      Height          =   600
      Left            =   1000
      TabIndex        =   0
      Top             =   300
      Width           =   2500
   End
   Begin VB.Image AboutIcon 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   300
      Picture         =   "AboutFrm.frx":08CA
      Top             =   300
      Width           =   480
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AboutConfirm_Click()

    Me.Hide

End Sub

Private Sub Form_Load()

    AboutMsg.Caption = "温度电阻检测是 版本 1.0" & vbCrLf _
                     & "版权所有(C) 2013-" & Year(Now) & "," & vbCrLf _
                     & "redraiment.com"

End Sub
