VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const PADDING = 400
Private Const POINT_SIZE = 10

Private box As PictureBox

Private xMin As Long, xMax As Long
Private yMin As Long, yMax As Long
Private xUnit As Double
Private yUnit As Double

Private dataSize As Long
Private dataCapacity As Long
Private data() As Double

Public Sub Init(picture As PictureBox)

    Set box = picture
    
    xMin = PADDING * 3 \ 2
    xMax = box.Width - xMin
    
    yMax = PADDING * 3 \ 2
    yMin = box.Height - yMax

    Reset

End Sub

Private Sub DrawXAxis()

    Dim xAxisYOffset As Long, i As Long
    xAxisYOffset = box.Height - PADDING

    box.Line (PADDING, xAxisYOffset)-(box.Width - PADDING, xAxisYOffset)
    box.PSet (box.Width - PADDING, xAxisYOffset - 100)
    box.Print "ÎÂ¶È"

    xUnit = (xMax - xMin) * 1# / (AppCfg.TempMax - AppCfg.TempMin + 1)

    For i = AppCfg.TempMin To AppCfg.TempMax Step AppCfg.TempStep
        box.PSet ((i - AppCfg.TempMin) * xUnit + xMin, xAxisYOffset + 100)
        box.Print i
    Next

End Sub

Private Sub DrawYAxis()

    Dim i As Long

    box.Line (PADDING, PADDING)-(PADDING, box.Height - PADDING)
    box.PSet (PADDING \ 2, PADDING \ 2)
    box.Print "µç×è"
    
    yUnit = (yMin - yMax) * 1# / (AppCfg.ResiMax - AppCfg.ResiMin + 1)

    For i = AppCfg.ResiMin To AppCfg.ResiMax Step AppCfg.ResiStep
        box.PSet (-80, yMin - (i - AppCfg.ResiMin) * yUnit)
        box.Print i
    Next

End Sub

Public Sub PutPoint(x As Double, y As Double)

    If AppCfg.TempMin <= x And x <= AppCfg.TempMax _
   And AppCfg.ResiMin <= y And y <= AppCfg.ResiMax Then
        addPoint x, y
        
        x = (x - AppCfg.TempMin) * xUnit + xMin
        y = yMin - (y - AppCfg.ResiMin) * yUnit
        
        box.Line (x - POINT_SIZE, y)-(x + POINT_SIZE, y), vbRed
        box.Line (x, y - POINT_SIZE)-(x, y + POINT_SIZE), vbRed
    End If

End Sub

Private Sub addPoint(x As Double, y As Double)

    If dataSize >= dataCapacity Then
        dataCapacity = dataCapacity * 2
        ReDim Preserve data(dataCapacity, 1) As Double
    End If

    dataSize = dataSize + 1
    data(dataSize, 0) = x
    data(dataSize, 1) = y

End Sub

Public Sub Reset()

    box.Cls
    box.AutoRedraw = True
    
    DrawXAxis
    DrawYAxis

    dataSize = -1
    dataCapacity = 10000
    ReDim data(dataCapacity, 1) As Double

End Sub

Public Sub Save(fileName As String)

    Dim s As String
    Dim i As Long

    Open fileName For Output As #1
    For i = 0 To dataSize
        s = CStr(data(i, 0)) & " " & CStr(data(i, 1))
        Print #1, s
    Next
    Close #1

End Sub

Public Sub Load(fileName As String)

    Dim point() As String
    Dim x As Double, y As Double
    Dim s As String

    Open fileName For Input As #1
    Do Until EOF(1)
        Line Input #1, s
        point = Split(s)
        x = CDbl(point(0))
        y = CDbl(point(1))
        PutPoint x, y
    Loop
    Close #1

End Sub

Public Sub Export(fileName As String)

    SavePicture box.Image, fileName

End Sub
