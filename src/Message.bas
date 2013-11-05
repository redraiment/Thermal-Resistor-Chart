Attribute VB_Name = "Message"
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Const INVAILD_DATA As Double = -100

Public Type CommMessage
    Temperature As Double
    Resistance As Double
End Type

Public Function ParseTemperature(bHigh As Byte, bLow As Byte) As Double

    ParseTemperature = (bHigh And &H7) * 16 + bLow / 16
    If (bHigh And &H8) <> 0 Then
        ParseTemperature = ParseTemperature * -1
    End If
    
    ParseTemperature = Round(ParseTemperature, 1)

End Function

Public Function ParseVoltage(bHigh As Byte, bLow As Byte) As Double

    ParseVoltage = (bHigh And &H7) * 256 + bLow

End Function

Public Function ParseResistance(dRealVoltage As Double, dTotalVoltage) As Double

    Const ReferenceResistor As Double = 676
    
    If dRealVoltage < dTotalVoltage Then
        ParseResistance = dRealVoltage * ReferenceResistor / (dTotalVoltage - dRealVoltage)
        ParseResistance = Round(ParseResistance, 1)
    Else
        ParseResistance = INVAILD_DATA
    End If

End Function

Public Function ParseMessage(data() As Byte) As CommMessage

    ParseMessage.Temperature = INVAILD_DATA
    ParseMessage.Resistance = INVAILD_DATA

    If SafeArrayGetDim(data) < 6 Then Exit Function

    Dim offset As Integer
    offset = UBound(data) - 5
    
    ParseMessage.Temperature = ParseTemperature(data(offset + 0), data(offset + 1))
    ParseMessage.Resistance = ParseResistance(ParseVoltage(data(offset + 2), data(offset + 3)), _
                                              ParseVoltage(data(offset + 4), data(offset + 5)))

End Function
