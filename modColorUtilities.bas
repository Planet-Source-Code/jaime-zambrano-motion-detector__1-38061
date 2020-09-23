Attribute VB_Name = "modColorUtilities"
Option Explicit
Public Type ParPromDest
    StdVar As Double
    Promedio As Double
End Type
Public Type ColorComponent
    RGBColor As Long
    RedInt As Integer
    GreenInt As Integer
    BlueInt As Integer
    SumaColInt As Integer

End Type
Public Sub GetColores(Canvas As Object, x As Single, Y As Single, _
                    Colors As ColorComponent)
    'On Error GoTo errorHandler
    Dim RedHex As String
    Dim GreenHex As String
    Dim BlueHex As String
    Dim RGBColorHex As String
    Dim C As ColorComponent

    C = Colors
    C.RGBColor = Canvas.Point(x, Y)

        'C.RGBColor = GetPixel(Canvas.hdc, x, y) ' probar
    'Debug.Print x, Y, C.RGBColor
    ' Use Format() in the case of
    ' the special value zero (0)
    RGBColorHex = CStr(Format(Hex(C.RGBColor), "000000"))
    ' Extract the component values


If Len(RGBColorHex) = 6 Then
    RedHex = "&H" & Right(RGBColorHex, 2)
    GreenHex = "&H" & Mid(RGBColorHex, 3, 2)
    BlueHex = "&H" & Left(RGBColorHex, 2)
Else
    If Len(RGBColorHex) = 4 Then
        RedHex = "&H" & Right(RGBColorHex, 2)
        GreenHex = "&H" & Left(RGBColorHex, 2)
        BlueHex = "&H00"
    Else
        If Len(RGBColorHex) = 2 Then
            RedHex = "&H" & Right(RGBColorHex, 2)
            GreenHex = "&H00"
            BlueHex = "&H00"
        Else
            RedHex = "&H00"
            GreenHex = "&H00"
            BlueHex = "&H00"
        End If

    End If
End If

    ' Convert the string to ints
    C.SumaColInt = CInt(RedHex) + CInt(GreenHex) + CInt(BlueHex)
    'C.RedInt = CInt(RedHex)
    'C.GreenInt = CInt(GreenHex)
    'C.BlueInt = CInt(BlueHex)


    Colors = C

    Exit Sub
errorHandler:
    MsgBox "Error en frmMain.GetColores ; " & Err.Number & vbCrLf & _
           Err.Description
End Sub


