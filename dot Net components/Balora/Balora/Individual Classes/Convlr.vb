Option Strict Off
Public Class Convlr
    Shared Function baseN2dec(ByVal value As String, ByVal inBase As Integer) As String
        'http://www.ecelab.com/vb-dec2hex.htm
        'Converts any base to base 10
        Dim strValue As String
        Dim i As Integer
        Dim x As String
        Dim y As Double
        strValue = StrReverse(CStr(UCase(value)))
        For i = 0 To Len(strValue) - 1
            x = Mid(strValue, i + 1, 1)
            If Not IsNumeric(x) Then
                y = y + ((Asc(x) - 65) + 10) * (inBase ^ i)
            Else
                y = y + ((inBase ^ i) * CInt(x))
            End If
        Next
        baseN2dec = y
    End Function

    Shared Function dec2baseN(ByVal value As String, ByVal outBase As Integer) As String
        'http://www.ecelab.com/vb-dec2hex.htm
        'Converts base 10 to any base
        Dim q As Double  'quotient
        Dim r As Object   'remainder
        Dim m As Integer   'denominator
        Dim y As String = vbNullString   'converted value
        m = outBase
        q = value
        Do
            r = q Mod m
            q = Int(q / m)
            If r >= 10 Then
                r = Chr(65 + (r - 10))
            End If
            y = y & CStr(r)
        Loop Until q = 0
        dec2baseN = StrReverse(y)
    End Function
End Class
