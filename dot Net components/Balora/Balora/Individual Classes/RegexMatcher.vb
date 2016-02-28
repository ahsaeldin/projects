Imports System.Text.RegularExpressions

Public Class RegexMatcher
    Inherits BaloraBase

    ''' <summary>
    ''' Matches the alphanumeric without underscore.
    ''' </summary>
    ''' <param name="str">The name.</param><returns></returns>
    Public Function MatchAlphanumeric(str As String) As Boolean
        Return Matcher(str, "^[0-9a-zA-Z ]+$")
    End Function

    Public Function MatchURL(str As String) As Boolean
        Dim checkWithoutWWW As Boolean = Matcher(str, "^([a-zA-Z0-9]+(\.[a-zA-Z0-9]+)+.*)$")
        Dim checkWithoutHttp As Boolean = Matcher(str, "^((https?|ftp)://|(www|ftp)\.)[a-z0-9-]+(\.[a-z0-9-]+)+([/?].*)?$")
        If checkWithoutHttp Or checkWithoutWWW Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function Matcher(Str As String, validationPattern As String) As Boolean
        Dim reg As New Regex(validationPattern)
        Return reg.IsMatch(Str)
    End Function
End Class
