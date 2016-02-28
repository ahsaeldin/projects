'To hide calling to a certain functions from Reflectors.
Friend Module BaloraPrivateCalls
    Friend Sub _privateRaiseErrorUp(errorDescription As String,
                             ex As Exception,
                             Optional broadCast As Boolean = False,
                             Optional justLog As Boolean = False)
        Balora.Alerter.REP(errorDescription, ex, broadCast, justLog)
    End Sub

    Friend Sub _privateShowMessageBox(msg As String)
        MsgBox(msg)
    End Sub
End Module
