'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
'                                                   CacadorePrivateCalls Module
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /	
'To hide calling to a certain functions from Reflectors.
Friend Module CacadorePrivateCalls
    Friend Sub _privateRaiseErrorUp(errorDescription As String,
                             ex As Exception,
                             Optional broadCast As Boolean = False,
                             Optional justLog As Boolean = False)
        Balora.Alerter.REP(errorDescription, ex, broadCast, justLog)
    End Sub

    Friend Sub _privateShowMessageBox(msg As String)
        MsgBox(msg)
    End Sub
    Friend Function _isCrudObjectDefined() As Boolean
        If Settings.CrudObject Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
End Module