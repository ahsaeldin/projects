Imports Cacadore
Imports System.ComponentModel

Friend Class TriggersStarter
    Implements IDisposable

    Private Shared bckGrndWrkHelper As Balora.BackgroundWorkerHelper
    Private Delegate Sub DoWorkDlg(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

#Region "Properties"
    Shared Property TriggersReStarted As Boolean = False
#End Region

#Region "Methods"
    Function TryStartTriggers() As Boolean
        Return Start()
    End Function

    Function Start() As Boolean
        If Cacadore.Settings.Enabled Then
            bckGrndWrkHelper = New Balora.BackgroundWorkerHelper
            bckGrndWrkHelper.WorkToDo = New DoWorkDlg(AddressOf WorkToDo)
            bckGrndWrkHelper.RunWorkerCompletedDlg = New RunWorkerCompletedEventHandler(AddressOf RunWorkerCompleted)

            Dim trgs = Commands.GetAllTriggers
            If trgs.Count > 0 AndAlso My.Settings.AMFBRTAS Then
                ShowAlertForm("Cacadu",
                              "Do you want to start your triggers?",
                              Cacadore.Actions.MessageBoxes.ShowAlert.ShowAlertsInputs.AlertLocationEnum.BottomRight,
                              "AskToStartTriggers",
                              True,
                              My.Resources.Cacadu_Image_46x46,
                              Me)
                Return False
            End If

            bckGrndWrkHelper.Start(trgs)
            Return True
        End If
        Return False
    End Function

    Private Sub WorkToDo(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

        Dim trgs As Dictionary(Of String, Trigger) = CType(e.Argument, Global.System.Collections.Generic.Dictionary(Of String, Global.Cacadore.Trigger))

        If trgs.Count > 0 AndAlso My.Settings.AST Then
            For Each trg In trgs.Values
                trg.StartListening()
                DeleteEndedTriggers(trg)
                If trg.MayFireAgain Then TriggersReStarted = True
            Next
        End If
    End Sub

    Private Sub DeleteEndedTriggers(trg As Trigger)
        If Not trg.MayFireAgain Then
            trg.Task.UnAttachTrigger(trg)
        End If
    End Sub

    Private Sub RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If TriggersReStarted Then ShowAlertForm("Cacadu", "All triggers are ready.", AlertLocation, , , My.Resources.Cacadu_Image_46x46)
        Dispose()
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                bckGrndWrkHelper.Dispose()
            End If
            bckGrndWrkHelper = Nothing
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
