Imports System.ComponentModel

'BackgroundWorkerHelper = BGH
Public Class BackgroundWorkerHelper
    Inherits BaloraBase
    Implements IDisposable

#Region "Fields"
    'The private field of WorkToDo, that will be used to...
    'Inject the method to be implemented in DoWorkEventHandler. 
    Private workToDoValue As [Delegate]
    Private progressChangedDlgValue As [Delegate]
    Private runWorkerCompletedDlgValue As [Delegate]
    'The private field of BackGroundWorkerThread, that will do the work threadly.
    Private WithEvents backGroundWorkerValue As New BackgroundWorker
#End Region

#Region "Constructors"
    Public Sub New()
        backGroundWorkerValue.WorkerReportsProgress = True
        backGroundWorkerValue.WorkerSupportsCancellation = True
    End Sub
#End Region

#Region "Properties"
    ''' <summary>
    ''' Used to inject the method to be called by BackGroundWorkerThread.
    ''' </summary>
    Public Property WorkToDo() As [Delegate]
        Get
            Return workToDoValue
        End Get
        Set(ByVal value As [Delegate])
            workToDoValue = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the run worker process changed event delegate for BackgroundWorker.
    ''' </summary>
    ''' <value>
    ''' The progress changed DLG.
    ''' </value>
    Public Property ProgressChangedDlg() As [Delegate]
        Get
            Return progressChangedDlgValue
        End Get
        Set(ByVal value As [Delegate])
            progressChangedDlgValue = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the run worker completed event delegate for BackgroundWorker.
    ''' </summary>
    ''' <value>
    ''' The run worker completed DLG.
    ''' </value>
    Public Property RunWorkerCompletedDlg() As [Delegate]
        Get
            Return runWorkerCompletedDlgValue
        End Get
        Set(ByVal value As [Delegate])
            runWorkerCompletedDlgValue = value
        End Set
    End Property

    ''' <summary>
    ''' The BackGroundWorker thread that will do the work threadly.
    ''' </summary>
    Public Property BackGroundWorkerObject() As BackgroundWorker
        Get
            Return backGroundWorkerValue
        End Get
        Set(ByVal value As BackgroundWorker)
            backGroundWorkerValue = value
        End Set
    End Property

#End Region

#Region "Methods"
    ''' <summary>
    ''' Start the async work.
    ''' </summary>
    ''' <returns>
    ''' If true then it has been started, else the background worker is busy now.
    ''' </returns>
    Public Function Start(Optional arguments As Object = Nothing) As Boolean
        If Not isBusyNow() Then
            Return callRunWorkerAsynic(arguments)
        Else
            Return onBusy()
        End If
    End Function

    Public Sub Kill()
        backGroundWorkerValue.Dispose()
    End Sub

    Private Function callRunWorkerAsynic(ByVal arguments As Object) As Boolean
        backGroundWorkerValue.RunWorkerAsync(arguments)
        Return True
    End Function

    Private Function isBusyNow() As Boolean
        Return backGroundWorkerValue.IsBusy()
    End Function

    Private Function onBusy() As Boolean
        'هام
        '#################################################################
        'In the future you can do the work just after
        'the worker is not busy be settings a manual
        'lock here and release in backGroundWorkerValue_RunWorkerCompleted
        'Also, you can add a property called
        '"Automatically do work if busy after compilation".
        '#################################################################
        RaiseEvent Busy(Me, New EventArgs)
        'Background worker is busy now.
        Return False
    End Function

    ''' <summary>
    ''' Handles the DoWork event of the backGroundWorkerValue control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.ComponentModel.DoWorkEventArgs" /> instance containing the event data.</param>
    Private Sub backGroundWorkerValue_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles backGroundWorkerValue.DoWork
        workToDoValue.DynamicInvoke(sender, e)
    End Sub

    ''' <summary>
    ''' Handles the ProgressChanged event of the backGroundWorkerValue control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.ComponentModel.ProgressChangedEventArgs" /> instance containing the event data.</param>
    Private Sub backGroundWorkerValue_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles backGroundWorkerValue.ProgressChanged
        ProgressChangedDlg.DynamicInvoke(sender, e)
    End Sub

    ''' <summary>
    ''' Handles the RunWorkerCompleted event of the backGroundWorkerValue control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.ComponentModel.RunWorkerCompletedEventArgs" /> instance containing the event data.</param>
    Private Sub backGroundWorkerValue_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backGroundWorkerValue.RunWorkerCompleted
        RunWorkerCompletedDlg.DynamicInvoke(sender, e)
        CType(sender, BackgroundWorker).Dispose()
    End Sub
#End Region

#Region "Events"
    Public Event Busy As EventHandler
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                backGroundWorkerValue.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            workToDoValue = Nothing
            progressChangedDlgValue = Nothing
            runWorkerCompletedDlgValue = Nothing
            backGroundWorkerValue = Nothing
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