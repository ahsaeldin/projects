#Region "Import Statements"
Imports System.Threading
#End Region

Public Class xTimer
    Implements IDisposable

#Region "Fields"
    Private _timer As System.Threading.Timer
    Private timerCallback As Threading.TimerCallback = AddressOf Timer_Callback
#End Region

#Region "Constructors"
    Public Sub New()
    End Sub
    Public Sub New(due As Integer, InterVal As Integer)
        DueTime = due
        TimerInterval = InterVal
    End Sub

    Public Sub New(due As Integer, InterVal As Integer, stateObj As Object)
        DueTime = due
        State = stateObj
        TimerInterval = InterVal
    End Sub

    Public Sub New(due As Integer, InterVal As Integer, startNow As Boolean, workToDo As [Delegate])
        DueTime = due
        TimerInterval = InterVal
        WorkDelegate = workToDo
        If startNow Then Start()
    End Sub

    Public Sub New(due As Integer, InterVal As Integer, workToDo As [Delegate])
        DueTime = due
        TimerInterval = InterVal
        WorkDelegate = workToDo
    End Sub
#End Region

#Region "Properties"
    Property DueTime As Integer = 0
    Property WorkDelegate As [Delegate]
    Property TimerInterval As Integer = 0
    Property State As Object = New Object
    Shared Property Switch As Boolean = True 'On.
#End Region

#Region "Methods"
    Sub Start()
        _timer = New System.Threading.Timer(timerCallback,
                                            State,
                                            DueTime,
                                            TimerInterval)
    End Sub

    Sub [Stop]()
        _timer.Change(Timeout.Infinite, Timeout.Infinite)
    End Sub

    Sub [Resume]()
        _timer.Change(_DueTime, _TimerInterval)
    End Sub

    Private Sub Timer_Callback(stateInfo As Object)
        If Switch = False Then
            'Balora.Hodhod.STOW("xTimer is disabled.")
            Exit Sub
        End If
        [Stop]()
        If Not IsNothing(WorkDelegate) Then
            WorkDelegate.DynamicInvoke()
        End If
        [Resume]()
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _timer.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            _timer = Nothing
            timerCallback = Nothing
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