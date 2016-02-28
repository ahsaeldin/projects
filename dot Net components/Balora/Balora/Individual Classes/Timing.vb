Public Class Timing
    Inherits BaloraBase

    Private duration As TimeSpan
    Private startingTime As TimeSpan

    Public Sub New()
        startingTime = New TimeSpan(0)
        duration = New TimeSpan(0)
    End Sub

    Public ReadOnly Property Result() As TimeSpan
        Get
            Return duration
        End Get
    End Property

    Public Sub StartTime()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        startingTime = Process.GetCurrentProcess.Threads(0).UserProcessorTime
    End Sub

    Public Sub StopTime()
        Dim endTime = Process.GetCurrentProcess.Threads(0).UserProcessorTime
        duration = Process.GetCurrentProcess.Threads(0).UserProcessorTime.Subtract(startingTime)
    End Sub


#Region "Unit Test"
    Private Sub Test()
        Dim nums(99999) As Integer
        BuildArray(nums)
        Dim tObj As New Timing()
        tObj.StartTime()
        'ToDo: add tested method here.
        tObj.StopTime()
        Console.WriteLine("time (.NET): " & _
        tObj.Result.TotalSeconds)
        Console.Read()
    End Sub

    Private Sub BuildArray(ByVal arr() As Integer)
        Dim index As Integer
        For index = 0 To 99999
            arr(index) = index
        Next
    End Sub

#Region "References"
    'From Data Structures and algorithms using visual basic.net page 10
#End Region
#End Region
End Class
