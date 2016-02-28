Imports System.ComponentModel

Friend Class HTMLReportEngine
    Inherits BaloraBase
    'http://www.codeproject.com/KB/cs/HTMLReportEngine.aspx

    ''' <summary>
    ''' Gets or Sets report source as DataSet.
    ''' </summary>
    Private reportSourceValue As DataSet
    Public Property ReportSource() As DataSet
        Get
            Return reportSourceValue
        End Get
        Set(ByVal value As DataSet)
            reportSourceValue = value
        End Set
    End Property

End Class