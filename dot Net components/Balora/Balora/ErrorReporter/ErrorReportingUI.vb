Imports System.ComponentModel
Imports System.Reflection

Namespace ErrorReporter
    Public Class ErrorReportingForm

#Region "Constructors"
        Public Sub New()
            'in designer-generated type 'Balora.ErrorReportingUI' should call InitializeComponent method.	
            InitializeComponent()
            'Set ErrorReporterUI Default Text
            Me.Text = Util.GetCallingAssemblyName
        End Sub
#End Region

    End Class
End Namespace