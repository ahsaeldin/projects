Imports System.ComponentModel

<LicenseProvider(GetType(Balora.Lixcer.RTALP))> <Serializable()>
Public Class BaloraBase

#Region "Fields"
    Shared x28549 As License = Nothing
#End Region

#Region "Constructors"
    Public Sub New()
        If IsNothing(x28549) Then
            is_x28549(GetType(BaloraBase), Me)
        End If
    End Sub
#End Region

#Region "Methods"
    ''' <summary>
    ''' This function to validate the license of the caller class
    ''' </summary>
    ''' <remarks>
    '''  How to call
    '''  is_x28549(Me.GetType, Me)
    ''' </remarks>
    Private Function is_x28549(typ As Type, instance As Object) As Object
        Try
            validate(typ, instance)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            killProcess()
        End Try
        Return x28549
    End Function

    Private Sub validate(typ As Type, instance As Object)
        x28549 = LicenseManager.Validate(typ, instance)
    End Sub

    Private Sub killProcess()
        'الرسالة' دي ممكن تكون ماركر للهاكر
        ' Catch any error, but especially licensing errors...
        'Dim strErr As String = String.Format("Error executing application: '{0}'", ex.Message)
        ' MessageBox.Show(strErr, "VB RegistryLicensedApplication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        'Environment.Exit()
        'If unauthorized use of Balore then kill the whole app.
        Process.GetCurrentProcess.Kill()
    End Sub
#End Region

End Class
