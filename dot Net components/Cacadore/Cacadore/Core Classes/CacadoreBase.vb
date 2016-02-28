'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   CacadoreBase Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports System.ComponentModel

''' <summary>
''' CacadoreBase works as stub for all Cacadorian classes.
''' </summary>
<LicenseProvider(GetType(Licenser.RunTimeActivationLicenseProvider))> <Serializable()>
Public Class CacadoreBase

#Region "Fields"
    'Shared to save between calls
    Shared x32651 As License = Nothing
#End Region

#Region "Constructors"
    Public Sub New()
        If IsNothing(x32651) Then
            'For unknown reason, I must pass CacadoreBase type not "Me.GetType()"
            is_x32651(GetType(CacadoreBase), Me)
        End If
    End Sub
#End Region

#Region "Methods"
    ''' <summary>
    ''' This function to validate the license of the caller class
    ''' </summary>
    ''' <remarks>
    '''  How to call
    '''  is_x32651("the base class type that implement the licensing here", Me)
    ''' </remarks>
    Private Function is_x32651(typ As Type, instance As Object) As Object
        Try
            validate(typ, instance)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            killCurrentProcess()
        End Try
        Return x32651
    End Function

    Private Sub validate(typ As Type, instance As Object)
        x32651 = LicenseManager.Validate(typ, instance)
    End Sub

    Private Sub killCurrentProcess()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'الرسالة' دي ممكن تكون ماركر للهاكر
        ' Catch any error, but especially licensing errors...
        'Dim strErr As String = String.Format("Error executing application: '{0}'", ex.Message)
        ' MessageBox.Show(strErr, "VB RegistryLicensedApplication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Environment.Exit()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If unauthorized use of Cacadore then kill the whole app.
        invokeMethod(Process.GetCurrentProcess, "Kill")
    End Sub

    'تحفة -- كل حاجة مخفية
    Private Function invokeMethod(instance As Object, methodName As String, Optional parameters() As Object = Nothing) As Object
        'Getting the method information using the method info class
        'invokeing the method
        'null- no parameter for the function [or] we can pass the array of parameters
        Return instance.[GetType]().GetMethod(methodName).Invoke(instance, parameters)
    End Function
#End Region
End Class

