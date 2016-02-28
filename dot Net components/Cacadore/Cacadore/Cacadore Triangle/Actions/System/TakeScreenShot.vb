'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ScheduleTrigger Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Namespace Actions
    Namespace System
        <Serializable()> Public Class TakeScreenShot
            Inherits Action
            Implements IDisposable

#Region "Properties"
            Public Overrides ReadOnly Property ActionType As IAction.ActionTypeEnum
                Get
                    Return IAction.ActionTypeEnum.TakeScreenShot
                End Get
            End Property
#End Region

            Public Overrides Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
                'TODO : contains the logic of execution
                'Set Input Property here
                'Set Output Property @ return
                'Set Result Property @ return
                'And after execution, you will have a result with input and output
                'And @ finish return "ITaskState.ExecutionState" that indicates the result
                Return Nothing
            End Function

            Public Overrides Property Inputs As IAction.IInputs
                Get
                    Return Nothing 'for now.
                End Get
                Set(value As IAction.IInputs)

                End Set
            End Property

            Public Overrides Property Outputs As IAction.IOutputs
                Get
                    Return Nothing 'for now.
                End Get
                Set(value As IAction.IOutputs)

                End Set
            End Property

            Public Overrides Property Errors As IAction.IErrors
                Get
                    Return Nothing 'for now.
                End Get
                Set(value As IAction.IErrors)

                End Set
            End Property

            Public Overrides Sub SetPossibleExceptions()

            End Sub

            Protected Overrides Sub WaitForExit(Optional ByRef state As Object = Nothing)

            End Sub

            Public Overrides Function GetFullDescription() As String
                Return "إبحث عني يا طور"
            End Function

            Public Overrides Function Clone() As Action
                ''Return "فاضي"
                Return Nothing
            End Function

#Region "IDisposable Support"
            Private disposedValue As Boolean ' To detect redundant calls

            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not Me.disposedValue Then
                    If disposing Then
                        ' TODO: dispose managed state (managed objects).
                    End If

                    ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                    ' TODO: set large fields to null.
                End If
                Me.disposedValue = True
            End Sub

            ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
            'Protected Overrides Sub Finalize()
            '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            '    Dispose(False)
            '    MyBase.Finalize()
            'End Sub

            ' This code added by Visual Basic to correctly implement the disposable pattern.
            Public Sub Dispose() Implements IDisposable.Dispose
                ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub
#End Region

        End Class
    End Namespace
End Namespace