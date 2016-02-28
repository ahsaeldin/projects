'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   CacadoreGloby Module                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Option Strict Off
'Global variables through out Cacadore.
Imports System.Drawing
Imports Cacadore.Commands
Imports System.ComponentModel

Module CacadoreGloby

#Region "Globy fields"
    Friend TriggersLastTimeUpdate As DateTime
    Friend TectonicCruding As New Crud()
    Friend InsertOperationDlg As New PersistenceOperationDlg(AddressOf TectonicCruding.Insert)
    Friend UpdateOperationDlg As New PersistenceOperationDlg(AddressOf TectonicCruding.Update)
    Friend DeleteOperationDlg As New PersistenceOperationDlg(AddressOf TectonicCruding.Delete)
    Friend QueryOperationDlg As New QueryPersistenceOperationDlg(AddressOf TectonicCruding.Query)
    Friend Delegate Function PersistenceOperationDlg(ByVal cacadoreObj As ICrudable, <[ParamArray]()> args() As String) As Boolean
    Friend Delegate Function QueryPersistenceOperationDlg(ByVal cacadoreObj As ICrudable, <[ParamArray]()> args() As String) As Object
#End Region

#Region "Globy Methods"
    ''' <summary>
    ''' Invokes a delegate that will execute a Crud operation against tectonic.
    ''' </summary>
    ''' <param name="callingObject">The calling object in which TypedDataSet will use to distinguish Cacacdore objects.</param>
    ''' <param name="args">Parameters that may be passed to the Crud methods.</param>
    ''' <param name="delegateObj">The CRUD operation you want to call.</param>
    ''' <param name="errorString">The error string used in RaiseErrorUp statement</param>
    ''' <returns>Return an object of type object and the caller method may convert it to the desired data type.</returns>
    Friend Function InvokeTectonicDlg(callingObject As Object,
                                           args() As String, delegateObj As [Delegate],
                                           errorString As String) As Object
        Try
            Dim resultantObject As Object
            SyncLock TectonicCruding
                'My genius idea of injecting Database stuff here.TectonicCaching.
                Dim Params() As Object = {callingObject, args}
                Watcher._isBusy = True
#If DEBUG Then
                Dim startTime As DateTime = DateTime.Now
#End If
                resultantObject = delegateObj.DynamicInvoke(Params)  'عنق الزجاجة'

#If DEBUG Then
                Dim executionTime As TimeSpan = DateTime.Now - startTime
#End If
                Watcher._isBusy = False
#If CONFIG = "Debug With Bottleneck" Then
                SendToBottleneckForm(callingObject, args, executionTime.TotalSeconds)
#End If
            End SyncLock
            Return resultantObject
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            If ex.Message.Contains(My.Resources.NoObjectFound) Then
                _privateRaiseErrorUp(errorString, ex, False)
            Else
                _privateRaiseErrorUp("", ex, True)
            End If
            Return Nothing
        End Try
    End Function

    <Conditional("DEBUG")>
    Sub SendToBottleneckForm(callingObject As Object,
                             args() As String,
                             _timing As Double)
        Dim bottleArgs(1) As String

        If args.Count > 0 Then
            If callingObject.ToString.Contains("Cacadore.Commands") Then
                Dim commandName As Commands.CommandsEnum = args(0)
                bottleArgs(0) = callingObject.ToString & "\" & commandName.ToString
            Else
                bottleArgs(0) = callingObject.ToString & "\" & args(0)
            End If
        Else
            bottleArgs(0) = callingObject.ToString
        End If

        bottleArgs(0) = bottleArgs(0).Replace("Cacadore.", "")
        bottleArgs(0) = bottleArgs(0).Replace("Commands\", "")
        If bottleArgs(0).Contains("Globals") Then bottleArgs(0) = "Globals\" & CType(args(0), IGlobals.GlobalType).ToString
        If bottleArgs(0).Contains("\") AndAlso Not bottleArgs(0).Contains("Globals") Then bottleArgs(0) = bottleArgs(0).Split("\")(0)
        bottleArgs(1) = _timing

        Balora.Alerter.SendMessageToCacadu("UpdateBottleneck", bottleArgs)
    End Sub
#End Region
End Module