'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskProperties Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports System.ComponentModel
Imports System.Globalization

<System.ComponentModel.DefaultProperty("TaskName")>
<TypeConverter(GetType(TaskPropertiesConverter))>
<Serializable()> Public Class TaskProperties
    Inherits CacadoreBase
    Implements ITaskProperties
    Implements ICrudable
    Implements IDisposable

    Enum PropertiesNames
        TaskName
        IsTaskEnabled
        LastTimeRun
        NextTimeToBeRun
        TriggerWhen
        LastTimeRunResult
    End Enum

#Region "Fields"
    'Mark all private fields as non serialiable to hide them from reflectors except the one
    'we will pass to Util.SetPrivateField function.
    Private _taskId As String
    Private _taskNameValue As String
    Private _lasttimeRunValue As Date?
    Private _triggersWhenValue As String
    Private _lastTimeRunResult As String
    Private _nextTimeToBeRunValue As Date?
    Private _isEnabledValue As Boolean = True
#End Region

#Region "Constructors"
    Public Sub New(taskId As String)
        init(taskId)
    End Sub

    Public Sub New(taskId As String, ByVal newTaskName As String)
        init(taskId, newTaskName)
    End Sub

    Private Sub init(taskId As String, Optional ByVal newTaskName As String = vbNullString)
        _taskId = taskId
        If newTaskName = vbNullString Then
            'Just the default value that can be
            'changed later.
            TaskName = "New Task " & DateTime.Now
        Else
            TaskName = newTaskName
        End If
    End Sub
#End Region

#Region "Properties"

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property taskId As String Implements ITaskProperties.taskId
        Get
            Return _taskId
        End Get
    End Property

    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Task Name")>
    <System.ComponentModel.Description("Selected task name.")>
    Public Property TaskName As String Implements ITaskProperties.TaskName
        Get
            Return _taskNameValue
        End Get
        Set(ByVal value As String)
            _taskNameValue = value
        End Set
    End Property

    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Enabled")>
    <System.ComponentModel.Description("Enable/Disable task execution.")>
    Public Property IsEnabled As Boolean Implements ITaskProperties.IsEnabled
        Get
            Return _isEnabledValue
        End Get
        Set(ByVal value As Boolean)
            _isEnabledValue = value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public Property LastTimeRun As Date? Implements ITaskProperties.LastTimeRun
        Get
            Return _lasttimeRunValue
        End Get
        Set(ByVal value As Date?)
            _lasttimeRunValue = value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public Property NextTimeToBeRun As Date? Implements ITaskProperties.NextTimeToBeRun
        Get
            Return _nextTimeToBeRunValue
        End Get
        Set(ByVal value As Date?)
            _nextTimeToBeRunValue = value
        End Set
    End Property

    'For both, they will be blank for now until I finish the logic of triggers
    'Tasks>=Triggers_When ... TaskProperties.Triggers_When
    'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
    <System.ComponentModel.ReadOnlyAttribute(True)>
    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Triggers When")>
    <System.ComponentModel.Description("Description of task's triggers settings.")>
    Public Property TriggersWhen As String Implements ITaskProperties.TriggersWhen
        Get
            Return _triggersWhenValue
        End Get
        Set(ByVal value As String)
            _triggersWhenValue = value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property LastTimeRunResult As String Implements ITaskProperties.LastTimeRunResult
        Get
            Return _lastTimeRunResult
        End Get
    End Property

    Public Sub SetLastTimeRunResult(result As String)
        _lastTimeRunResult = result
    End Sub
#End Region

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

'To display TaskProperties in a property grid as an Expandable Object
Friend Class TaskPropertiesConverter
    Inherits ExpandableObjectConverter

    Public Overloads Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext, _
                                                     ByVal destinationType As Type) As Boolean
        If (destinationType Is GetType(TaskProperties)) Then
            Return True
        End If
        Return MyBase.CanConvertFrom(context, destinationType)
    End Function

    Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext, _
                                                  ByVal culture As CultureInfo, _
                                                  ByVal value As Object, _
                                                  ByVal destinationType As System.Type) _
                     As Object
        If (destinationType Is GetType(System.String) _
            AndAlso TypeOf value Is TaskProperties) Then

            Dim so As TaskProperties = CType(value, TaskProperties)

            Return "TaskName: " & so.TaskName & _
                   ", TriggersWhen: " & so.TriggersWhen & _
                   ", IsEnabled: " & so.IsEnabled
        End If
        Return MyBase.ConvertTo(context, culture, value, destinationType)
    End Function
End Class