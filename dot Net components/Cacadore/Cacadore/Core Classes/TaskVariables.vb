Imports System.ComponentModel
Imports System.Globalization
Imports System.Drawing
Imports System.ComponentModel.Design

'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskVariables Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Very important 
'Task Variables must be attached to Task class in order to be saved in database.
<TypeConverter(GetType(TaskVariablesConverter))>
<Serializable()> Public Class TaskVariables
    Inherits CacadoreBase
    Implements ITaskVariables
    Implements ICrudable
    Implements IDisposable

    <NonSerialized()> Private bckGrndWrkHelper As Balora.BackgroundWorkerHelper
    Delegate Sub doWordDelegateOfDeleteVariables(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

#Region "Constructors"
    'This is very important
    'Direct assignment of Task.TaskVariables property to any
    'other TaskVariable class is acceptable only if both has
    'the same id, otherwise it won't set. Check TaskVariables 
    'Property in Task Class.
    Public Sub New(taskId As String)
        TaskIdValue = taskId
    End Sub
#End Region

#Region "Fields"
    Friend TaskIdValue As String
#End Region

#Region "Properties"
    <System.ComponentModel.Editor(GetType(VariablesCollectionEditor), GetType(System.Drawing.Design.UITypeEditor))>
    Public ReadOnly Property Variables As VariablesCollection Implements ITaskVariables.Variables
        Get
            Return createVariablesCollection()
        End Get
    End Property

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property TaskId As String Implements ITaskVariables.taskId
        Get
            Return TaskIdValue
        End Get
    End Property
#End Region

#Region "Methods"
    'Not e that if there is no task with this id, it won't be saved to database.
    Public Function AddVariable(name As String, value As String) As Boolean Implements ITaskVariables.AddVariable
        Return setVariable(name, value, True)
    End Function

    Public Function UpdateVariable(name As String, newValue As String) As Boolean Implements ITaskVariables.UpdateVariable
        Return setVariable(name, newValue)
    End Function

    'Return false if it has no changes.
    Public Function UpdateVariables(newVariablesCollection As VariablesCollection) As Boolean Implements ITaskVariables.UpdateVariables
        Dim newVariables As Dictionary(Of String, String) = newVariablesCollection.InnerDictionary
        Dim storedVariables As Dictionary(Of String, String) = Variables.InnerDictionary
        Dim res1 = checkForVariablesToBeAddedOrUpdated(newVariables, storedVariables)
        Dim res2 = checkForVariablesToBeDeleted(newVariables, storedVariables)
        If res1 OrElse res2 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function DeleteVariable(ByVal variableName As String) As Boolean Implements ITaskVariables.RemoveVariable
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {variableName}, DeleteOperationDlg, My.Resources.VarsNotPresisted))
        End If
    End Function

    Public Function DeleteVariables() As Boolean Implements ITaskVariables.DeleteVariables
        Dim variablesCount As Integer = Variables.Count
        If variablesCount <= 0 Then Return True
        bckGrndWrkHelper = New Balora.BackgroundWorkerHelper
        AddHandler bckGrndWrkHelper.BackGroundWorkerObject.RunWorkerCompleted, AddressOf RunWorkerCompleted
        bckGrndWrkHelper.WorkToDo = New doWordDelegateOfDeleteVariables(AddressOf backGroudWorkerThread_DoWork)
        bckGrndWrkHelper.RunWorkerCompletedDlg = New RunWorkerCompletedEventHandler(AddressOf RunWorkerCompleted)
        Return bckGrndWrkHelper.Start()
    End Function

    Public Function GetVariableValue(variableName As String) As String Implements ITaskVariables.GetVariableValue
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CStr(InvokeTectonicDlg(Me, {"GetVariableValue", Me.TaskId, variableName}, QueryOperationDlg, My.Resources.VarsNotPresisted))
        End If
    End Function

    Public Shared Function AddVariableNotation(ByRef name As String, ByRef errorMessage As String) As Boolean 'Implements ITaskVariables.AddVariableNotation
        If isVariableNameHasNotation(name) AndAlso validateVariableName(name) Then Return True
        checkForSingleBracketCase(name)
        If Not validateVariableName(name) Then
            errorMessage = My.Resources.InVaNa
            Return False
        End If
        name = String.Format("<{0}>", name)
        Return True
    End Function

    ''' <summary>
    ''' Checks for single bracket case. having a single bracket in the right or left.
    ''' </summary>
    ''' <param name="name">The name.</param>
    Private Shared Sub checkForSingleBracketCase(ByRef name As String)
        If name.First = "<" Then
            name = name.Remove(0, 1)
        ElseIf name.Last = ">" Then
            Dim lastCharIndex As Integer = name.Length - 1
            name = name.Remove(lastCharIndex)
        End If
    End Sub

    Private Sub backGroudWorkerThread_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Dim tskName As String = Commands.GetTaskProperty(Me.TaskId, TaskProperties.PropertiesNames.TaskName).ToString
        Balora.Alerter.ShowWaitWinodw(String.Format("Deleting the variables of task {0}{1}{0}...", Chr(34), tskName))
        e.Result = Commands.DeleteTaskVariables(Me.TaskId)
    End Sub

    Private Sub RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        Balora.Alerter.CloseWaitWinodw()
    End Sub

    ''' <summary>
    ''' Checks for variables to be added or updated.
    ''' </summary>
    ''' <param name="newVariables">The new variables.</param>
    ''' <param name="storedVariables">The stored variables.</param><returns></returns>
    ''' <remarks>
    ''' Compare the new variables dictionary returned by a collection editor 
    ''' with a the stored dictionary stored in Tectonic. This check can get
    ''' newly added or updated variables only.
    '''  </remarks>
    Private Function checkForVariablesToBeAddedOrUpdated(newVariables As Dictionary(Of String, String), storedVariables As Dictionary(Of String, String)) As Boolean
        For Each var In newVariables
            If storedVariables.ContainsKey(var.Key) Then
                If Not newVariables(var.Key) = storedVariables(var.Key) Then
                    'Change variable value here...
                    UpdateVariable(var.Key, var.Value)
                    Return True
                End If
            Else
                'Add new variable here...
                AddVariable(var.Key, var.Value)
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Checks for variables to be deleted.
    ''' </summary>
    ''' <param name="newVariables">The new variables.</param>
    ''' <param name="storedVariables">The stored variables.</param><returns></returns>
    ''' <remarks>
    ''' Compare the stored variables of a given task with the new variables returned by
    ''' a collection edtior. This check can get deleted variables only.
    ''' </remarks>
    Private Function checkForVariablesToBeDeleted(newVariables As Dictionary(Of String, String), storedVariables As Dictionary(Of String, String)) As Boolean
        For Each var In storedVariables
            If Not newVariables.ContainsKey(var.Key) Then
                'Delete variable here...
                DeleteVariable(var.Key)
                Return True
            End If
        Next
        Return False
    End Function

    Private Function createVariablesCollection() As VariablesCollection
        Dim vasrColl As New VariablesCollection
        Dim varsDic = getVariables()
        For Each var In varsDic
            vasrColl.Add(createNewVariableObject(var.Key, var.Value))
        Next
        Return vasrColl
    End Function

    Private Function getVariables() As Dictionary(Of String, String)
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CType(InvokeTectonicDlg(Me, {Me.TaskId}, QueryOperationDlg, My.Resources.VarsNotPresisted), Dictionary(Of String, String))
        End If
    End Function

    Private Shared Function validateVariableName(name As String) As Boolean
        name = name.Replace("<", "")
        name = name.Replace(">", "")
        Dim matcher As New Balora.RegexMatcher
        Return matcher.MatchAlphanumeric(name)
    End Function

    Private Function createNewVariableObject(ByVal name As String, ByVal value As String) As Variable
        Dim var As New Variable
        var.TaskId = Me.TaskId
        var.Name = name
        var.Value = value
        Return var
    End Function

    Private Shared Function sendToTectonic(ByVal insert As Boolean,
                                           ByVal _variable As Variable) As Boolean
        If insert Then
            Return CBool(InvokeTectonicDlg(_variable, {}, InsertOperationDlg, My.Resources.VarsNotPresisted))
        Else
            Return CBool(InvokeTectonicDlg(_variable, {}, UpdateOperationDlg, My.Resources.VarsNotPresisted))
        End If
    End Function

    Private Function setVariable(name As String, value As String,
                                 Optional insert As Boolean = False,
                                 Optional errorMessage As String = vbNullString) As Boolean
        If Not AddVariableNotation(name, errorMessage) Then Return False
        Dim _variable As Variable = createNewVariableObject(name, value)
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return sendToTectonic(insert, _variable)
        End If
    End Function

    Private Shared Function isVariableNameHasNotation(varName As String) As Boolean
        If Strings.Left(varName, 1) = "<" AndAlso Strings.Right(varName, 1) = ">" Then Return True
        Return False
    End Function
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
            If Not IsNothing(bckGrndWrkHelper) Then RemoveHandler bckGrndWrkHelper.BackGroundWorkerObject.RunWorkerCompleted, AddressOf RunWorkerCompleted

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

#Region "Helper Classes"

'To display TaskVariables in a property grid as an Expandable Object
Friend Class TaskVariablesConverter
    Inherits ExpandableObjectConverter

    Public Overloads Overrides Function CanConvertTo(ByVal context As ITypeDescriptorContext,ByVal destinationType As Type) As Boolean
        If (destinationType Is GetType(TaskVariables)) Then
            Return True
        End If
        Return MyBase.CanConvertFrom(context, destinationType)
    End Function

    Public Overloads Overrides Function ConvertTo(ByVal context As ITypeDescriptorContext,
                                                  ByVal culture As CultureInfo,
                                                  ByVal value As Object,
                                                  ByVal destinationType As System.Type) As Object
        If (destinationType Is GetType(System.String) _
            AndAlso TypeOf value Is TaskVariables) Then
            Dim tskVars As TaskVariables = CType(value, TaskVariables)
            Dim varsCount = tskVars.Variables.Count
            If varsCount = 1 Then
                Return String.Format("Variables: {0}{1}", tskVars.Variables.Item(0).Name, tskVars.Variables.Item(0).Value)
            ElseIf varsCount > 1 Then
                Return String.Format("Variables: {0}{1}...", tskVars.Variables.Item(0).Name, tskVars.Variables.Item(0).Value)
            ElseIf varsCount = 0 Then
                Return "Variables: " & "No Variable Defined"
            End If
        End If
        Return MyBase.ConvertTo(context, culture, value, destinationType)
    End Function
End Class

Public Class VariablesCollection
    Inherits CollectionBase
    Default Public ReadOnly Property Item(index As Integer) As Variable
        Get
            Return DirectCast(List(index), Variable)
        End Get
    End Property

    Public Sub Add(emp As Variable)
        List.Add(emp)
    End Sub

    Public Sub Remove(emp As Variable)
        List.Remove(emp)
    End Sub

    Public ReadOnly Property InnerDictionary() As Dictionary(Of String, String)
        Get
            Dim _variablesDictionaryValue As New Dictionary(Of String, String)
            For Each var In List
                Dim varName As String = DirectCast(var, Variable).Name
                Dim varValue As String = DirectCast(var, Variable).Value
                'Add it, if a new variable (null variable names aren't allowed.)
                If varName <> vbNullString AndAlso Not _variablesDictionaryValue.ContainsKey(varName) Then
                    _variablesDictionaryValue.Add(varName, varValue)
                ElseIf varName <> vbNullString AndAlso _variablesDictionaryValue.ContainsKey(varName) Then
                    'Update variable, if the newly added variable is already registered.
                    _variablesDictionaryValue(varName) = varValue
                End If
            Next
            Return (_variablesDictionaryValue)
        End Get
    End Property
End Class

Friend Class VariablesCollectionEditor
    Inherits CollectionEditor
    Public Sub New(type As Type)
        MyBase.New(type)
    End Sub

    Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext,
                                        provider As IServiceProvider,
                                        value As Object) As Object
        Dim vars As VariablesCollection = DirectCast(value, VariablesCollection)
        Dim tskVars As TaskVariables = DirectCast(context.Instance, TaskVariables)
        Dim innerDictionaryBeforeEditing = vars.InnerDictionary
        'Very important:-
        'http://social.msdn.microsoft.com/Forums/zh/winformsdatacontrols/thread/571fcee9-652b-43a0-b71e-b3ab956f388d
        'Code following this line has a guarantee that the collection editor has already
        'been dismissed, and changes to the collection have been applied.
        'context.Instance is the object that holds the property just edited.
        'In the case of the CodeProject example, it is of type Organization.
        'value is the collection itself.
        'Do whatever you need to do to save the collection here.  Either by a method
        'provided by the object owner of the collection, or the collection itself.
        'Example:  Let's imagine that there is a SaveCollection() method in the object
        'holding the collection.
        '((Organization)context.Instance).SaveCollection();
        Dim obj = MyBase.EditValue(context, provider, value)

        updateVariablesAfterEditing(vars)
        Return obj
    End Function

    Private Sub updateVariablesAfterEditing(vars As VariablesCollection)
        Dim tskVars As TaskVariables = DirectCast(Context.Instance, TaskVariables)
        If vars.Count <> 0 Then
            tskVars.UpdateVariables(vars)
        ElseIf vars.Count = 0 Then
            tskVars.DeleteVariables()
        End If
    End Sub

    Protected Overrides Function GetDisplayText(value As Object) As String
        Dim item As New Variable()
        item = DirectCast(value, Variable)
        Return MyBase.GetDisplayText(String.Format("{0}, {1}", item.Name, item.Value))
    End Function
End Class
#End Region
