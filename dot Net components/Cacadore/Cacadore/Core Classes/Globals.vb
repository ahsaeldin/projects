'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Globals Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Stores the global variables and constants among all tasks!
''' </summary>
''' <remarks>All variables and constants are shared among all tasks.</remarks>
Public Class Globals
    Inherits CacadoreBase
    Implements IGlobals
    Implements ICrudable
    Implements IDisposable

#Region "Fields"
    Private _constantsValue As Dictionary(Of String, String)
    Private _variablesValue As Dictionary(Of String, String)
#End Region

#Region "Properties"
    Public ReadOnly Property Constants As Dictionary(Of String, String) Implements IGlobals.Constants
        Get
            Return getGlobal(CStr(IGlobals.GlobalType.Constant))
        End Get
    End Property

    Public ReadOnly Property Variables As Dictionary(Of String, String) Implements IGlobals.Variables
        Get
            Return getGlobal(CStr(IGlobals.GlobalType.Variable))
        End Get
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Adds a new constant.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Function AddConstant(name As String,
                                value As String,
                                Optional errorMessage As String = vbNullString) As Boolean Implements IGlobals.AddConstant

        If Not AddGlobalConstantNotation(name, errorMessage) Then Return False

        Dim newConstant = createNewMember(IGlobals.GlobalType.Constant, name, value)

        If Not _isCrudObjectDefined() Then
            '++One day I want to implement this 
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(newConstant, {}, InsertOperationDlg, My.Resources.GlobNotPresisted))
        End If
    End Function

    Public Shared Function AddGlobalVariableNotation(ByRef name As String, ByRef errorMessage As String) As Boolean
        If Not validateName(name) Then
            errorMessage = My.Resources.InVaNa
            Return False
        End If

        If isMemeberHasNotation(name) = IGlobals.GlobalType.UnKnown Then
            name = "G<" & name & ">"
        ElseIf IsGlobalConstant(name) Then
            name = ConvertToOppositeType(name)
        End If

        Return True
    End Function

    Public Shared Function AddGlobalConstantNotation(ByRef name As String, ByRef errorMessage As String) As Boolean
        If Not validateName(name) Then
            errorMessage = My.Resources.InVaNa
            Return False
        End If

        If isMemeberHasNotation(name) = IGlobals.GlobalType.UnKnown Then
            name = "G{" & name & "}"
        ElseIf IsGlobalVariable(name) Then
            name = ConvertToOppositeType(name)
        End If

        Return True
    End Function

    Public Shared Function IsGlobalConstant(member As String) As Boolean
        Dim memberType As IGlobals.GlobalType = isMemeberHasNotation(member)
        If memberType = IGlobals.GlobalType.Constant Then Return True
        Return False
    End Function

    Public Shared Function IsGlobalVariable(member As String) As Boolean
        Dim memberType As IGlobals.GlobalType = isMemeberHasNotation(member)
        If memberType = IGlobals.GlobalType.Variable Then Return True
        Return False
    End Function

    Public Shared Function ConvertToOppositeType(member As String) As String
        If IsGlobalConstant(member) Then
            member = member.Replace("{", "<")
            member = member.Replace("}", ">")
        ElseIf IsGlobalVariable(member) Then
            member = member.Replace("<", "{")
            member = member.Replace(">", "}")
        End If
        Return member
    End Function

    Private Function createNewMember(memeberType As IGlobals.GlobalType, ByVal name As String, ByVal value As String) As GlobalEntry
        Select Case memeberType
            Case IGlobals.GlobalType.Constant

                Dim newConstant As New GlobalEntry
                newConstant.Name = name
                newConstant.Value = value
                newConstant.type = IGlobals.GlobalType.Constant
                Return newConstant

            Case IGlobals.GlobalType.Variable

                Dim newVariable As New GlobalEntry
                newVariable.Name = name
                newVariable.Value = value
                newVariable.type = IGlobals.GlobalType.Variable
                Return newVariable

        End Select
        Return Nothing
    End Function

    ''' <summary>
    ''' Adds a new variable.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Function AddVariable(name As String,
                                value As String,
                                Optional errorMessage As String = vbNullString) As Boolean Implements IGlobals.AddVariable
        If Not AddGlobalVariableNotation(name, errorMessage) Then
            Return False
        End If

        Dim newVariable As GlobalEntry = createNewMember(IGlobals.GlobalType.Variable, name, value)

        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(newVariable, {}, InsertOperationDlg, My.Resources.GlobNotPresisted))
        End If
    End Function

    Public Function EditVariable(name As String,
                                 newValue As String,
                                 Optional errorMessage As String = vbNullString) As Boolean Implements IGlobals.EditVariable
        If Not AddGlobalVariableNotation(name, errorMessage) Then
            Return False
        End If

        Dim variable = createNewMember(IGlobals.GlobalType.Variable, name, newValue)

        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(variable, {}, UpdateOperationDlg, My.Resources.GlobNotPresisted))
        End If
    End Function

    Public Function RemoveMemberByName(ByVal memberName As String) As Boolean Implements IGlobals.RemoveMemberByName
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {memberName}, DeleteOperationDlg, My.Resources.GlobNotPresisted))
        End If
    End Function

    Public Function IsRegisteredGlobalMember(memberName As String) As Boolean Implements IGlobals.IsRegisteredMember
        Dim key = From var In Variables Where var.Key = memberName

        If key.Count = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function getGlobal(ParamArray args() As String) As Dictionary(Of String, String)
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CType(InvokeTectonicDlg(Me, args, QueryOperationDlg, My.Resources.GlobNotPresisted), Dictionary(Of String, String))
        End If
    End Function

    Private Shared Function validateName(name As String, Optional memberType As IGlobals.GlobalType = Nothing) As Boolean
        Dim matcher As New Balora.RegexMatcher
        name = removeMemberNotation(name)
        Return matcher.MatchAlphanumeric(name)
    End Function

    Private Shared Function isMemeberHasNotation(memberName As String) As IGlobals.GlobalType
        If Strings.Left(memberName, 2) = "G{" AndAlso Strings.Right(memberName, 1) = "}" Then Return IGlobals.GlobalType.Constant

        If Strings.Left(memberName, 2) = "G<" AndAlso Strings.Right(memberName, 1) = ">" Then Return IGlobals.GlobalType.Variable

        Return IGlobals.GlobalType.UnKnown
    End Function

    Private Shared Function removeMemberNotation(memberName As String) As String
        If isMemeberHasNotation(memberName) <> IGlobals.GlobalType.UnKnown Then
            memberName = Strings.Mid(memberName, 3)
            memberName = Strings.Left(memberName, memberName.Length - 1)
        End If
        Return memberName
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