'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Settings Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Option Strict Off
Imports System.Drawing
Public Class Settings
    Inherits CacadoreBase
    Implements ISettings
    Implements ICrudable
    Implements IDisposable

    Structure SettingEntry
        Public Enabled As Boolean
        Public HistoryEnabled As Boolean
        Public TriggersEnabled As Boolean
    End Structure

#Region "Fields"
    Private Shared isEnabledValue As Boolean
    Private Shared isHistoryEnabled As Boolean
    Private Shared isTriggersEnabled As Boolean
    'Default values is true in case you use the second
    'constructor New(update).
    Private Shared _crudObjectValue As ICrud
#End Region

#Region "Properties"
    Private Shared Property IsLoaded As Boolean = False
    ''' <summary>
    ''' Enable/Disable The execution of Cacadore Tasks however Triggers is still listening.
    ''' </summary>
    Public Shared Property Enabled As Boolean
        Get
            If Not IsLoaded Then loadAllSettings()
            Return isEnabledValue
        End Get
        Set(ByVal value As Boolean)
            isEnabledValue = value
            Save(ISettings.SettingsEnum.CacadoreEnabled, value)
            RaiseEvent CacadoreEnabledChanged(value)
            History.OnEnabledChanged(value)
        End Set
    End Property

    ''' <summary>
    ''' Enable/Disable all Cacadore's triggers, 
    ''' even if all of them are enabled.
    ''' </summary>
    Public Shared Property TriggersEnabled As Boolean
        Get
            If Not IsLoaded Then loadAllSettings()
            Return isTriggersEnabled
        End Get
        Set(ByVal value As Boolean)
            Select Case value
                Case True
                    Scheduler.ResumeAll()
                Case False
                    Scheduler.PauseAll()
            End Select
            isTriggersEnabled = value
            Save(ISettings.SettingsEnum.TriggersEnabled, value)
            RaiseEvent TriggersEnabledChanged(value)
        End Set
    End Property

    ''' <summary>
    ''' Enable/Disable History logging.
    ''' </summary>
    Public Shared Property HistoryEnabled As Boolean
        Get
            If Not IsLoaded Then loadAllSettings()
            Return isHistoryEnabled
        End Get
        Set(value As Boolean)
            isHistoryEnabled = value
            Save(ISettings.SettingsEnum.HistoryEnabled, value)
        End Set
    End Property

    '<summary>
    'A hidden property used in Cacadore that will set by the calling
    'Application that will set it to a md5 value of a guid
    'The idea here
    'is to check the equity of a md5 value of a GUID specified here by me
    'And Cacadu will be the only who knows this vlaue and pass its hash to me.
    '</summary>
    '<remarks>
    'CAG means Calling Assembly Guid.
    'We make the property type of Object
    'rather than string to hide any hint
    ''about what CAG do
    '</remarks>
    Private Shared cAGValue As Object
    Public Shared Property CAG() As Object
        Get
            Return cAGValue
        End Get
        Set(ByVal value As Object)
            cAGValue = value
        End Set
    End Property

    '1.Crud & ICrud is defined and implemented in Cacadore.
    '2.Cacadore classes invoke CRUD operations using Delegates.
    '3.Delegates execuate corresponding methods in an instance of Crud
    '4.Now the instance of CRUD execute the corresponding CRUD method in
    'the shared CrudObject property of settings class.
    '5.So that's it, all you need to do from Cacadu is to set CrudObject to an ICrud
    'class the implements the CRUD functions.
    Public Shared Property CrudObject As ICrud
        Get
            Return _crudObjectValue
        End Get
        Set(value As ICrud)
            _crudObjectValue = value
            TectonicCruding.CrudObject = value
        End Set
    End Property

    'How To Use:
    '   Dim sf As ISchedulerFactory = New StdSchedulerFactory
    '   Cacadore.Settings.Scheduler = sf.GetScheduler
    Shared Property Scheduler As Object
#End Region

#Region "Methods"
    Public Shared Function Save(settingName As ISettings.SettingsEnum, settingValue As Boolean) As Boolean
        Return _save({settingName, CInt(settingValue)})
    End Function

    Public Shared Function Save(settingEntry As SettingEntry) As Boolean
        Dim res As Boolean = False
        res = _save({ISettings.SettingsEnum.CacadoreEnabled, CInt(settingEntry.Enabled)})
        If Not res Then Return False
        res = _save({ISettings.SettingsEnum.HistoryEnabled, CInt(settingEntry.HistoryEnabled)})
        If Not res Then Return False
        res = _save({ISettings.SettingsEnum.TriggersEnabled, CInt(settingEntry.TriggersEnabled)})
        If Not res Then Return False
        Return True
    End Function

    ''' <summary>
    ''' Disable Cacadore (all tasks disabled, all triggers disables.)
    ''' </summary>
    Public Sub ShutdownCacadore() Implements ISettings.ShutdownCacadore
        Enabled = False
    End Sub

#Region "Privates"
    Private Shared Function loadSetting(ParamArray args() As String) As Boolean
        Try
            Return _processOperation(args, QueryOperationDlg)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            _processExcpetion(ex)
            Return Nothing
        End Try
    End Function

    Private Shared Sub loadAllSettings()
        isTriggersEnabled = loadSetting(CStr(ISettings.SettingsEnum.TriggersEnabled))
        isHistoryEnabled = loadSetting(CStr(ISettings.SettingsEnum.HistoryEnabled))
        isEnabledValue = loadSetting(CStr(ISettings.SettingsEnum.CacadoreEnabled))
        IsLoaded = True
    End Sub

    Private Shared Function _save(args() As String) As Boolean
        Try
            Return _processOperation(args, UpdateOperationDlg)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            _processExcpetion(ex)
            Return Nothing
        End Try
    End Function

    'Process Crud operations
    Private Shared Function _processOperation(args() As String, presistanceOperation As [Delegate]) As Boolean
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(New Settings, args, presistanceOperation, My.Resources.SetNotPresisted))
        End If
    End Function

    Private Shared Sub _processExcpetion(ex As Exception)
        If ex.Message.Contains(My.Resources.NoObjectFound) Then
            _privateRaiseErrorUp(My.Resources.SetNotPresisted, ex, False)
        Else
            _privateRaiseErrorUp("", ex, True)
        End If
    End Sub
#End Region
#End Region

#Region "Events"
    'Shared event is a powerful techniques 
    'that all events in all instances will 
    'be triggered.
    Public Shared Event CacadoreEnabledChanged(ByVal enabled As Boolean)
    Public Shared Event TriggersEnabledChanged(ByVal enabled As Boolean)
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