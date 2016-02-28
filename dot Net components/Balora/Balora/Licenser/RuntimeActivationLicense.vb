Imports System.ComponentModel
'هنا براحتك .. دي الرخصة نفسها .. عايز تحط فيها اي معلومات مشفرة تبع الترخيص براحتك برضه
Namespace Lixcer 'Licenser
    Friend Class RAL 'RuntimeActivationLicense -- rename again to the old name whenever you buy babelfor .net
        Inherits License
        Public Overrides ReadOnly Property LicenseKey As String
            Get
                'Just return nothing until you implement your lic system.
                Return Nothing
            End Get
        End Property

        Public Overrides Sub Dispose()

        End Sub
    End Class
End Namespace

'Example
'Friend Class RuntimeRegistryLicense
'	Inherits License

'	Private type As type = Nothing

'	Public Sub New(ByVal type As type)
'		Me.type = type
'	End Sub

'	Public Overrides ReadOnly Property LicenseKey() As String
'		Get
'			Return type.GUID.ToString()
'		End Get
'	End Property

'	Public Overrides Sub Dispose()

'	End Sub

'End Class