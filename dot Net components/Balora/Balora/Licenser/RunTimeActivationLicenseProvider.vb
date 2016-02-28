Imports System.ComponentModel

Namespace Lixcer
    Friend Class RTALP ' RunTimeActivationLicenseProvider -- rename again to the old name whenever you buy babelfor .net
        Inherits LicenseProvider
        Public Overrides Function GetLicense(ByVal context As System.ComponentModel.LicenseContext, ByVal type As System.Type, ByVal instance As Object, ByVal allowExceptions As Boolean) As System.ComponentModel.License
            'Just return nothing until you implement your lic system.

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '+A Future vision of how to combine both intelliProtector with Licenser. 
            'Hi,

            'You should protect it the same way as you protect the ordinary executable. 
            'Just make sure that the IntelliProtector.Init() is called before entering 
            'any method with [Encrypt] attribute. Also the Init method must be called 
            'before any call to the IntelliProtector API. That’s all.

            'Let me know if you have more question.

            'Serge(Levin)
            'Dim initValue As Boolean = IntelliProtectorService.IntelliProtector.Init()
            'If initValue Then
            'Return New RuntimeActivationLicense()
            'Else
            'Return Nothing
            'End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '++woooooooooooooowz, this is my awesome check details
            '+The Story:-
            'Simply I want Balora or Cacadore to run only via Cacadu, in other words...
            'I don't want a developer to simply import and call them. The idea here
            'is to check the equity of a md5 value of a GUID specified here by me
            'And Cacadu will be the only who knows this vlaue and pass its hash to me.
            If Not Settings.CAG Is Nothing Then
                'This is the hash of {38D4C0D0-8E45-40E6-B40A-67AAFFE264E1} directly 7E3434FCB6CFD72CE0C22F8DDB9CBB9D
                'so I can avoid hard coding directly in balora
                Dim md5_1 = "7E3434FCB6CFD72CE0C22F8DDB9CBB9D" 'Serializer.BinarySerializer.MD5SerializedObject("{38D4C0D0-8E45-40E6-B40A-67AAFFE264E1}")
                Dim md5_2 = Balora.Settings.CAG
                If Object.Equals(md5_1, md5_2) Then
                    'Clear the value in case of a مختلسي النظر للقيم في الذاكرة
                    Settings.CAG = Nothing
                    Return New RAL()
                Else
                    'If CA property isn't equal to assemblyThatCallsThisLibrary then no license here ya m3lm.
                    Return (Nothing)
                End If
            Else
                'If CA property isn't set then no license here ya m3lm.
                Return (Nothing)
            End If

        End Function
    End Class
End Namespace

'Example
'Friend Class RegistryLicenseProvider
'	Inherits LicenseProvider

'	Public Overloads Overrides Function GetLicense( _
'	ByVal context As LicenseContext, _
'	ByVal type As Type, _
'	ByVal instance As Object, _
'	ByVal allowExceptions As Boolean) As License
'		Dim licenseKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\\Acme\\HostKeys")
'		If context.UsageMode = LicenseUsageMode.Runtime Then
'			If Not licenseKey Is Nothing Then
'				Dim strLic As String = CType(licenseKey.GetValue(type.GUID.ToString()), String)
'				If Not strLic Is Nothing Then
'					If String.Compare("Installed", strLic, False) = 0 Then
'						GetLicense = New RuntimeRegistryLicense(type)
'						Exit Function
'					End If
'				End If

'				If allowExceptions = True Then
'					Throw New LicenseException(type, instance, "Your license is invalid")
'				End If

'				GetLicense = Nothing
'			End If
'		Else
'			GetLicense = New DesigntimeRegistryLicense(type)
'		End If
'	End Function
'End Class