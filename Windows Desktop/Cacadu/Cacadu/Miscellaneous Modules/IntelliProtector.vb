REM ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
REM //
REM //		Copyright 2010. IntelliProtector.com
REM //		Web site: http://intelliprotector.com
REM //		Version: v2.6
REM //
REM ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Option Strict Off

Imports System
Imports System.Text
Imports System.Runtime.InteropServices

Namespace netprotector.api
    Friend Delegate Sub dVoid_NoParams()
    Friend Delegate Sub dVoid_BoolParam(ByRef b1 As Boolean)
    Friend Delegate Sub dVoid_IntParam(ByRef i1 As Integer)
    Friend Delegate Sub dVoid_2IntParam(ByRef i1 As Integer, ByVal i2 As Integer)
    Friend Delegate Sub dVoid_5IntParams(ByRef i1 As Integer, ByRef i2 As Integer, ByRef i3 As Integer, ByRef i4 As Integer, ByRef i5 As Integer)
    Friend Delegate Sub dVoid_StringBuilderParam(<MarshalAs(UnmanagedType.LPWStr)> ByVal s1 As StringBuilder, ByVal iMaxLength As Integer)
    Friend Delegate Sub dBool_StringParam(ByRef b1 As Boolean, <MarshalAs(UnmanagedType.LPWStr)> ByVal s1 As [String], ByVal iMaxLength As Integer)
    Friend Delegate Sub dBool_StringStringParam(ByRef b1 As Boolean, ByVal s1 As [String], ByVal s2 As [String])
    Friend Delegate Sub dBool_StringOnlyParam(ByRef b1 As Boolean, ByVal s1 As [String])
End Namespace
Namespace IntelliProtectorService.attributes
    Friend Class Encrypt
        Inherits Attribute
    End Class
    Friend Class GetBuyNowLinkAttribute
        Inherits Attribute
    End Class
    Friend Class IsSoftwareProtectedAttribute
        Inherits Attribute
    End Class
    Friend Class IsSoftwareRegisteredAttribute
        Inherits Attribute
    End Class
    Public Class GetTrialDaysCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetTrialUnitsCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetTrialDaysLeftCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetTrialUnitsLeftCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetTrialLaunchesCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetTrialLaunchesLeftCountAttribute
        Inherits Attribute
    End Class
    Friend Class ShowRegistrationWindowAttribute
        Inherits Attribute
    End Class
    Friend Class IsTrialElapsedAttribute
        Inherits Attribute
    End Class
    Friend Class GetRenewalPurchaseLinkAttribute
        Inherits Attribute
    End Class
    Friend Class GetLicenseTypeAttribute
        Inherits Attribute
    End Class
    Friend Class GetLicenseExpirationDaysCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetLicenseExpirationDaysLeftCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetSupportExpirationDaysCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetSupportExpirationDaysLeftCountAttribute
        Inherits Attribute
    End Class
    Friend Class GetCurrentProductVersionAttribute
        Inherits Attribute
    End Class
    Friend Class GetCustomerNameAttribute
        Inherits Attribute
    End Class
    Friend Class GetCustomerEMailAttribute
        Inherits Attribute
    End Class
    Friend Class GetLicenseCodeAttribute
        Inherits Attribute
    End Class
    Friend Class RegisterSoftwareAttribute
        Inherits Attribute
    End Class
    Friend Class RenewLicenseCodeAttribute
        Inherits Attribute
    End Class
    Friend Class GetCurrentActivationDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetCurrentRegistrationDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetFirstRegistrationDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetOrderDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetLicenseExpirationDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetProtectionDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetSupportExpirationDateAttribute
        Inherits Attribute
    End Class
    Friend Class GetSupportExpirationProductVersionAttribute
        Inherits Attribute
    End Class
    Friend Class GetIntelliProtectorVersionAttribute
        Inherits Attribute
    End Class
    Friend Class CreateRegistrationRequestCertificateAttribute
        Inherits Attribute
    End Class
    Friend Class UseRegistrationResponseCertificateAttribute
        Inherits Attribute
    End Class
    Friend Class WaitForInitializationAttribute
        Inherits Attribute
    End Class
    Friend Class SystemFunction1Attribute
        Inherits Attribute
    End Class
End Namespace
Namespace IntelliProtectorService
    Friend Class IntelliProtector
        Public Shared Function Version_2_6() As Double
            Return 2.6F
        End Function
#If DEBUG Then
        Public Shared Function IsInDebugMode() As Boolean
            Return True
        End Function
#End If

        Public Enum enUnitDimension
            eudMinutes
            eudHours
            eudDays
            eudWeeks
            eudMonths
        End Enum

        Private Class Init32
            <DllImport("JitHookCore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="Init")> _
            Public Shared Sub Init()
            End Sub
        End Class

        Private Class Init64
            <DllImport("JitHookCoreX64.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="Init")> _
            Public Shared Sub Init()
            End Sub
        End Class

        Public Shared Function Init() As Boolean
            Try
                Dim tryLoadX64 As Boolean
                tryLoadX64 = False

                Try
                    Init32.Init()
                Catch generatedExceptionName As BadImageFormatException
                    tryLoadX64 = True
                End Try

                If tryLoadX64 Then
                    Try
                        Init64.Init()
                    Catch generatedExceptionName As BadImageFormatException
                        Return False
                    End Try
                End If

                Return True
            Catch ex As Exception
                Return True
            End Try
        End Function

        <attributes.GetBuyNowLinkAttribute()> _
        Public Shared Function GetBuyNowLink() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetBuyNowLinkTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetBuyNowLinkTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.IsSoftwareProtected()> _
        Public Shared Function IsSoftwareProtected() As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False}
            Dim tp As Type = GetType(netprotector.api.dVoid_BoolParam)

            IsSoftwareProtectedTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub IsSoftwareProtectedTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = False
        End Sub
        <attributes.IsSoftwareRegistered()> _
        Public Shared Function IsSoftwareRegistered() As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False}
            Dim tp As Type = GetType(netprotector.api.dVoid_BoolParam)

            IsSoftwareRegisteredTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub IsSoftwareRegisteredTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = False
        End Sub
        <attributes.GetTrialDaysCount()> _
        Public Shared Function GetTrialDaysCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)

            GetTrialDaysCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetTrialDaysCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetTrialUnitsCount()> _
        Public Shared Function GetTrialUnitsCount(ByVal dimensions As Integer) As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, dimensions}
            Dim tp As Type = GetType(netprotector.api.dVoid_2IntParam)
            GetTrialUnitsCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetTrialUnitsCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetTrialDaysLeftCount()> _
        Public Shared Function GetTrialDaysLeftCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetTrialDaysLeftCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetTrialDaysLeftCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetTrialUnitsLeftCount()> _
        Public Shared Function GetTrialUnitsLeftCount(ByVal dimension As Integer) As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_2IntParam)
            GetTrialUnitsLeftCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetTrialUnitsLeftCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetTrialLaunchesCount()> _
        Public Shared Function GetTrialLaunchesCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetTrialLaunchesCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetTrialLaunchesCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetTrialLaunchesLeftCount()> _
        Public Shared Function GetTrialLaunchesLeftCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetTrialLaunchesLeftCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetTrialLaunchesLeftCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.ShowRegistrationWindow()> _
        Public Shared Sub ShowRegistrationWindow()
            Dim address As UInt64 = 0
            Dim tp As Type = GetType(netprotector.api.dVoid_NoParams)
            ShowRegistrationWindowTest(address, Nothing, tp)
        End Sub
        <attributes.Encrypt()> _
        Public Shared Sub ShowRegistrationWindowTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
        End Sub
        <attributes.IsTrialElapsed()> _
        Public Shared Function IsTrialElapsed() As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False}
            Dim tp As Type = GetType(netprotector.api.dVoid_BoolParam)

            IsTrialElapsedTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub IsTrialElapsedTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = False
        End Sub
        <attributes.GetRenewalPurchaseLink()> _
        Public Shared Function GetRenewalPurchaseLink() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetRenewalPurchaseLinkTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetRenewalPurchaseLinkTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.GetLicenseType()> _
        Public Shared Function GetLicenseType() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetLicenseTypeTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetLicenseTypeTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetLicenseExpirationDaysCount()> _
        Public Shared Function GetLicenseExpirationDaysCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetLicenseExpirationDaysCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetLicenseExpirationDaysCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetLicenseExpirationDaysLeftCount()> _
        Public Shared Function GetLicenseExpirationDaysLeftCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetLicenseExpirationDaysLeftTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetLicenseExpirationDaysLeftTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetSupportExpirationDaysCount()> _
        Public Shared Function GetSupportExpirationDaysCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetSupportExpirationDaysCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetSupportExpirationDaysCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetSupportExpirationDaysLeftCount()> _
        Public Shared Function GetSupportExpirationDaysLeftCount() As Integer
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0}
            Dim tp As Type = GetType(netprotector.api.dVoid_IntParam)
            GetSupportExpirationDaysLeftCountTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Public Shared Sub GetSupportExpirationDaysLeftCountTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
        End Sub
        <attributes.GetCurrentProductVersion()> _
        Public Shared Function GetCurrentProductVersion() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetCurrentProductVersionTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetCurrentProductVersionTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.GetCustomerName()> _
        Public Shared Function GetCustomerName() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetCustomerNameTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetCustomerNameTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.GetCustomerEMail()> _
        Public Shared Function GetCustomerEMail() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetCustomerEMailTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetCustomerEMailTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.GetLicenseCode()> _
        Public Shared Function GetLicenseCode() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetLicenseCodeTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetLicenseCodeTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.RegisterSoftware()> _
        Public Shared Function RegisterSoftware(ByVal licenseKey As String) As Boolean
            Dim address As UInt64 = &HDEADBEAFDEADBEAFUL
            Dim obj As Object() = New Object() {False, licenseKey, licenseKey.Length}
            Dim tp As Type = GetType(netprotector.api.dBool_StringParam)

            RegisterSoftwareTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub RegisterSoftwareTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = False
        End Sub
        <attributes.RenewLicenseCode()> _
        Public Shared Function RenewLicenseCode(ByVal renewalCode As String) As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False, renewalCode, renewalCode.Length}
            Dim tp As Type = GetType(netprotector.api.dBool_StringParam)

            RenewLicenseCodeTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub RenewLicenseCodeTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = False
        End Sub
        <attributes.GetCurrentActivationDate()> _
        Public Shared Function GetCurrentActivationDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetCurrentActivationDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetCurrentActivationDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetCurrentRegistrationDate()> _
        Public Shared Function GetCurrentRegistrationDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetCurrentRegistrationDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetCurrentRegistrationDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetFirstRegistrationDate()> _
        Public Shared Function GetFirstRegistrationDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetFirstRegistrationDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetFirstRegistrationDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetOrderDate()> _
        Public Shared Function GetOrderDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetOrderDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetOrderDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetLicenseExpirationDate()> _
        Public Shared Function GetLicenseExpirationDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetLicenseExpirationDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetLicenseExpirationDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetProtectionDate()> _
        Public Shared Function GetProtectionDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetProtectionDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetProtectionDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetSupportExpirationDate()> _
        Public Shared Function GetSupportExpirationDate() As DateTime
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {0, 0, 0, 0, 0}
            Dim tp As Type = GetType(netprotector.api.dVoid_5IntParams)

            GetSupportExpirationDateTest(address, obj, tp)
            If obj(0) < 0 Then
                Return DateTime.MinValue
            End If
            Return New DateTime(obj(0), obj(1), obj(2), obj(3), obj(4), 0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetSupportExpirationDateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = -1
            obj(1) = -1
            obj(2) = -1
            obj(3) = -1
            obj(4) = -1
        End Sub
        <attributes.GetSupportExpirationProductVersion()> _
        Public Shared Function GetSupportExpirationProductVersion() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetSupportExpirationProductVersionTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetSupportExpirationProductVersionTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub
        <attributes.GetIntelliProtectorVersion()> _
        Public Shared Function GetIntelliProtectorVersion() As [String]
            Dim address As UInt64 = 0
            Dim ciMinLength As Integer = 0
            Dim ciMaxLength As Integer = 1024
            Dim rValue As New StringBuilder(ciMaxLength)
            Dim obj As Object() = New Object() {rValue, ciMaxLength - ciMinLength}
            Dim tp As Type = GetType(netprotector.api.dVoid_StringBuilderParam)

            GetIntelliProtectorVersionTest(address, obj, tp)
            Return rValue.ToString()
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub GetIntelliProtectorVersionTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            Dim strLink As String = ""
            If DirectCast(obj(1), Integer) >= strLink.Length Then
                DirectCast(obj(0), StringBuilder).Append(strLink)
            End If
        End Sub

        <attributes.CreateRegistrationRequestCertificateAttribute()> _
        Public Shared Function CreateRegistrationRequestCertificate(ByVal path As String, ByVal licenseKey As String) As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False, path, licenseKey}
            Dim tp As Type = GetType(netprotector.api.dBool_StringStringParam)

            CreateRegistrationRequestCertificateTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub CreateRegistrationRequestCertificateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = True
        End Sub
        <attributes.UseRegistrationResponseCertificateAttribute()> _
        Public Shared Function UseRegistrationResponseCertificate(ByVal path As String) As Boolean
            Dim address As UInt64 = 0
            Dim obj As Object() = New Object() {False, path}
            Dim tp As Type = GetType(netprotector.api.dBool_StringOnlyParam)

            UseRegistrationResponseCertificateTest(address, obj, tp)
            Return obj(0)
        End Function
        <attributes.Encrypt()> _
        Protected Shared Sub UseRegistrationResponseCertificateTest(ByVal address As UInt64, ByVal obj As Object(), ByVal delegateType As Type)
            obj(0) = True
        End Sub
        <attributes.WaitForInitializationAttribute()> _
        Public Shared Sub WaitForInitialization()
            Dim address As UInt64 = 0
            Dim tp As Type = GetType(netprotector.api.dVoid_NoParams)
            ShowRegistrationWindowTest(address, Nothing, tp)
        End Sub
        <attributes.SystemFunction1()> _
        Protected Shared Sub SystemFunction1(ByVal address As UInt64, ByVal parameters As Object(), ByVal delegateType As Type)
            Dim delegateFunc As [Delegate] = Marshal.GetDelegateForFunctionPointer(address, delegateType)
            delegateFunc.DynamicInvoke(parameters)
        End Sub
    End Class
End Namespace

