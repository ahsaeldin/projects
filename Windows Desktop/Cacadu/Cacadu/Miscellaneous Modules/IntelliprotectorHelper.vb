Module IntelliprotectorHelper
    '*.How to play with the free version in the future.
    '1.Inside your code you will use "GetLicenseType" to protect paid areas in your code 
    '2.When you finish at publishing a new version, branch it to "E:\Cacadu\Code\Branches"
    '3.Branch the free version to Free Versions "E:\Cacadu\Code\Branches\Free Versions"
    '4.Branch the lite version to "E:\Cacadu\Code\Branches\Paid Versions\Lite Versions"

    Sub ShowRegistrationWindow()
        IntelliProtectorService.IntelliProtector.ShowRegistrationWindow()
    End Sub

    Function IsCacaduRegistered() As Boolean
        Return IntelliProtectorService.IntelliProtector.IsSoftwareRegistered
    End Function

    Function GetRemainingDays() As Integer
        Return IntelliProtectorService.IntelliProtector.GetTrialDaysLeftCount
    End Function

    Function GetCustomerName() As String
        Return IntelliProtectorService.IntelliProtector.GetCustomerName
    End Function

    Function GetLicenseType() As Integer
        'Returns -1 if license code is not available (software is not purchased, trial period).
        Return IntelliProtectorService.IntelliProtector.GetLicenseType
    End Function
End Module
