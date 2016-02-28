Module ScraperSettings
    Property LogFilePath As String = "log.txt"
    Property TargetWebSiteName As String = "Tmart"
    Property TargetWebsiteUrl As String = "http://www.tmart.com/"
    Property SiteMapURL As String = "http://www.tmart.com/site_map.html"
    Property CatList As New Dictionary(Of String, String)
    Property ExchangeRateFromEuroToUsd As Double = 0
End Module
