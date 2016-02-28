Imports System.IO
Imports System.Text
Imports MyScraper.MainForm
Imports System.Text.RegularExpressions

Friend Class Scraper
#Region "Properties"
    Shared Property IsPaused As Boolean = False
    Shared Property IsCancelled As Boolean = False
    'Holds site map html list.
    Shared Property SiteMapElement As HtmlElement
    'Hold items that is out of stock.
    'Shared Property OutOfStockItems As New Dictionary(Of String, String)
    'Holds non USA pieces items.
    'Shared Property excludedWarehousesItems As New Dictionary(Of String, String)
    'Holds items URLs in the keys, and corresponding department URL in value.
    Shared Property ProductsPagesList As New List(Of String)
    'Holds department name in the keys, and Department objects in values.
    Shared Property DepartmentDict As New Dictionary(Of String, Department)
    'Holds scraped products url & data scraped.
    Shared Property ScrapedProdcuts As New List(Of String)
#End Region

#Region "Methods"
    Shared Sub ScrapeDepartments(document As HtmlDocument,
                                 siteMapListId As String,
                                 singleListId As String)
        Dim lastLevel0Node As TreeNode = Nothing
        Dim lastLevel1Node As TreeNode = Nothing
        Dim lastLevel2Node As TreeNode = Nothing
        Dim lastLevel3Node As TreeNode = Nothing

        Dim isEnough As Boolean = False
        SiteMapElement = document.GetElementById(siteMapListId)
        For Each ele As HtmlElement In SiteMapElement.All

            If ele.InnerText.Contains("Home, Office & Garden") Then
                Console.WriteLine("Asdfasdf")
            End If

            Dim departName As String = ele.InnerText
            Dim departUrl As String = ele.GetAttribute("href")
            If Strings.Right(departUrl, 1) = "/" Then
                departUrl = Strings.Left(departUrl, Len(departUrl) - 1) + "-301list/"
            Else
                departUrl = departUrl + "-301list/"
            End If

            If ele.TagName = "A" AndAlso ele.Parent.OuterHtml.Contains("level0") Then
                If Not MainForm.ExcludedDepartmentsDict.Contains(departName) Then
                    lastLevel0Node = MainForm.DepartmentsTreeView.Nodes.Add(departName, departName)
                    lastLevel0Node.Tag = departUrl
                    DepartmentDict.Add(departName, CreateDepartment(departName, departUrl, 0))
                    MainForm.TreeNodesList.Add(lastLevel0Node)
                End If
            ElseIf ele.TagName = "A" AndAlso ele.Parent.OuterHtml.Contains("level1") Then
                Dim name As String = lastLevel0Node.Text + "/" + departName
                If Not MainForm.ExcludedDepartmentsDict.Contains(name) Then
                    lastLevel1Node = lastLevel0Node.Nodes.Add(name, departName)
                    lastLevel1Node.Tag = departUrl
                    DepartmentDict.Add(name, CreateDepartment(departName, departUrl, 1))
                    MainForm.TreeNodesList.Add(lastLevel1Node)
                End If
            ElseIf ele.TagName = "A" AndAlso ele.Parent.OuterHtml.Contains("level2") Then
                Dim path As String = lastLevel0Node.Text + "/" + lastLevel1Node.Text + "/" + departName
                If Not MainForm.ExcludedDepartmentsDict.Contains(path) Then
                    lastLevel2Node = lastLevel1Node.Nodes.Add(path, departName)
                    lastLevel2Node.Tag = departUrl
                    DepartmentDict.Add(path, CreateDepartment(departName, departUrl, 2))
                    MainForm.TreeNodesList.Add(lastLevel2Node)
                End If
            ElseIf ele.TagName = "A" AndAlso ele.Parent.OuterHtml.Contains("level3") Then
                Dim path As String = lastLevel0Node.Text + "/" + lastLevel1Node.Text + "/" + lastLevel2Node.Text + "/" + departName
                If Not MainForm.ExcludedDepartmentsDict.Contains(path) Then
                    lastLevel3Node = lastLevel2Node.Nodes.Add(path, departName)
                    lastLevel3Node.Tag = departUrl
                    DepartmentDict.Add(path, CreateDepartment(departName, departUrl, 3))
                    MainForm.TreeNodesList.Add(lastLevel3Node)
                End If
            End If

            'Added for a strange bug sent by Daniel that lets more of category added to the last sub category.
            'May be I fix that in the future.
            If departName = "Sportswear" Then
                ' Exit For
            End If
            Application.DoEvents()
        Next

        If IO.File.Exists("Cats.txt") Then IO.File.Delete("Cats.txt")
        If IO.File.Exists("CatsTree.txt") Then IO.File.Delete("CatsTree.txt")
        For Each depart In DepartmentDict
            GetAppFolderPath()

            ' My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "Cats.txt", depart.Value.DepartmentName + vbNewLine, True)
            'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "CatsTree.txt", depart.Key + vbNewLine, True)
        Next
    End Sub

    Shared Function CreateDepartment(name As String,
                                     url As String,
                                     level As Integer) As Department
        Dim depart As New Department
        With depart
            .DepartmentName = name
            .DepartmentUrl = url
            .Level = level
        End With
        Return depart
    End Function

    Shared Function DownloadPageSource(url As String) As String
        Dim webClient As New System.Net.WebClient
        Try
            webClient.Encoding = Encoding.UTF8
            Dim result As String = webClient.DownloadString(url)
            Return result
        Catch ex As Exception
            'ChangeStatus("DownloadPageSource::" + ex.Message)
            Return False
        End Try
    End Function

    Shared Function CreateHtmlDocument(pageSource As String) As HtmlAgilityPack.HtmlDocument
        Dim doc As New HtmlAgilityPack.HtmlDocument
        Try
            doc.Load(New MemoryStream(Encoding.UTF8.GetBytes(pageSource)))
            Return doc
        Catch ex As Exception
            LogState(ex.Message + vbNewLine + ex.StackTrace)
            Return Nothing
        End Try
    End Function

    Shared Function GetSearchPagesCount(doc As HtmlAgilityPack.HtmlDocument) As Integer
        Dim searchNodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//*[@id=%productsListingListingTopLinks%]".Replace("%", Chr(34)))
        If Not IsNothing(searchNodes) Then
            Dim pagesCountQuery = From ele As HtmlAgilityPack.HtmlNode In searchNodes.Nodes
                                  Where ele.Name = "a"
            If pagesCountQuery.Count = 0 Then Return 1 Else Return pagesCountQuery.Count
        Else
            Return 1
        End If
    End Function

    Sub ScrapeDepartment(url As String, departmentName As String, Optional workerThread As System.ComponentModel.BackgroundWorker = Nothing)
        '#Add a switch here to check for Auto Save.
        '#We need to save ProductsPagesDict & and the last key scraped in it.
        '#We want to start from the last point.
        '#We need to bypass the first loop
        '#We need to save iOffer and Magneto files frequently

        'Dim pass As Boolean = True
        'If Not pass OrElse Not departmentName = "iPhone Accessories" Then Exit Sub
        'pass = True

        Dim urlsCount As Integer = 0
        Dim listingPagesToBeScraped As New List(Of String)

        Dim searchPageSource As String = DownloadPageSource(url)
        Dim doc As HtmlAgilityPack.HtmlDocument = CreateHtmlDocument(searchPageSource)
        Dim pagesCount As Integer = GetSearchPagesCount(doc)
        doc = Nothing
        'LogState("pagesCountQuery = " + pagesCount.ToString + " in URL::" + url)


        For i As Integer = 1 To pagesCount
            Try
                If IsCancelled Then
                    IsCancelled = False
                    Exit For
                End If

                Do While IsPaused
                    Application.DoEvents()
                Loop
                Dim token As String = "30" + i.ToString + "list"
                listingPagesToBeScraped.Add(url.Replace("301list", token))

                If Not IsNothing(workerThread) AndAlso workerThread.IsBusy Then
                    workerThread.ReportProgress(i, "Scraping items Urls from " + departmentName + String.Format(" page: {0} of {1}", i, pagesCount))
                End If

                If i = 1 Then
                    'Scrape first page...
                    urlsCount = ScrapeProductsUrl(searchPageSource, departmentName, url)
                Else
                    Dim nextUrl As String = listingPagesToBeScraped(i - 1)
                    searchPageSource = DownloadPageSource(nextUrl)
                    urlsCount = urlsCount + ScrapeProductsUrl(searchPageSource, departmentName, nextUrl)
                End If
            Catch ex As Exception
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next

        If Not IsNothing(workerThread) AndAlso workerThread.IsBusy Then
            workerThread.ReportProgress(0, urlsCount.ToString + " items Urls collected for department '" + departmentName + "'")
        End If

        Dim urlNo As Integer = 0
        'If IO.File.Exists("Excluded Warehouses Items.txt") Then IO.File.Delete("Excluded Warehouses Items.txt")

        For Each str As String In ProductsPagesList
            Try
                If IsCancelled Then
                    IsCancelled = False
                    Exit For
                End If
                Do While IsPaused
                    Application.DoEvents()
                Loop
                urlNo = urlNo + 1
                If Not IsNothing(workerThread) AndAlso workerThread.IsBusy Then
                    workerThread.ReportProgress(0, String.Format("item {0} of {1} items in department {2}", urlNo, ProductsPagesList.Count, departmentName))
                End If

                Dim productPageSource As String = DownloadPageSource(str)

                If productPageSource.Contains("4GB 1.8% MP4 with FM Function Black".Replace("%", Chr(34))) Then
                    Console.WriteLine("asdfasdf")
                End If

                Dim excludedReason As String = vbNullString

                If ScrapedProdcuts.Contains(str) Then
                    Continue For
                End If


                'Remove Chinese Characters first.
                Dim regEx1 As New Regex("[\u4e00-\u9fa5]+", RegexOptions.IgnoreCase)
                Dim matchesCol As MatchCollection = regEx1.Matches(productPageSource)
                'If successful, write the group.
                If (matchesCol.Count > 0) Then
                    For Each matchy As Match In matchesCol
                        productPageSource = productPageSource.Replace(matchy.Value, "")
                    Next
                    'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "Description with Chinese characters.txt", kvp.Key + vbNewLine, True)
                End If

                Dim productData As BusinessItem = ScrapeProductPage(CreateHtmlDocument(productPageSource),
                                                                    str,
                                                                    productPageSource,
                                                                    "disable for now",
                                                                    excludedReason)

                If Not IsNothing(productData) Then
                    'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "page" + ".txt", productPageSource, False)
                    ScrapedProdcuts.Add(str)
                    Dim magentoOutput As New MagentoOutput
                    With magentoOutput
                        .sku = productData.SKUNumber
                        If Not IsNothing(productData.Price) AndAlso productData.Price.Contains("&euro;") Then
                            productData.Price = productData.Price.Replace("&euro;", "")
                            Dim convertedPrice As Double = CDbl(productData.Price) * ExchangeRateFromEuroToUsd
                            productData.Price = convertedPrice.ToString("0.00")
                        End If

                        If My.Settings.EnablePriceRange Then
                            If Not CDbl(productData.Price) >= My.Settings.MinPrice OrElse Not CDbl(productData.Price) <= My.Settings.MaxPrice Then
                                Continue For
                            End If
                        End If

                        'LogState("+++++++++++++++++++++++++++++++")
                        'LogState("Price URL::" + productData.ProductUrl)
                        'LogState("Price Type::" + My.Settings.PriceChangingType.ToString)
                        'LogState("Scraped Price::" + productData.Price)
                        If My.Settings.PriceChangingType = 0 Then
                            Dim percentageValue As Double = (1 + My.Settings.PriceChangeByPercentage / 100)
                            LogState("percentageValue::" + percentageValue.ToString)
                            .price = (CDbl(productData.Price) * percentageValue).ToString("0.00")  'السعر مضروبا في النسبة
                        ElseIf My.Settings.PriceChangingType = 1 Then
                            .price = CDbl(productData.Price) + CDbl(My.Settings.PriceChangeByValue)
                        End If

                        If My.Settings.ShippingCostIsSet Then
                            Dim price As Double = My.Settings.ShippingCost
                            If price > 0 Then
                                .price = .price + price
                            End If
                        End If
                        'LogState("Modified Price::" + .price)
                        'LogState("Modified Magento Price::" + .price)
                        'Dim dateFileName = Utils.GetDateTimeFileName(".html")
                        'LogState("HTML File Name::" + dateFileName)
                        'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + dateFileName, productPageSource, False)

                        productData.Price = .price
                   
                        .description = productData.ProductDescriptionHtml
                        .short_description = Strings.Left(productData.ProductDescriptionText, InStr(productData.ProductDescriptionText, "."))
                        .name = productData.ProductTitle
                        .image = productData.ProdcutLargeImageUrl
                        .small_image = productData.ProdcutLargeImageUrl
                        .thumbnail = productData.ProdcutLargeImageUrl
                        .categories = productData.Categories
                        .media_gallery = Strings.Join(productData.Images.ToArray, ";")
                        If excludedReason = "Out Of Stock" Then
                            .qty = 0
                            .min_qty = 0
                            .is_in_stock = 0
                        End If
                    End With

                    Dim outputsArr As New List(Of Object)
                    outputsArr.Add(productData)
                    outputsArr.Add(magentoOutput)
                    If excludedReason <> "Out Of Stock" Then
                        Dim iOfferOutput As New iOfferOutput
                        With iOfferOutput
                            If CatList.ContainsKey(magentoOutput.categories) Then
                                .Category = CatList(magentoOutput.categories)
                            Else
                                ' LogState("Categories::" & magentoOutput.categories + " not found in the cat list.")
                                '    My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "N-A categories", "Categories::" & magentoOutput.categories + " not found in the cat list.", True)
                                .Category = "N/A"
                            End If
                            .Title = magentoOutput.name
                            .Description = magentoOutput.description
                            .Price = magentoOutput.price
                            .Quantity = "25"
                            .Shipping = "0"
                            .Image1 = magentoOutput.image
                            If magentoOutput.media_gallery <> "" Then
                                Dim imagesArr() As String = magentoOutput.media_gallery.Split(";")
                                If imagesArr.Count > 1 Then
                                    .Image2 = imagesArr(1)
                                    If imagesArr.Count > 2 Then
                                        .Image3 = imagesArr(2)
                                    End If
                                End If
                            End If

                            If productData.Keywords <> vbNullString Then
                                Dim keywordsArr() As String = productData.Keywords.Split(";")
                                .Keyword1 = keywordsArr(0)
                                If keywordsArr.Count > 1 Then .Keyword2 = keywordsArr(1)
                                If keywordsArr.Count > 2 Then .Keyword3 = keywordsArr(2)
                                If keywordsArr.Count > 3 Then .Keyword4 = keywordsArr(3)
                            End If
                        End With
                        outputsArr.Add(iOfferOutput)
                    End If

                    If Not IsNothing(workerThread) AndAlso workerThread.IsBusy Then workerThread.ReportProgress(urlNo, outputsArr)
                Else
                    'LogState("productData is nothing::" + kvp.Key + "::" + kvp.Value)
                    'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "Excluded Warehouses Items.txt", excludedReason + " Item::" + kvp.Key + vbNewLine, True)
                End If
            Catch ex As Exception
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next
    End Sub

    Shared Function ScrapeProductsUrl(pageText As String,
                                      departmentName As String,
                                      departmentUrl As String) As Integer
        Dim regex As Regex = New Regex("<div class=%itemTitle%><a href=%(.+)%>(.+)</a></div>".Replace("%", Chr(34)))
        Dim [matches] As MatchCollection = regex.Matches(pageText)
        For Each [group] As Match In [matches]
            If Not ProductsPagesList.Contains([group].Groups(1).Value) Then
                ProductsPagesList.Add([group].Groups(1).Value)
            Else
                Dim url As String = [group].Groups(1).Value
                Dim firstDepartMent As String = ProductsPagesList(url)
                Dim duplicatedIn As String = departmentName
                'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "duplicated-Item.txt", "Url::" + url + ", found in::" + firstDepartMent + ", duplicated In::" + duplicatedIn + vbNewLine, True)
            End If
        Next
        Return [matches].Count
    End Function

    Shared Function ScrapeProductPage(pageDocument As HtmlAgilityPack.HtmlDocument,
                                      pageUrl As String,
                                      pageText As String,
                                      departmentUrl As String,
                                      ByRef excludedType As String) As BusinessItem
        If pageUrl = "http://www.tmart.com/16GB-2.8-LCD-Screen-MP5-Player-with-Digital-Camera-White_p135213.html" Then
            Console.WriteLine("asdfadf")
        End If
        Dim busItem As New BusinessItem
        If pageText.Contains("this item is temporarily out of stock") Then
            'If Not OutOfStockItems.ContainsKey(pageUrl) Then OutOfStockItems.Add(pageUrl, departmentUrl)
            excludedType = "Out Of Stock"
            'Return Nothing
        ElseIf My.Settings.PriceType = "US" AndAlso Not pageText.Contains("United States Warehouse Price") Then
            'If Not excludedWarehousesItems.ContainsKey(pageUrl) Then excludedWarehousesItems.Add(pageUrl, departmentUrl)
            excludedType = "Non USA Priced"
            Return Nothing
        ElseIf My.Settings.PriceType = "UK" AndAlso Not pageText.Contains("United Kingdom Warehouse Price") Then
            'If Not excludedWarehousesItems.ContainsKey(pageUrl) Then excludedWarehousesItems.Add(pageUrl, departmentUrl)
            excludedType = "Non UK Warehouse"
            Return Nothing
        ElseIf My.Settings.PriceType = "HK" AndAlso Not pageText.Contains("Hong Kong Warehouse Price") Then
            'If Not excludedWarehousesItems.ContainsKey(pageUrl) Then excludedWarehousesItems.Add(pageUrl, departmentUrl)
            excludedType = "Non HK Warehouse"
            Return Nothing
        End If

        'Scraping Description
        Dim searchNodes As HtmlAgilityPack.HtmlNodeCollection = pageDocument.DocumentNode.SelectNodes("//*[@id=%tab_description%]".Replace("%", Chr(34)))
        If searchNodes.Count = 0 Then
            ChangeStatus("Error while scraping; cannot find product descrption.")
            LogState("URL::" + pageUrl)
            Return Nothing
        End If
        busItem.ProductDescriptionHtml = "<font face=%arial% size=%10%>".Replace("%", Chr(34)) + searchNodes.First.InnerHtml.Trim + "</font>"
        If busItem.ProductDescriptionHtml.Contains("href") Then
            Dim regx As New System.Text.RegularExpressions.Regex("<a.+?(href=%.+?%).+?</a>".Replace("%", Chr(34)))
            Dim mats = regx.Matches(busItem.ProductDescriptionHtml)
            For Each mat As System.Text.RegularExpressions.Match In mats
                busItem.ProductDescriptionHtml = busItem.ProductDescriptionHtml.Replace(mat.Groups(1).Value, "href=%#%".Replace("%", Chr(34)))
            Next
            'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "Description with urls items.txt", busItem.ProductUrl + vbNewLine, True)
        End If
        If busItem.ProductDescriptionHtml.Contains("<strong style=%font-size:14px;%>Extra Info</strong>".Replace("%", Chr(34))) Then
            Dim whereExtraInfo As Integer = InStr(busItem.ProductDescriptionHtml, "<strong style=%font-size:14px;%>Extra Info</strong>".Replace("%", Chr(34)))
            busItem.ProductDescriptionHtml = Strings.Left(busItem.ProductDescriptionHtml, whereExtraInfo)
        ElseIf busItem.ProductDescriptionHtml.Contains("<font color=%#07519a%>Details: </font>".Replace("%", Chr(34))) Then
            Dim whereExtraInfo As Integer = InStr(busItem.ProductDescriptionHtml, "<font color=%#07519a%>Details: </font>".Replace("%", Chr(34)))
            busItem.ProductDescriptionHtml = Strings.Left(busItem.ProductDescriptionHtml, whereExtraInfo)
        ElseIf busItem.ProductDescriptionHtml.Contains("<font color=%#07519a%>Note: </font>".Replace("%", Chr(34))) Then
            Dim whereExtraInfo As Integer = InStr(busItem.ProductDescriptionHtml, "<font color=%#07519a%>Note: </font>".Replace("%", Chr(34)))
            busItem.ProductDescriptionHtml = Strings.Left(busItem.ProductDescriptionHtml, whereExtraInfo)
        End If
        busItem.ProductDescriptionText = searchNodes.First.InnerText.Trim

        If DecriptionFooter <> vbNullString Then
            busItem.ProductDescriptionHtml = busItem.ProductDescriptionHtml + "<br />" + DecriptionFooter
            busItem.ProductDescriptionText = busItem.ProductDescriptionText + vbNewLine + DecriptionFooter
        End If

        'Scraping small image.
        searchNodes = pageDocument.DocumentNode.SelectNodes("//*[@id=%smallpic_show%]".Replace("%", Chr(34)))
        If searchNodes.Count = 0 Then
            ChangeStatus("Error while scraping; cannot find product small image.")
            LogState("URL::" + pageUrl)
            Return Nothing
        End If
        busItem.ProdcutSmallImageUrl = searchNodes.First.GetAttributeValue("src", "")
        busItem.ProdcutLargeImageUrl = busItem.ProdcutSmallImageUrl.Replace("320x320", "650x650")
        busItem.ProductUrl = pageUrl


        'Scraping small Title.
        searchNodes = pageDocument.DocumentNode.SelectNodes("//*[@id=%container%]/div[2]/div[1]/div[1]/h1/span".Replace("%", Chr(34)))
        If searchNodes.Count = 0 Then
            ChangeStatus("Error while scraping; cannot find product title.")
            LogState("URL::" + pageUrl)
            Return Nothing
        End If
        busItem.ProductTitle = searchNodes.First.InnerText.Trim

        For Each prod As String In ExcludedProducts
            If busItem.ProductTitle = prod.Trim Then
                Return Nothing
            End If
        Next

        'Scraping SKU
        Dim regEx As Regex = New Regex("SKU:.+?</div>")
        Dim [match] As Match = regEx.Match(pageText)
        If [match].Success Then
            busItem.SKUNumber = [match].Value.Replace("&nbsp;", "").Replace("</div>", "").Replace("SKU:", "")
            'busItem.ProductDescriptionHtml = busItem.ProductDescriptionHtml + String.Format("<p>SKU:{0}</p>", busItem.SKUNumber)
        End If


        'Scraping Price
        Select Case My.Settings.PriceType
            Case "US"
                regEx = New Regex("United States Warehouse Price: <span class=%redBold%>(.+?)</span>".Replace("%", Chr(34)))
            Case "UK"
                regEx = New Regex("United Kingdom Warehouse Price: <span class=%redBold%>(.+?)</span>".Replace("%", Chr(34)))
            Case "HK"
                regEx = New Regex("Hong Kong Warehouse Price: <span class=%redBold%>(.+?)</span>".Replace("%", Chr(34)))
        End Select
        [match] = regEx.Match(pageText)
        If [match].Success Then
            busItem.Price = [match].Groups(1).Value.Replace("$", "").Trim
        Else
            If excludedType = "Out Of Stock" Then
                regEx = New Regex("itemprop=%price%>(.+)</span>".Replace("%", Chr(34)))
                [match] = regEx.Match(pageText)
                If [match].Success Then
                    busItem.Price = [match].Groups(1).Value.Replace("$", "").Trim
                End If
            End If
        End If

        'Scraping Categories
        searchNodes = pageDocument.DocumentNode.SelectNodes("//*[@id=%navBreadCrumb%]".Replace("%", Chr(34)))
        If Not IsNothing(searchNodes) Then
            Dim catsQuery = From ele As HtmlAgilityPack.HtmlNode In searchNodes.Nodes
                            Where ele.Name = "a" AndAlso ele.GetAttributeValue("title", "") <> "Tmart.com"
                            Select ele.InnerText
            busItem.Categories = Strings.Join(catsQuery.ToArray, "/").Trim
        End If


        'Scraping Images Urls
        Dim imagesDict As New Dictionary(Of String, String)
        regEx = New Regex("id=%(thumb\d)% href=%(.+)% onclick=%return false%>".Replace("%", Chr(34)))
        Dim thumbsMatches = regEx.Matches(pageText)
        For Each matchy As Match In thumbsMatches
            Dim imageLink As String = matchy.Groups(2).Value.Replace("320x320", "650x650")
            If Not imagesDict.ContainsKey(matchy.Groups(2).Value) Then imagesDict.Add(imageLink, "")
        Next
        busItem.Images = imagesDict.Keys.ToList


        'Scraping Keywords
        searchNodes = pageDocument.DocumentNode.SelectNodes("/html/head/meta[11]".Replace("%", Chr(34)))
        Dim metaTagQuery = searchNodes.First.GetAttributeValue("content", "")

        'Fine tuning keywords.
        Dim keywords As String = vbNullString
        If metaTagQuery.Count > 0 Then
            Dim keywordsArr() As String = metaTagQuery.Split(",")
            For Each keyword In keywordsArr
                keyword = keyword.Trim
                Dim spaceCount As Integer = keyword.Count(Function(x) x = " ")
                If spaceCount <= 1 Then
                    If keywords = vbNullString Then
                        keywords = keyword
                    Else
                        keywords = keywords + ";" + keyword
                    End If
                End If
            Next
            If Not IsNothing(keywords) Then keywords = keywords.Trim
        End If
        busItem.Keywords = keywords
        Return busItem
    End Function

    Shared Function ScrapeEuroToUsdConversionRate() As Double
        Const url As String = "http://rate-exchange.appspot.com/currency?from=EUR&to=USD"
        Dim pageSource As String = DownloadPageSource(url)
        If pageSource = vbNullString Then
            MsgBox("Error, scraping exchange rate.")
            Return 0
        End If
        Dim regex As New Regex("%rate%:\s(\d+\.\d+),".Replace("%", Chr(34)))
        Dim match As Match = regex.Match(pageSource)
        If match.Success Then
            Dim exchangeRate As String = match.Groups(1).Value.Replace("&nbsp;", "")
            If IsNumeric(exchangeRate) Then
                Return CDbl(exchangeRate)
            Else
                MsgBox("Not valid exchange rate")
                Return 0
            End If
        Else
            MsgBox("Cannot extract exchange rate.")
            Return 0
        End If
    End Function
#End Region
End Class
