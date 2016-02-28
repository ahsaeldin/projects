'Holds data to be scraped.
Class BusinessItem
    Property SKUNumber As String
    Property Categories As String
    Property ProductTitle As String
    Property Price As String
    Property ProductUrl As String
    Property ProductDescriptionText As String
    Property ProductDescriptionHtml As String
    Property ProdcutSmallImageUrl As String
    Property ProdcutLargeImageUrl As String
    Property Images As New List(Of String)
    Property Keywords As String
End Class

'Holds every department data.
<Serializable()>
Friend Class Department
    Property Level As Integer
    Property DepartmentUrl As String
    Property DepartmentName As String
End Class

Friend Class MagentoOutput
    Property sku As String
    Property _store As String
    Property categories As String
    Property _attribute_set As String = "Default"
    Property _type As String = "simple"
    Property _category As String
    Property _root_category As String
    Property _product_websites As String = "base"
    Property description As String
    Property image As String
    Property name As String
    Property price As String
    Property short_description As String
    Property small_image As String
    Property status As String = "1"
    Property tax_class_id As String = "2"
    Property thumbnail As String
    Property visibility As String = "4"
    Property weight As String = "0"
    Property qty As String = "25"
    Property min_qty As String = "1"
    Property use_config_min_qty As String = "1"
    Property is_qty_decimal As String = "0"
    Property backorders As String = "0"
    Property use_config_backorders As String = "1"
    Property min_sale_qty As String = "1"
    Property use_config_min_sale_qty As String = "1"
    Property max_sale_qty As String = "0"
    Property use_config_max_sale_qty As String = "1"
    Property is_in_stock As String = "1"
    Property notify_stock_qty As String
    Property use_config_notify_stock_qty As String = "1"
    Property manage_stock As String = "0"
    Property use_config_manage_stock As String = "1"
    Property stock_status_changed_auto As String = "0"
    Property use_config_qty_increments As String = "1"
    Property qty_increments As String = "0"
    Property use_config_enable_qty_inc As String = "1"
    Property enable_qty_increments As String = "0"
    Property media_gallery As String
    Property is_decimal_divided As String = "0"
    Property _links_related_sku As String
    Property _links_related_position As String
    Property _links_crosssell_sku As String
    Property _links_crosssell_position As String
    Property _links_upsell_sku As String
    Property _links_upsell_position As String
    Property _associated_sku As String
    Property _associated_default_qty As String
    Property _associated_position As String
    Property _tier_price_website As String
    Property _tier_price_customer_group As String
    Property _tier_price_qty As String
    Property _tier_price_price As String
    Property _group_price_website As String
    Property _group_price_customer_group As String
    Property _group_price_price As String
    Property _media_attribute_id As String
    Property _media_image As String
    Property _media_lable As String
    Property _media_position As String
    Property _media_is_disabled As String

    Public Shared Function CreateMagentoObject(data As String()) As MagentoOutput
        Dim magen As New MagentoOutput
        With magen
            .sku = data(0)
            ._store = data(1)
            .categories = data(2)
            ._attribute_set = data(3) '= "Default"
            ._type = data(4) '= "simple"
            ._category = data(5)
            ._root_category = data(6)
            ._product_websites = data(7) '= "base"
            .description = data(8)
            .image = data(9)
            .name = data(10)
            .price = data(11)
            .short_description = data(12)
            .small_image = data(13)
            .status = data(14) '= "1"
            .tax_class_id = data(15) '= "2"
            .thumbnail = data(16)
            .visibility = data(17) '= "4"
            .weight = data(18) '= "0"
            .qty = data(19) '= "1"
            .min_qty = data(20) '= "1"
            .use_config_min_qty = data(21) '= "1"
            .is_qty_decimal = data(22) '= "0"
            .backorders = data(23) '= "0"
            .use_config_backorders = data(24) '= "1"
            .min_sale_qty = data(25) '= "1"
            .use_config_min_sale_qty = data(26) '= "1"
            .max_sale_qty = data(27) '= "0"
            .use_config_max_sale_qty = data(28) '= "1"
            .is_in_stock = data(29) '= "1"
            .notify_stock_qty = data(30)
            .use_config_notify_stock_qty = data(31) '= "1"
            .manage_stock = data(32) '= "0"
            .use_config_manage_stock = data(33) '= "1"
            .stock_status_changed_auto = data(34) '= "0"
            .use_config_qty_increments = data(35) '= "1"
            .qty_increments = data(36) '= "0"
            .use_config_enable_qty_inc = data(37) '= "1"
            .enable_qty_increments = data(38) '= "0"
            .media_gallery = data(39)
            .is_decimal_divided = data(40) '= "0"
            ._links_related_sku = data(41)
            ._links_related_position = data(42)
            ._links_crosssell_sku = data(43)
            ._links_crosssell_position = data(44)
            ._links_upsell_sku = data(45)
            ._links_upsell_position = data(46)
            ._associated_sku = data(47)
            ._associated_default_qty = data(48)
            ._associated_position = data(49)
            ._tier_price_website = data(50)
            ._tier_price_customer_group = data(51)
            ._tier_price_qty = data(52)
            ._tier_price_price = data(53)
            ._group_price_website = data(54)
            ._group_price_customer_group = data(55)
            ._group_price_price = data(56)
            ._media_attribute_id = data(57)
            ._media_image = data(58)
            ._media_lable = data(59)
            ._media_position = data(60)
            ._media_is_disabled = data(61)
        End With
        Return magen
    End Function
End Class

Friend Class iOfferOutput
    Property Category As String
    Property Title As String
    Property Description As String
    Property Shipping As String
    Property Price As String
    Property Quantity As String
    Property Condition As String = "New"
    Property Image1 As String
    Property Image2 As String
    Property Image3 As String
    Property Keyword1 As String
    Property Keyword2 As String
    Property Keyword3 As String
    Property Keyword4 As String

    Shared Function CreateiOfferObject(data As String()) As iOfferOutput
        Dim iOfferOutput As New iOfferOutput
        With iOfferOutput
            .Category = data(0)
            .Title = data(1)
            .Description = data(2)
            .Shipping = data(3)
            .Price = data(4)
            .Quantity = data(5)
            .Condition = data(6) '"New"
            .Image1 = data(7)
            .Image2 = data(8)
            .Image3 = data(9)
            .Keyword1 = data(10)
            .Keyword2 = data(11)
            .Keyword3 = data(12)
            .Keyword4 = data(13)
        End With
        Return iOfferOutput
    End Function
End Class
