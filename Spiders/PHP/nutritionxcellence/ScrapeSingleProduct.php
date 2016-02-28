<?php

$cookies = '';
$setCookies = false;


ScrapeProductPage('https://www.nutritionxcellence.com/accessories/flexsports-international-power-wrist-wrap-gloves-black-small-1-1033-detail#.Uo4291DPHGA', 'Amino Acid');

die('d');

# Get a list of Cats URLs.
$startURL = 'https';
$catsPageSource = GetPageSource($startURL);
if (empty($catsPageSource)){
    die('Sorry cURL failed to reterive page source');
}

$filePath = dirname(__FILE__).'/data.csv';
deleteFileIfExists($filePath);

$dom = GetDomObject($catsPageSource);
$xpath = GetXPathObject($dom);

echo ('Getting categories links').'<br />';
$categoriesLinks = GetCategoriesLinks($xpath);
echo ($categoriesLinks->length.' Found.').'<br />';

# Loop through them one by one
$products = array();
$cats = array();
foreach($categoriesLinks as $catName=>$catURL){
    
    $arr = array();
    
    array_push($arr, $catURL);
    
    $catURL = preg_replace("/http:/", "https:", $catURL);
    $listingPageSource = GetProductsLinks($catURL, $products);
    
    while(true){
        $nextURL = GetNextURL($listingPageSource);
        $nextURL = preg_replace("/http:/", "https:", $nextURL);
        if(!empty($nextURL) && !in_array($nextURL, $arr)){
            array_push($arr, $nextURL);
            $listingPageSource = GetProductsLinks($nextURL, $products);
            $nextURL = '';
        }   
        else{
            break;
        }
    }
    
    $cats[$catName] = $products;
    echo 'Prodcuts colected from '.$catName.'<br />';
}

echo sizeof($products).' found in all catgories'.'<br />';

foreach($cats as $catName=>$prodsURLs){
    echo $catName.':<br />';
    foreach($prodsURLs as $productURL){
        $productURL = preg_replace("/http:/", "https:", $productURL);
        ScrapeProductPage($productURL);
    }
}

function GetPageSource($URL)
{
    // is cURL installed yet?
    if (!function_exists('curl_init')) {
        die('Sorry cURL is not installed!');
    }
    
    // OK cool - then let's create a new cURL resource handle
    $ch = curl_init();
    
    // Now set some options (most are optional)
    
    // Set URL to download
    curl_setopt($ch, CURLOPT_URL, $URL);
                
    // Set a referer
    //curl_setopt($ch, CURLOPT_REFERER, "http://www.example.org/yay.htm");
    
    // User agent
    curl_setopt($ch, CURLOPT_USERAGENT, "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.69 Safari/537.36");
    
    if (!empty($GLOBALS['cookies'])){
        curl_setopt($ch, CURLOPT_COOKIE,  $GLOBALS['cookies']);
    } 

    // Include header in result? (0 = yes, 1 = no)
    curl_setopt($ch, CURLOPT_HEADER, 1);
    
    // Should cURL return or print out the data? (true = return, false = print)
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
      
    // Timeout in seconds
    curl_setopt($ch, CURLOPT_TIMEOUT, 30);
    
    # Read more about 
    #http://curl.haxx.se/docs/sslcerts.html
    #http://stackoverflow.com/questions/15201755/i-had-a-errors-with-curl-and-ssl-certificate
    #http://stackoverflow.com/questions/6400300/php-curl-https-causing-exception-ssl-certificate-problem-verify-that-the-ca-cer
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
    
    // Download the given URL, and return output
    $output = curl_exec($ch);
    
    if(curl_errno($ch)){
        echo 'Curl error: ' . curl_error($ch).'<br />';
    }

    preg_match_all('|Set-Cookie: (.*);|U', $output, $results);   
   
    $GLOBALS['cookies'] = implode(';', $results[1]);  
    
    // Close the cURL resource, and free system resources
    curl_close($ch);
    
    return $output;
}

function GetDomObject($pageSource){
    $dom = new DOMDocument('1.0');
    libxml_use_internal_errors(true);
    $dom->loadHTML($pageSource);
    return $dom;
}

function GetXPathObject($dom){
    $xpath = new DOMXpath($dom);
    return $xpath;
}

function GetCategoriesLinks($xpath){
    $linksArr = array();
    $categoriesLinksNodes = $xpath->query("//*[contains(@class, 'VMmenu shop_by_category')]/li/div/a");
    foreach ($categoriesLinksNodes as $link){
        $catLink = 'http://www.nutritionxcellence.com'.$link->getAttribute('href').'/results1-80';
        $catName = trim($link->nodeValue);
        $linksArr[$catName] = $catLink;
    }
    return $linksArr;
}

function GetProductsLinks($catURL, &$products){
    $listingPageSource = GetPageSource($catURL);
    if (empty($listingPageSource)){
        die('Sorry cURL failed to reterive page source');
    }
    
    $dom = GetDomObject($listingPageSource);
    $xpath = GetXPathObject($dom);
    $imgs = $dom->getElementsByTagName('img');
    
    foreach($imgs as $img){
        $imgParentAnchor = $img->parentNode->getAttribute('href');
        $imgParentAnchor = 'http://www.nutritionxcellence.com'.$imgParentAnchor;
        if (strpos($imgParentAnchor,'/index.php/') != false) {
            array_push($products, $imgParentAnchor);
        }
    }
    
    return $listingPageSource;
}

function GetNextURL($listingPageSource){
    if (preg_match('/<a\stitle="Next"\shref="(.+?)"\sclass="pagenav">Next</', $listingPageSource, $matches)) {
        $nextURL = 'http://www.nutritionxcellence.com'.$matches[1];
        return $nextURL;
    }
}

function PrepareField($field){
    $field = trim(preg_replace('/"/', "'", $field));
    $field = trim(preg_replace('/,/', ";", $field));
    return $field;
}

function ScrapeProductPage($URL, $catName){
    $prodPageSource = GetPageSource($URL);
    $dom = GetDomObject($prodPageSource);
    $xPath = GetXPathObject($dom);

    $Category = '';
    $Manufacturer = '';
    $ProductTitle = '';
    $ProductDescription = '';
    $SalesPrice = ''; #(USD)
    $MSRPPrice = '';
    $InStock = 'In stock';
    $ProductImages = '';
    $NutritionFacts = '';
    $OtherIngredients = '';
    $AllerginInfo = '';
    $Warnings = '';
    
    if (strpos($prodPageSource,'Notify Me') != false) {
        $InStock = "Out of stock";
    }
       
    $Category = $catName;
    $manufacturerNode = $xPath->query("//*[contains(@class, 'manufacturer')]");
    if (!is_null($manufacturerNode)){
        $Manufacturer  = $manufacturerNode->item(0)->nodeValue;
        $Manufacturer = trim(str_replace('Manufacturer:','',$Manufacturer));
    }
    
    $titleNode = $xPath->query('//*[@id="main"]/div[2]/h1');
    if (!is_null($titleNode)){
        $title = $titleNode->item(0)->nodeValue;
        $ProductTitle  = trim(str_replace("$Manufacturer:",'',$title));
    }        

    $descNode = $xPath->query("//*[contains(@class, 'product-description')]");
    if (!is_null($descNode)){
        $ProductDescription = $descNode->item(0)->nodeValue;
        $ProductDescription = trim(preg_replace('/Description/', '', $ProductDescription));
        $ProductDescription = PrepareField($ProductDescription);
    }
    
    $SalesPriceNode = $xPath->query("//*[contains(@class, 'PricesalesPrice')]");
    if (!is_null($SalesPriceNode)){
        $SalesPrice = $SalesPriceNode->item(0)->nodeValue;
        if( preg_match("/\d+\.\d+/", $SalesPrice, $matches)){
            $SalesPrice = $matches[0];
        }
    }
    
    preg_match("/MSRP:.+?<s>(.+?)<\/s>/", $prodPageSource, $matchs);
    if (!is_null($matchs)){
        $MSRPPrice = $matchs[1];
    }
    
    $imageNode = $dom->getElementById('medium-image');
    if (!is_null($imageNode)){
        $ProductImages = 'http://www.nutritionxcellence.com'.$imageNode->getAttribute('src');
    }   
  
    for ($i=1; $i<=4; $i++){
        $tabNode = $dom->getElementById("tabs-$i");
        if (!is_null($tabNode)){
            switch ($i){
                case 1:
                    $NutritionFacts = PrepareField($tabNode->nodeValue);
                    break;
                case 2:
                    $OtherIngredients = PrepareField($tabNode->nodeValue);
                    break;
                case 3:
                    $AllerginInfo = PrepareField($tabNode->nodeValue);
                    break;
                case 4:
                    $Warnings = PrepareField($tabNode->nodeValue);
            }       
        }          
    }

    # Push data to a csv file.
    $headerArray = array("Category","Manufacturer","ProductTitle","ProductDescription",
                         "SalesPrice","MSRPPrice","InStock", "ProductImages","NutritionFacts","OtherIngredients","AllerginInfo",
                         "Warnings");

    $rowArray = array("$Category","$Manufacturer","$ProductTitle","$ProductDescription",
                      "$SalesPrice","$MSRPPrice","$InStock","$ProductImages","$NutritionFacts","$OtherIngredients","$AllerginInfo",
                      "$Warnings");

    PushDataToCSV('data.csv', $headerArray, $rowArray);   
     
    echo $title.' -- Done'.'<br />';
}

function deleteFileIfExists($filePath){
    if(is_writable($filePath)){
        unlink($filePath); 
    }
}

function PushDataToCSV($fileName, $headerArray, $rowArray){
    $fileExists = false;
    if (file_exists($fileName)) $fileExists = true;
        
    $fp = fopen($fileName, 'a');
    
    if (!$fileExists) fputcsv($fp, $headerArray); 
       
    fputcsv($fp, $rowArray);

    fclose($fp);    
}

function innerHTML( $contentdiv ) {
    $r = '';
    $elements = $contentdiv->childNodes;
    foreach( $elements as $element ) { 
            if ( $element->nodeType == XML_TEXT_NODE ) {
                    $text = $element->nodeValue;
                    // IIRC the next line was for working around a
                    // WordPress bug
                    //$text = str_replace( '<', '&lt;', $text );
                    $r .= $text;
            }	 
            // FIXME we should return comments as well
            elseif ( $element->nodeType == XML_COMMENT_NODE ) {
                    $r .= '';
            }	 
            else {
                    $r .= '<';
                    $r .= $element->nodeName;
                    if ( $element->hasAttributes() ) { 
                            $attributes = $element->attributes;
                            foreach ( $attributes as $attribute )
                                    $r .= " {$attribute->nodeName}='{$attribute->nodeValue}'" ;
                    }	 
                    $r .= '>';
                    $r .= innerHTML( $element );
                    $r .= "</{$element->nodeName}>";
            }	 
    }	 
    return $r;
}
?>