<?php

$cookies = '';
$setCookies = false;

ScrapeProductPage('http://www.nutritionxcellence.com/index.php/amino-acids/alri-chaind-out-blue-raspberry-60-srv-1253-detail#.UoYPlVDPHGA');

function ScrapeProductPage($URL){
    $prodPageSource = GetPageSource($URL);
    $dom = GetDomObject($prodPageSource);
    $xPath = GetXPathObject($dom);

    $Category = '';
    $Manufacturer = '';
    $ProductTitle = '';
    $ProductDescription = '';
    $SKUNumber = 'N/A';
    $SalesPrice = ''; #(USD)
    $MSRPPrice = '';
    $ProductImages = '';
    $NutritionFacts = '';
    $OtherIngredients = '';
    $AllerginInfo = '';
    $Warnings = '';
    
    $title = $dom->getElementsByTagName('title')->item(0)->nodeValue;
    $titleArr = explode( ':', $title);
    if (sizeof($titleArr) == 3){
        $Category = trim($titleArr[0]);
        $Manufacturer  = trim($titleArr[1]);
        $ProductTitle = trim($titleArr[2]);
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
    $headerArray = array("Category","Manufacturer","ProductTitle","ProductDescription","SKUNumber",
                         "SalesPrice","MSRPPrice","ProductImages","NutritionFacts","OtherIngredients","AllerginInfo",
                         "Warnings");

    $rowArray = array("$Category","$Manufacturer","$ProductTitle","$ProductDescription","$SKUNumber",
                      "$SalesPrice","$MSRPPrice","$ProductImages","$NutritionFacts","$OtherIngredients","$AllerginInfo",
                      "$Warnings");

    PushDataToCSV('data.csv', $headerArray, $rowArray);   
     
    echo $title.' -- Done'.'<br />';
}

function PushDataToCSV($fileName, $headerArray, $rowArray){
    $fileExists = false;
    if (file_exists($fileName)) $fileExists = true;
        
    $fp = fopen($fileName, 'a');
    
    if (!$fileExists) fputcsv($fp, $headerArray); 
       
    fputcsv($fp, $rowArray);

    fclose($fp);    
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
    //curl_setopt($ch, CURLOPT_USERAGENT, "MozillaXYZ/1.0");
    
    if (!empty($GLOBALS['cookies'])){
        curl_setopt($ch, CURLOPT_COOKIE,  $GLOBALS['cookies']);
    } 

    // Include header in result? (0 = yes, 1 = no)
    curl_setopt($ch, CURLOPT_HEADER, 1);
    
    // Should cURL return or print out the data? (true = return, false = print)
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
      
    // Timeout in seconds
    curl_setopt($ch, CURLOPT_TIMEOUT, 30);
    
    // Download the given URL, and return output
    $output = curl_exec($ch);
    
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
?>