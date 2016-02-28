<?php
    require_once('simpletest/browser.php');
    
    $browser = &new SimpleBrowser();
    
    $browser->useCookies();
    
    $source = GetPageSource('https://www.nutritionxcellence.com/protein');
    
    $dom = GetDomObject($source);
    $xpath = GetXPathObject($dom);

    $subcCategoriesLinksNodes = $xpath->query('//a');

    foreach($subcCategoriesLinksNodes as $node1){
       $ttt = $node1->attributes->item(0)->nodeValue;
       echo $ttt.'<br />';
    }
    echo 'asdf';
    
    function GetPageSource($URL)
    {
        $GLOBALS['browser']->get($URL);
        $output = $GLOBALS['browser']->getContent();
        return $output;
    }

    function GetDomObject($pageSource){
        $dom = new DOMDocument('1.0');
        libxml_use_internal_errors(true);
        if (!is_null($pageSource)){
            $dom->loadHTML($pageSource);
            return $dom;       
        }
        else{
            return;
        }
    }

    function GetXPathObject($dom){
        $xpath = new DOMXpath($dom);
        return $xpath;
    }
?>