<?php
	
	switch ($_GET["c"])
	{
	case 'time':
  		echo getYear();
  		break;
	default:
  		/*echo "No number between 1 and 3";*/
	}

	function getYear(){
		echo date("Y");
	}
	
?>