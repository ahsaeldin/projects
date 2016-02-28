	function get(key_str)//Get the query String 
	{
		var query = window.location.search.substr(1);
		var pairs = query.split("&");
		for(var i = 0; i < pairs.length; i++) 
		{
			var pair = pairs[i].split("=");
			if(unescape(pair[0]) == key_str)
				return unescape(pair[1]);//return the query string
		}
		return 'home.html';//If no query string then load home.html
	
	}