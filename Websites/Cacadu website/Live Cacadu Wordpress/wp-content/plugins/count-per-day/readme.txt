=== Count per Day ===
Contributors: Tom Braider
Tags: counter, count, posts, visits, reads, dashboard, widget, shortcode
Requires at least: 3.0
Tested up to: 3.4.1
Stable tag: 3.2.4
License: Postcardware :)
Donate link: http://www.tomsdimension.de/postcards

Visit Counter, shows reads and visitors per page, visitors today, yesterday, last week, last months and other statistics.

== Description ==

* count reads and visitors
* shows reads per page
* shows visitors today, yesterday, last week, last months and other statistics on dashboard
* shows country of your visitors
* you can show these statistics on frontend per widget or shortcodes too
* Plugin: http://www.tomsdimension.de/wp-plugins/count-per-day
* Donate: http://www.tomsdimension.de/postcards

"Count per Day" counts 1 visit per IP per day. So any reload of the page do not increment the counter.

= Languages, Translators =

- up to date translations:
- Bulgarian		- joro -							http://www.joro711.com
- German		- Tom -								http://www.tomsdimension.de
- Japanese		- Juno Hayami						http://juno.main.jp/blog/
- Portuguese	- Beto Ribeiro -					http://www.sevenarts.com.br
- Russian		- Ilya Pshenichny -					http://iluhis.com

- older, incomplete translations:
- Azerbaijani	- Bohdan Zograf -					http://wwww.webhostingrating.com
- Belarusian	- Alexander Alexandrov -			http://www.designcontest.com
- Dansk			- Jonas Thomsen -					http://jonasthomsen.com
- Dutch NL		- Rene -							http://wpwebshop.com
- Espanol		- Juan Carlos del R&iacute;o -
- Finnish		- Jani Alha -						http://www.wysiwyg.fi
- France		- Bjork -							http://www.habbzone.fr
- Greek			- Essetai_Imar -					http://www.elliniki-grothia.com
- Hindi			- Love Chandel -					http://outshinesolutions.com
- Italian		- Gianni Diurno -					http://gidibao.net
- Lithuanian	- Nata Strazda - 					http://www.webhostinghub.com
- Norwegian		- Stein Ivar Johnsen -				http://iDyrøy.no
- Polish 		- LeXuS -							http://intrakardial.de
- Romanian		- Alexander Ovsov -					http://webhostinggeeks.com
- Swedish		- Magnus Suther -					http://www.magnussuther.se
- Turkish		- Emrullah Tahir Ekmek&ccedil;i - 	http://emrullahekmekci.com.tr
- Ukrainian		- Iflexion design -					http://iflexion.com


== Installation ==

1. unzip plugin directory into the '/wp-content/plugins/' directory
1. activate the plugin through the 'Plugins' menu in WordPress
1. on every update you have to deactivate and reactivate the plugin to update some settings!

The activation will create or update a table wp_cpd_counter.

The Visitors-per-Day function use 7 days as default. So don't surprise about a wrong value in the first week.

**Configuration**

See the Options Page and check the default values.

== Frequently Asked Questions ==

= Need Help? Find Bug? =

read and write comments on http://www.tomsdimension.de/wp-plugins/count-per-day

== Screenshots ==

1. Statistics on Count-per-Day Dashboard
2. Options
3. Widget sample

== Arbitrary section ==

**Shortcodes**

You can use these shortcodes in the content of your posts to show a number or list
or in your theme files while adding e.g. '&lt;?php echo do_shortcode("[THE_SHORTCODE]"); ?>'.
To use the shortcodes within a text widget you have to add 'add_filter("widget_text", "do_shortcode");' to the 'functions.php' of your theme.

[CPD_READS_THIS]
[CPD_READS_TOTAL]
[CPD_READS_TODAY]
[CPD_READS_YESTERDAY]
[CPD_READS_LAST_WEEK]
[CPD_READS_THIS_MONTH]
[CPD_READS_PER_MONTH]
[CPD_VISITORS_TOTAL]
[CPD_VISITORS_ONLINE]
[CPD_VISITORS_TODAY]
[CPD_VISITORS_YESTERDAY]
[CPD_VISITORS_LAST_WEEK]
[CPD_VISITORS_THIS_MONTH]
[CPD_VISITORS_PER_MONTH]
[CPD_VISITORS_PER_DAY]
[CPD_VISITORS_PER_POST]
[CPD_FIRST_COUNT]
[CPD_MOST_VISITED_POSTS]
[CPD_POSTS_ON_DAY]
[CPD_CLIENTS]
[CPD_COUNTRIES]
[CPD_COUNTRIES_USERS]
[CPD_REFERERS]
[CPD_POSTS_ON_DAY date="2010-10-06" limit="3"]
- date (optional), format: year-month-day, default = today
- limit (optional): max records to show, default = all 
[CPD_MAP width="500" height="340" what="reads" min=1]
- width and height: size, default 500x340 px
- what: map content - reads|visitors|online, default reads
- min: 1 (disable title, legend and zoombar), default 0
[CPD_SEARCHES days="14" limit="20"]
- days (optional), show last x days
- limit (optional): show x most searched strings 

**Functions**

You can place these functions in your template.
Use
<code>&lt;?php
global $count_per_day;
if(method_exists($count_per_day,"show")) echo $count_per_day->getReadsAll(true);
?></code>
to check if plugin is activated.

'show( $before, $after, $show, $count, $page )'

* $before = text before number e.g. '&lt;p&gt;' (default "")
* $after = text after number e.g. 'reads&lt;/p&gt;' (default " reads")
* $show = true/false, "echo" complete string or "return" number only (default true)
* $count = true/false, false will not count the reads (default true)
* $page (optional) PostID

'count()'

* only count reads, without any output
* 'show' call it

'getFirstCount( $return )'

* shows date of first count
* $return: 0 echo, 1 return output

'getUserPerDay( $days, $return )'

* shows average number of visitors per day of the last _$days_ days
* default on dashboard (see it with mouse over number) = "Latest Counts - Days" in options
* $return: 0 echo, 1 return output

'getReadsAll( $return )'
 
* shows number of total reads
* $return: 0 echo, 1 return output

'getReadsToday( $return )'

* shows number of reads today
* $return: 0 echo, 1 return output

'getReadsYesterday( $return )'
 
* shows number of reads yesterday
* $return: 0 echo, 1 return output

'getReadsLastWeek( $return )'

* shows number of reads last week (7 days)
* $return: 0 echo, 1 return output

'getReadsThisMonth( $return )'

* shows number of reads current month
* $return: 0 echo, 1 return output

'getReadsPerMonth( $return )'

* lists number of reads per month
* $return: 0 echo, 1 return output

'getUserAll( $return )'

* shows number of total visitors
* $return: 0 echo, 1 return output

'getUserOnline( $frontend, $country, $return )'

* shows number of visitors just online
* $frontend: 1 no link to map
* $country: 0 number, 1 country list
* $return: 0 echo, 1 return output

'getUserToday( $return )'

* shows number of visitors today
* $return: 0 echo, 1 return output

'getUserYesterday( $return )'
 
* shows number of visitors yesterday
* $return: 0 echo, 1 return output

'getUserLastWeek( $return )'

* shows number of visitors last week (7 days)
* $return: 0 echo, 1 return output

'getUserThisMonth( $return )'

* shows number of visitors current month
* $return: 0 echo, 1 return output

'getUserPerMonth( $frontend, $return )'

* lists number of visitors per month
* $frontend: 1 no links
* $return: 0 echo, 1 return output

'getUserPerPost( $limit, $frontend, $return )'

* lists _$limit_ number of posts, -1: all, 0: get option from DB, x: number
* $frontend: 1 no links
* $return: 0 echo, 1 return output

'getMostVisitedPosts( $days, $limits, $frontend, $postsonly, $return )'

* shows a list with the most visited posts in the last days
* $days = days to calc (last days), 0: get option from DB
* $limit = count of posts (last posts), 0: get option from DB
* $frontend: 1 no links
* $postsonly: 0 show, 1 don't show categories and taxonomies
* $return: 0 echo, 1 return output

'getVisitedPostsOnDay( $date, $limit, $show_form, $show_notes, $frontend, $return )'

* shows visited pages at given day
* $date day in MySQL date format yyyy-mm-dd, 0 today
* $limit count of posts
* $show_form show form for date selection, default on, in frontend set it to 0
* $show_notes show button to add notes in form, default on, in frontend set it to 0
* $frontend: 1 no links
* $return: 0 echo, 1 return output

'getClients( $return )'

* shows visits per client/browser in percent
* $return: 0 echo, 1 return output

'getReferers( $limit, $return, $days )'

* lists top _$limit_ referrers of the last $days days, 0: get option from DB, x: number
* $return: 0 echo, 1 return output

'getMostVisitedPostIDs( $days, $limit, $cats, $return_array )'

* $days last x days, default = 365
* $limit return max. x posts, default = 10
* $cats IDs of categories to filter, array or number
* $return_array true returns an array with Post-ID, title and count, false returns comma separated list of Post-IDs

'function getMap( $what, $width, $height, $min )'

* gets a world map
* $what visitors|reads|online
* $width size in px
* $height size in px
* $min : 1 disable title, legend and zoombar

'getDayWithMostReads( $return )'

* shows day with most Reads
* $return: 0 echo, 1 return output

'getDayWithMostVisitors( $return )'

* shows day with most Visitors
* $return: 0 echo, 1 return output

**GeoIP**

* With GeoIP you can associate your visitors to an country using the ip address.
* In the database a new column 'country' will be insert on plugin activation.
* On options page you can update you current visits. This take a while! The Script checks 100 IP addresses at once an reload itself until less then 100 addresses left. Click the update button to check the rest.
* If the rest remains greater than 0 the IP address is not in GeoIP database (accuracy 99.5%).
* You can update the GeoIP database from time to time to get new IP data. This necessitates write rights to geoip directory (e.g. chmod 777).
* If the automatically update don't work download <a href="http://geolite.maxmind.com/download/geoip/database/GeoLiteCountry/GeoIP.dat.gz">GeoIP.dat.gz</a> and extract it to the "geoip" directory.
* More information about GeoIP on http://www.maxmind.com/app/geoip_country

== Changelog ==

= 3.2.4 =
+ Bigfix: security fix, check user permissions

= 3.2.3 =
+ Bugfix: security fix, XSS in search words, thanks to http://www.n0lab.com/?p=163

= 3.2.2 =
+ New: counter column in custom post lists
+ Bugfix: errors in search words
+ Bugfix: wrong counts in posts lists

= 3.2.1 =
+ Bugfix: massbot delete error
+ Bugfix: search words array sometimes corrupt
+ Bugfix: add collected data to reads per post, thanks to Suzakura Karin http://yumeneko.pmfan.jp / http://is.gd/VWNyLq 
+ Language update: Japanese, thanks to Juno Hayami
+ Language update: Portuguese, thanks to Beto Ribeiro
+ Language update: Russian, thanks to Ilya Pshenichny
+ Language update: Bulgarian, thanks to joro

= 3.2 =
+ New: save search words
+ New shortcode: CPD_COUNTRIES_USERS
+ New: flags for Bahamas, Mongolia, Cameroon and Kazakhstan
+ Bugfix: can't move widgets
+ Bugfix: visitors per post list
+ Bugfix: "Clean Database" deleted collection too
+ Bugfix: browser summary Chrome/Safari fixed
+ Bugfix: get real remote IP address, not local server
+ Bugfix: security fixes
+ Change: create collection functions optimized
+ New language: Romanian, thanks to Alexander Ovsov
+ New language: Hindi, thanks to Love Chandel
+ New language: Finnish, thanks to Jani Alha
+ Language update: Ukrainain, thanks to Iflexion design

= 3.1.1 =
+ Bugfix: important fixes in map.php and download.php, thanks to http://6scan.com

= 3.1 =
+ New: memory check before backup to avoid "out of memory" error
+ New: create temporary backup files for download only
+ New: delete backup files in wp-content on settings page
+ Bugfix: all posts shows 1 read in posts list
+ Bugfix: clean database shows 0 entries deleted

= 3.0 =
+ New: use now default WordPress database functions to be compatible to e.g. multi-db plugins
+ New: backup your counter data
+ New: collect entries of counter table per month and per post to reduce the database and speed up the statistics
+ New: functions and shortcodes [CPD_DAY_MOST_READS] [CPD_DAY_MOST_USERS] to shows days with most reads/visitors
+ New: option to cut referrer on "?" to not store query strings
+ New: parameter '$postsonly' for 'getMostVisitedPosts' function to list single posts and pages only
+ New: flags for Moldavia and Nepal
+ New language: Norwegian, thanks to Stein Ivar Johnsen and Tore Johnny Bråtveit
+ New language: Azerbaijani, thanks to Bohdan Zograf
+ New language: Japanese, thanks to Juno Hayami
+ Bugfix: visitors per month list
+ Change: some function parameters
+ Change: little memory optimizing
+ Change: visitors currently online and notes will now managed per option, without seperate tables in database
+ Change: design updated
+ Change: old bar charts deleted
