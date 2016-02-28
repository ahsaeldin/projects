=== Plugin Name ===
Contributors: Hetjens
Tags: feed, redirect, feedburner, location, rss, atom
Requires at least: 2.7.0
Tested up to: 2.9.2
Stable tag: 0.4

This plug-in redirect reqests to the main or comment feed of the blog to Feedburner or a similar service.

== Description ==

This plug-in redirects requests to the main or comment feeds, and only theses two, to Feedburner or a similar service. In comparision to the original Feedsmith-plug-in is has three advantages:

1. It redirects only the two main feeds of the blog and not all. The category, tag and author feeds are still accessible

2. If you need access to the original blog feed you can add the URL parameter "noredirect" (i.e. /feed/?noredirect) to your feed address and it will not forward your request.

3. The Plugin is more integrated into the administration interface. It adds its settings to the default "reading" settings pages and do not create an own page.

== Installation ==

1. Upload the directory hetjens-feed_redirect to '/wp-content/plugins/' directory
2. Activate the plugin through the 'Plugins' menu in WordPress
3. Set up your Feed-URLs in Settings -> Reading
3. You can use it

== Changelog ==

= 0.4 =
* First Plugin directory version
