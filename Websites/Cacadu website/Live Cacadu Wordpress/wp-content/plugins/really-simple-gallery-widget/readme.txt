=== Really Simple Gallery Widget ===
Contributors: helenyhou
Donate link: http://www.helenhousandi.com/wordpress/donate/
Tags: gallery, widget
Requires at least: 2.8
Tested up to: 3.2
Stable tag: 1.1

Simple widget for displaying images from the Media Library with plenty of configurable options.

== Description ==

Really Simple Gallery Widget adds a widget to display images from the Media Library, no extra uploading or creating custom post types required. You can choose to show images from a specific post/page or the entire media library, and/or show only images attached to the post/page currently being viewed. Especially helpful if you create photo gallery pages using the built-in WordPress media library and just want to be able to display those images in a widget area.

= Features =
* Add as many widgets as you want, wherever you want
* Display images from a specific post/page
* Display images from the entire Media Library
* Display images from the post/page currently being viewed
* Select a number of images
* Select any registered size in WordPress
* Display the images in ascending, descending, or random order
* Show or hide captions
* Link the images to the original file, post, anchor in the post, attachment page, or nothing. NOTE: If the image is not attached to a post, the file link will be used instead.
* Add a prefix to the link and image title (appears as a tooltip)
* Use a rel attribute for the link - great for lightboxes

== Installation ==

Really Simple Gallery Widget is most easily installed automatically via the Plugins tab in your blog administration panel.

= Manual Installation =

1. Upload the `really-simple-gallery-widget` folder to the `/wp-content/plugins/` directory
1. Activate the plugin through the 'Plugins' menu in WordPress
1. Head over to Appearance &rarr; Widgets to set up one or more Really Simple Gallery Widgets

== Frequently Asked Questions ==

= Why won't images I've hotlinked or copy-pasted into my content appear? =
This widget can only grab images that were uploaded into your Media Library. See the [WordPress Codex](http://codex.wordpress.org/Using_Image_and_File_Attachments#Attachment_to_a_Post "WordPress Codex") for more information.

= Why is the wrong kind of link being used? =
The image is probably not associated with a post or page and will default to linking to the file if a link option is chosen. To associate (attach) an image with a post or page, go to your Media Library and click on the Unattached link next to the item to choose a post to attach it to. See the [WordPress Codex](http://codex.wordpress.org/Using_Image_and_File_Attachments#Attachment_to_a_Post "WordPress Codex") for more information. An option to choose what to link to in the case of not having an associated post or page will probably be added in the future.

= How do I get the ID for the post or page? =
The easiest way is to mouseover or click an edit link for the post or page in question. The ID number will appear in the URL; e.g. `http://yoursite.com/wp-admin/post.php?post=41&action=edit` indicates that the ID of the post or page you want to reference is 41.

= Why is the anchor link not working? =
The anchor link relies on the ID that WordPress automatically generates when you insert an image with a caption. If you inserted the image manually or without a caption, the anchor won't jump you to the spot in the page. The ability to specify an anchor may be added at a later time, or you can just add the ID (attachment_##) to the img tag.

= I selected a registered size but the images are showing up huge or in the wrong size. =
The images may be missing the thumbnails of that size and by default will pull the full size image instead. Try using Viper007Bond's fantastic [Regenerate Thumbnails](http://wordpress.org/extend/plugins/regenerate-thumbnails/ "Regenerate Thumbnails") plugin to create new versions for any new or changed image sizes.

= How can I make the widget look prettier? =
See [Styling the Really Simple Gallery Widget](http://www.helenhousandi.com/wordpress/2011/02/styling-the-really-simple-gallery-widget/ "Styling the Really Simple Gallery Widget") for more information and some examples of what you might do.

== Screenshots ==

1. Widget options
2. Sample display with prefixed title showing

== Changelog ==

= 1.1 =
* New options for entire media library and single post/page
* HTML output cleanup
* Better security
* Smaller memory footprint (hopefully)

= 1.0 =
* First version

== Upgrade Notice ==

= 1.1 =
More display options