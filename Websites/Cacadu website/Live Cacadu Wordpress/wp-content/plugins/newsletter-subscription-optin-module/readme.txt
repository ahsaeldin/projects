=== Newsletter subscription optin module ===
Contributors: nonletter
Donate link: http://www.sendblaster.com/
Tags: sidebar, widget, newsletter, email, bulk email, management
Requires at least: 2.0.0
Tested up to: 2.9.1
Stable tag: 1.1.6

Plugin for managing subscriptions to a mailing list. It provides a simple form for subscription to your mailing list through single or double opt-in.

== Description ==

This plugin enables you to create simple forms for subscription to your newsletter. It includes a sidebar widget with customizable fields (up to 15) to gather from your users all the information you need for your email marketing campaigns. Users can also use the form to unsubscribe.

From the options page you can:

* choose whether you require a double opt-in (users must follow a link in an email message, in order to complete subscription)
* specify name and number of text fields in the form
* customize text messages and labels
* manage subscribed email addresses 

= Related features =

(Un)subscription requests can be directly processed by SendBlaster Pro [bulk email software](http://www.sendblaster.com "bulk email software") :

* [bulk email software download](http://www.sendblaster.com/free-bulk-emailer-download/ "bulk email software download")

This widget sidebar form for newsletter subscription is also compatible with SendBlaster Free [Newsletter software](http://www.sendblaster.com/newsletter-software-no-recurring-fees/ "newsletter software").

= Plugin Features =

* Subscribe new members
* Unsubscribe existing members
* Stores subscribed email addresses (as a useful backup against mail delivery failures)
* Purges old email addresses
* Customizable sidebar appearance
* Customizable texts and labels
* Adds to your form up to 15 custom fields

= Plugin Options =

* E-mail address for managing subscriptions
* Message to subscriber - subject
* Message to subscriber - content
* Double Opt-in
* Link Love (enable and disable)
* Front side messages
* Front side appearance and custom fields
* Temporary db of newly subscribed members
* Automatic temporary db cleanup


== Installation ==

This section describes how to install the plugin and get it working.

1. Unzip the downloaded archive
2. Upload the script `wpsb-opt-in.php` to the `/wp-content/plugins/` directory
3. Activate the plugin through the **Plugins** menu in WordPress Admin
4. You will find a new item under 'Options' (or 'Settings') menu called 'Wp SendBlaster Opt-in'.
5. Go to `Wp SendBlaster Opt-in` and set the **email address** that you intend to use for managing subscriptions. This should be the same address you have provided in your [bulk email software Manage subscriptions](http://www.sendblaster.com/bulk-email-software/wp-content/manage-subscriptions.gif "manage subscriptions help screenshot") configuration
6. In the `Appearance->Widgets` menu, drag to your sidebar the newly added plugin

== Screenshots ==

1. This is the newsletter subscription form inside your WordPress sidebar
2. This is the admin panel for managing options and texts

== Frequently Asked Questions ==

= I am unable to delete the temporary opted in users. I get this message: Cannot load wpsb-opt-in.php =

The wpsb-opt-in.php file must be in the `/wp-content/plugins/` folder and not in a subdirectory.

= Is it possible to show the form in the body of a page rather than on the sidebar? =

Yes: to have the newsletter widget inside a post you must enable PHP execution inside posts. First install this PHPexec plugin:

http://wordpress.org/extend/plugins/exec-php/

then write this code in your post:

`<?php wpsb_opt_in(); ?>`

