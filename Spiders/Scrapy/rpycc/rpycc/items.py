# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field

class RpyccItem(Item):
    # define the fields for your item here like:
    profile_url = Field()
    name = Field()
    location = Field()
    email = Field()
    member_since = Field()
    address = Field()
    address_phone1 = Field()
    address_phone2 = Field()
    address_phone3 = Field()
    profile_image = Field()

# <h1 class="memberprofile-user-name">MR GREGORY TALBOTT</h1>
#
# ----------------------
#
# <span class="memberprofile-user-location">BOCA RATON, FL</span>
#
# ----------------------
#
# <span class="memberprofile-user-email"><a href="mailto:gtalbott@talbottrealty.com">gtalbott@talbottrealty.com</a></span>
#
# ----------------------
#
# <span class="memberprofile-user-content-misc"><em>Member Since:</em> 2/27/2003</span>
#
# ----------------------
#
# <address>
# 		250 NE 5TH AVENUE<br>
#
#
# 		<span class="memberprofile-primary-adr-city-state">BOCA RATON, FL</span>  33432<br>
# 		<br>
# 		</address>
#
# ----------------------
#
# 		<span class="memberprofile-user-address-phone"><em>Home:&nbsp;</em>561-289-7157 C</span>
#
# ----------------------
#
# 		<span class="memberprofile-user-address-phone"><em>Other:&nbsp;</em>561-392-8525 O</span>
#
# ----------------------
#
# 		<span class="memberprofile-user-address-phone"><em>Other 2:&nbsp;</em>561-361-3081 W</span>
# ----------------------
#
# <div class="memberprofile-user-pic-mask"><a href="/Images/Library/1792.jpg" title="Click to Enlarge this Photo of: MR DAVID ADAMS"></a></div>
#
# ----------------------