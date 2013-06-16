favourites_xls
==============

A script to download all your favourited tweets to an excel spreadsheet. You will need to:

1. Get Python (if you don't already have it) - Please download a version as appropriate for your OS - http://www.python.org/getit/
2. Go to dev.twitter.com, login, create an application, and get the consumer key, secret, access token and secret
3. Download this script and put your access token, secret blah blah where it says you should
4. Run it by doing "python favourites_xls.py all" from the command line. It will generate a twitter_favourites.xls in the same folder. 
3. Um...not that I'm an XBox One fan but be online when you run the script. In case you have more than 15 pages of favourites (each page has 10) in Twitter, the script will take a while because I've put in a sleep(900) to ensure that Twitter doesn't throw the Rate-limit book at you

**Usage**
	
1. python favourites_xls.py all (for all faves)
2. python favourites_xls.py 2 (for 2 pages of recently faved tweets)
