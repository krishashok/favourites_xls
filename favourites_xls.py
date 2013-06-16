from twython import Twython
import oauth2 as oauth
from xlwt.Workbook import *
from xlwt import easyxf,Formula
import time
import sys


# Uncomment these lines and fill in your details before running
# access_token_key = <Your access token key>
# access_token_secret = <Your access token secret>

# consumer_key = <Your consumer key>
# consumer_secret = <Your consumer secret>

oauth_token    = oauth.Token(key=access_token_key, secret=access_token_secret)
oauth_consumer = oauth.Consumer(key=consumer_key, secret=consumer_secret)

def favourites_xls(opt):
	t = Twython(consumer_key, consumer_secret, access_token_key, access_token_secret)

	# Let's create an empty xls Workbook and define formatting
	wb = Workbook()
	ws = wb.add_sheet('0')
	
	# Column widths
	ws.col(0).width = 256 * 30
	ws.col(1).width = 256 * 30
	ws.col(2).width = 256 * 60
	ws.col(3).width = 256 * 30
	ws.col(4).width = 256 * 30

	# Stylez
	style_link = easyxf('font: underline single, name Arial, height 280, colour_index blue')
	style_heading = easyxf('font: bold 1, name Arial, height 280; pattern: pattern solid, pattern_fore_colour yellow, pattern_back_colour yellow')
	style_wrap = easyxf('align: wrap 1; font: height 280')

	# Headings in proper MBA spreadsheet style - Bold with yellow background
	ws.write(0,0,'Author',style_heading)
	ws.write(0,1,'Twitter Handle',style_heading)
	ws.write(0,2,'Text',style_heading)
	ws.write(0,3,'Embedded Links',style_heading)

	# Let's start at page 1 of your favourites because you know, it's a very good place to start
	count = 1
	pagenum = 1

	# Now, let's start an infinite loop and I don't mean the one with Apple's HQ
	while True:

		# Get your favourites from Twitter
		faves = t.get_favorites(page=pagenum)

		# If there's no favourites left from this page OR we've reached the page number specified by our user, do the Di Caprio and jump out
		if len(faves) == 0 or (opt != 'all' and pagenum > int(opt)):
			break

		# Programmers have been doing inception for ages before Nolan did. Let's go deeper and get
		# into another loop now
		for fav in faves:
			ws.write(count, 0, fav['user']['name'],style_wrap)
			ws.write(count, 1, fav['user']['screen_name'],style_wrap)
			ws.write(count, 2, fav['text'],style_wrap)
			links = fav['entities']['urls']
			i = 0
			for link in links:
				formatted_link = 'HYPERLINK("%s";"%s")' % (link['url'],"link")
				ws.write(count, 3+i, Formula(formatted_link), style_link)
				i += 1
			
			count += 1

		pagenum += 1

		

		# Now comes the all important part of the code. As my grandmother once told me, it is
		# important to get good sleep, and we shall precisely do that for 15 minutes because you know
		# Twitter mama hates rate-limit violators. We will use 12 requests for every 15 minutes just to be 
		# on the safe side
		if (pagenum % 12 == 0):
			time.sleep(900)

	
	# Now for the step that has caused untold misery and suffering to people who forget to do it at work
	wb.save('twitter_favourites.xls')

if __name__ == '__main__':
	
	if len(sys.argv) == 1:
		print 'Usage is python favourites_xls.py all OR python favourites_xls <number of pages of recent faves>'
	else:
		if (sys.argv[1].isdigit()):
			favourites_xls(sys.argv[1])
		else:
			if (sys.argv[1] == 'all'):
				favourites_xls(sys.argv[1])
			else:
				print 'Please use either all or a page number. This script is not Niels Bohr ok?'

	

