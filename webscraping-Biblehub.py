import random
from urllib.request import urlopen
from bs4 import BeautifulSoup
from urllib.request import urlopen, Request

# webpage = 'https://biblehub.com/asv/john/1.htm'
# headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
# req = Request(webpage, headers=headers)

# webpage = urlopen(req).read()

# soup = BeautifulSoup(webpage,'html.parser')

# print(soup.title.text)

# verses_list = soup.findAll('p',class_='reg')

# for verse in verses_list:
#     print(verse.text)
#     input()

# #for verse in verses_lsit:
#     print(verse.text)

# verse_list = [v.text.split('.') for v in verses_list ]
# print(verse_list)

# print(random.choice(random.choice(verse_list)))


chapters = list(range(1,22))
random_chapter = random.choice(chapters)

if random_chapter < 10:
    random_chapter = '0'+str(random_chapter)

else: 
    random_chapter = str(random_chapter)

url = 'https://ebible.org/asv/JHN' + random_chapter + '.htm'
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.3'}
req = Request(url, headers=headers)

webpage = urlopen(req).read()

soup = BeautifulSoup(webpage,'html.parser')

print(soup.title.text)

page_verses = soup.findAll('div', class_='p')

for verses in page_verses:
    verse_list = verses.text.split(".")

mychoice = random.choice(verse_list[:5])

verse = f'Chapter: {random_chapter} Verse: {mychoice}'
print(verse)

import keys
from twilio.rest import Client 

client = Client(keys.accountSID, keys.authToken)

TwilioNumber = '+15074163593'

mycellphone = '+16613901505'

textmessage = client.messages.create(to=mycellphone, from_=TwilioNumber, body=verse)

