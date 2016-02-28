import re
import csv
import time
import requests
import MyLines2
from bs4 import BeautifulSoup

def main():
    targetURL = 'http://stackoverflow.com'
    proxiesFile = 'proxies.dat'
    dataFileName = 'data.csv'
    prodsURLsFile ='StackOverFlow URLs'

    headers = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
               "Accept-Encoding": "gzip,deflate,sdch",
               "Accept-Language": "en-US,en;q=0.8",
               "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36"}

    proxies = MyLines2.loadProxies(proxiesFile)
    MyLines2.removeDataFile(dataFileName)
    profile_urls = MyLines2.loadFileIntoLIst(prodsURLsFile)
    profile_urls = set(profile_urls)

    oFile = open(dataFileName, "ab")
    writer = csv.writer(oFile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)

    headers_row = ['UserName',
                   'URL',
                   'Location',
                   'Age',
                   'MemberSince',
                   'AccountCount',
                   'TagCount',
                   'VotesCast',
                   'Percentile',
                   'Qcount',
                   'Acount',
                   'QARatio',
                   'Date',
                   'ReputationDay',
                   'UpVoteCount',
                   'AcceptCount',
                   'DownVoteCount',
                   'AwardedGold?',
                   'Awarded Silver?',
                   'Awarded Bronze?',
                   'Other Awarded?',
                   'Area Awarded',
                   'Bounty Offered',
                   'Bounty Earned',
                   'Area of Bounty']

    writer.writerow(headers_row)

    counter = 0
    #3.Visit each product and scrape data
    for profile_url in profile_urls:
        user_name = ''
        url = ''
        location = ''
        age = ''
        member_since = ''
        account_count = ''
        tag_count = ''
        votes_cast = ''
        percentile = ''
        q_count = ''
        a_count = ''
        qa_ratio = ''
        date = ''
        reputation_day = ''
        up_vote_count = ''
        accept_count = ''
        down_vote_count = ''
        awarded_gold = ''
        awarded_silver = ''
        awarded_bronze = ''
        other_awarded = ''
        area_awarded = ''
        bounty_offered = ''
        bounty_earned = ''
        area_of_bounty = ''

        req = request_page(profile_url, headers, proxies)
        if req.status_code == 200:
            counter += 1

            print 'Scraping::' + profile_url + ' #' + str(counter)
            page_source = req.text
            if page_source:
                soup = BeautifulSoup(page_source)

                url = profile_url

                #User Name
                user_name_tag = soup.find('h1', id='user-displayname')
                if user_name_tag:
                    user_name = user_name_tag.text.strip()

                #Location
                location_tag = soup.find('td', {'class': 'label adr'})
                if location_tag:
                    location = location_tag.text.strip()

                #Age
                age_tag = soup.find('td', text='age')
                if age_tag:
                    age_row = age_tag.find_parent()
                    if age_row:
                        age_arr = [x for x in age_row.findChildren() if re.match('(\d+)', x.text)]
                        if age_arr:
                            age = age_arr[0].text

                #Member Since
                member_since_tag = soup.find('td', text='member for')
                if member_since_tag:
                    member_since_row = member_since_tag.find_parent()
                    if member_since_row:
                        mem_arr = [x for x in member_since_row.findChildren() if x.attrs.has_key('title')]
                        if mem_arr:
                            member_since = mem_arr[0].text

                #Accounts Count
                accounts_count_tag = [x.findChildren() for x in soup.find_all('a') if 'Accounts' in x.text]
                if accounts_count_tag:
                    for x in accounts_count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    account_count = y.text
                                    break

                #Tags Count
                tags_count_tag = [x.findChildren() for x in soup.find_all('a') if 'Tags' in x.text]
                if tags_count_tag:
                    for x in tags_count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    tag_count = y.text
                                    break

                #Tags Count
                count_tag = [x.findChildren() for x in soup.find_all('h1') if 'Votes Cast' in x.text]
                if count_tag:
                    for x in count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    votes_cast = y.text
                                    break

                #QCount
                count_tag = [x.findChildren() for x in soup.find_all('h1') if 'Question' in x.text]
                if count_tag:
                    for x in count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    q_count = y.text
                                    break

                #ACount
                count_tag = [x.findChildren() for x in soup.find_all('h1') if 'Answers' in x.text]
                if count_tag:
                    for x in count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    a_count = y.text
                                    break

                #ACount
                count_tag = [x.findChildren() for x in soup.find_all('h1') if 'Reputation' in x.text]
                if count_tag:
                    for x in count_tag:
                        if len(x) > 0:
                            for y in x:
                                if y.attrs.has_key('class') and 'count' in y.attrs['class']:
                                    reputation_day = y.text
                                    break

                qa_ratio = float(q_count.replace(',', '')) / float(a_count.replace(',', ''))
                qa_ratio = str(qa_ratio)

                date = str(time.strftime("%d-%b-%y"))

                if user_name:
                    row = [user_name.strip().strip().encode('utf-8'),
                           url.strip().strip().encode('utf-8'),
                           location.strip().strip().encode('utf-8'),
                           age.strip().encode('utf-8'),
                           member_since.strip().encode('utf-8'),
                           account_count.strip().encode('utf-8'),
                           tag_count.strip().encode('utf-8'),
                           votes_cast.strip().encode('utf-8'),
                           percentile.strip().encode('utf-8'),
                           q_count.strip().encode('utf-8'),
                           a_count.strip().encode('utf-8'),
                           qa_ratio.strip().encode('utf-8'),
                           date.strip().encode('utf-8'),
                           reputation_day.strip().encode('utf-8'),
                           up_vote_count.strip().encode('utf-8'),
                           accept_count.strip().encode('utf-8'),
                           down_vote_count.strip().encode('utf-8'),
                           awarded_gold.strip().encode('utf-8'),
                           awarded_silver.strip().encode('utf-8'),
                           awarded_bronze.strip().encode('utf-8'),
                           other_awarded.strip().encode('utf-8'),
                           area_awarded.strip().encode('utf-8'),
                           bounty_offered.strip().encode('utf-8'),
                           bounty_earned.strip().encode('utf-8'),
                           area_of_bounty.strip().encode('utf-8')]

                    writer.writerow(row)
                    print row

                else:
                    print 'No name found for::' + profile_url

    oFile.close()
    print 'Done'

def request_page(target_url, headers, proxies):
    if proxies and len(proxies) > 0:
        rndProxy = 'http://' + MyLines2.PickRandomProxy(proxies)
        #print 'request using proxy::' + rndProxy
        req = requests.get(target_url, headers=headers, proxies={'http': rndProxy})
    else:
        req = requests.get(target_url, headers=headers)
    return req

if __name__ == '__main__':
    main()
    print 'Done'