# -*- coding: utf-8 -*-

import re
import os
import ujson
import shutil
import requests
from lxml import html

def download_file(url, filename):
    file_req = requests.get(url, stream=True)
    if file_req.status_code == 200:
        with open(filename, 'wb') as f:
            file_req.raw.decode_content = True
            shutil.copyfileobj(file_req.raw, f)

def extract_data(url):

        download_subs = True

        req = requests.get(url)
        if req.status_code == 200:
            content = req.text

            #load movie data.
            movie_data = ujson.loads(content)

            #Extract movie data from json
            IMDBRating = movie_data['imdbRating'] if movie_data.has_key('imdbRating') else 'N/A'
            Poster = movie_data['Poster'] if movie_data.has_key('Poster') else 'N/A'
            Year = movie_data['Year'] if movie_data.has_key('Year') else 'N/A'
            Title = movie_data['Title'] if movie_data.has_key('Title') else 'N/A'
            Genre = movie_data['Genre'] if movie_data.has_key('Genre') else 'N/A'
            Actors = movie_data['Actors'] if movie_data.has_key('Actors') else 'N/A'
            Awards = movie_data['Awards'] if movie_data.has_key('Awards') else 'N/A'
            imdbID = movie_data['imdbID'] if movie_data.has_key('imdbID') else 'N/A'

            if imdbID != 'N/A':
                cert_url = 'http://www.imdb.com/title/%s/parentalguide?ref_=tt_stry_pg#certification' % imdbID
                cert_req = requests.get(cert_url)
                if cert_req:
                    cert_body = cert_req.text
                    lhtml = html.fromstring(cert_body)
                    parent_guide = lhtml.xpath('.//*[@id="swiki.2.1"]/descendant::*/text()')
                    if parent_guide:
                        parent_guide = [re.sub('\n',' ',x).strip() for x in parent_guide]
                        parent_guide = [x for x in parent_guide if x != '']
                        parent_guide = '; '.join(parent_guide)
                        pFile = open('ParentGuide.csv', 'ab')
                        pFile.write(movie_name.replace(',', ';') + ',' + parent_guide.replace(',', ';') + '\r\n')
                        pFile.close()
            try:
                #Create Genre folder and downlaod posters into them.
                if Genre != 'N/A':
                    Genre = Genre.split(',', 1)[0]
                    if not os.path.exists('Posters/' + Genre):
                        os.makedirs('Posters/' + Genre)
                    if Poster != 'N/A' and Title != 'N/A':
                        Title = re.sub('<|>|:|"|/|\|\?|\*', '', Title).strip()
                        file_name = 'Posters/' + Genre + '\\' + Title + '.jpg'
                        if IMDBRating != 'N/A' and  Year != 'N/A':
                            file_name = 'Posters/' + Genre + '\\' + Title + ' - ' + Year + ' - ' + IMDBRating + '.jpg'
                        download_file(Poster, file_name)

            except:
                #Report failed posters and skip this one.
                f = open('Failed.txt', 'ab')
                f.write(Title + '\n')
                f.close()
                return

            #Use Title to get list of srt files on subscen, (if YIFY on it select, else pick the first one.)
            Choosed = list()
            if Title != 'N/A' and download_subs:
                sub_title = Title.lower().replace(' ', '-').strip()
                ara_subs_url = 'http://subscene.com/subtitles/%s/arabic' % sub_title
                sub_req = requests.get(ara_subs_url)
                if sub_req.status_code == 200:
                    lhtml = html.fromstring(sub_req.text)
                    links = lhtml.xpath('.//td/a')
                    for link in links:
                        anchors_texts = link.xpath('.//span/text()')
                        name = [re.sub('\n|\r|\t', '', x).strip() for x in anchors_texts]
                        name = '-'.join(name)
                        if not name == '':
                            href = 'http://subscene.com' + link.xpath('@href')[0]
                            if 'YIFY' in name or 'yify' in name:
                                Choosed.append(name)
                                Choosed.append(href)
                                print '\t' + name + ' ' + href

                    if len(Choosed) == 0 and len(links) > 0:
                       link = links[0]
                       anchors_texts = link.xpath('.//span/text()')
                       name = [re.sub('\n|\r|\t', '', x).strip() for x in anchors_texts]
                       name = '-'.join(name)
                       if not name == '':
                           href = 'http://subscene.com' + link.xpath('@href')[0]
                           Choosed.append(name)
                           Choosed.append(href)
                           print '\t' + name + ' ' + href

            #Download the sub titles files on srt folder.
            if len(Choosed) == 2:
                if not os.path.exists('srt'):
                    os.makedirs('srt')
                    os.makedirs('srt/zip')

                file_name = re.sub('<|>|:|"|/|\|\?|\*', '', Choosed[0]).strip()
                srt_url = Choosed[1]
                srt_req = requests.get(srt_url)
                if srt_req.status_code == 200:
                    s = html.fromstring(srt_req.text)
                    button = s.xpath('//*[@id="downloadButton"]/@href')
                    if button:
                        print '\t' + 'Download srt file for ' + Title
                        download_link = 'http://subscene.com' + button[0]
                        download_file(download_link, 'srt/zip/' + file_name + '.zip')

#Read all unwathched movies paths from UnWatchedDirectoriesPaths.txt
state_dict = dict()
UnWatchedDirectories = 'UnWatchedDirectoriesPaths.txt'
with open(UnWatchedDirectories) as f:
    skipped_movies = list()
    dirs = f.read().splitlines()

    for dir in dirs:
        #It means I commented it out.
        if dir.startswith('#'):
            continue #skip this one

        print 'Check dir:' + dir
        movies_directory_path = dir

        #All paths here now
        movies = os.listdir(movies_directory_path)
        movies_count = len(movies)

        #State
        state_dict['TotalMoviesCount'] = movies_count + state_dict['TotalMoviesCount'] \
            if state_dict.has_key(('TotalMoviesCount')) \
            else movies_count

        print '\t' + 'Unwatched Movies Count: ' + str(movies_count)

        #Get every movie name and year then create an api url from it then pass to extract_data or to skipped_movies file.
        api_url = 'N/A'
        for movie in movies:
            try:
                movie_name = movie
                movie_name = re.sub('\[\w+\]', '', movie_name).strip()
                movie_year = re.search('\(*(\d{4})\)*', movie)
                if movie_year:
                    year = movie_year.group(1)
                    movie_name = movie_name.replace(movie_year.group(0), '').strip()
                    print '\t' + movie_name + ' ' + year
                    movie_name = movie_name.replace(' ', '+').strip()
                    api_url = 'http://www.omdbapi.com/?t=%s&y=%s&plot=full&r=json' % (movie_name, year)
                    extract_data(api_url)
                else:
                    skipped_movies.append(movie_name)
            except:
                #Report failed posters and skip this one.
                f = open('Failed.txt', 'ab')
                f.write(movie + '\n')
                f.close()

        #Print state
        print state_dict

        #Print and log skipped.
        print 'Skipped Count:' + str(len(skipped_movies))
        for m in skipped_movies:
            print '\t' + m
            f = open('Skipped.txt', 'ab')
            f.write(m + '\n')
            f.close()
