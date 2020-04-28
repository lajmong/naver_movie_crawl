# -*- coding: utf-8 -*-
import urllib.request
import json
import re
import pandas as pd
import numpy as np
import os
import sys
import requests
import openpyxl
from bs4 import BeautifulSoup

def search_movie(client_id, client_secret, title):
    '''Naver 검색 API를 이용해서 영화 제목을 치면 영화 고유값을 돌려줌.'''

    encText = urllib.parse.quote(title)
    url = "https://openapi.naver.com/v1/search/movie?query=" + encText  # json 결과
    request = urllib.request.Request(url)
    request.add_header("X-Naver-Client-Id", client_id)
    request.add_header("X-Naver-Client-Secret", client_secret)
    response = urllib.request.urlopen(request)
    rescode = response.getcode()

    if (rescode == 200):
        response_body = response.read()
        fix = response_body.decode('utf-8')
        fix = json.loads(fix)  # json 형태로 변환
        for each in fix['items']:
            if each['title'] == '<b>'+title+'</b>' and (each['pubDate'] == '2017' or each['pubDate'] == '2018'):
                full_link = each['link']  # link만 따로 뽑음.
                mov_num = re.search(r'code=(\d+)', full_link)  # 영화 고유 코드
                movie_id = mov_num.group(1)
                return movie_id
    else:
        return print("Error Code:" + rescode)


def movie_info(id):
    '''requests를 이용해서 영화별 네티즌 평점, 기자 평점, 장르, 상영시간, 관람등급을 뽑음.'''
    
    if type(id) != str:  # id가 None으로 나올 경우
        n_score, s_score, genre, time, age, nvideo = '-', '-', '-', '-', '-', '-'
    else:
        raw = requests.get('https://movie.naver.com/movie/bi/mi/basic.nhn?code=' + id)
        html = BeautifulSoup(raw.text, 'html.parser')
        for s in html.select('div.netizen_score div.star_score em'):  # 네티즌 평점
            n_score = s.text

        for s in html.select('div.special_score div.star_score em'):  # 기자 평점
            s_score = s.text

        info1 = html.select('dl.info_spec dd span')  # 장르, 상영시간
        genre = ''
        for g in info1[0].select('a'):  # 장르
            genre += g.text + ' '

        time = info1[2].text  # 상영시간

        info2 = html.select('dl.info_spec dd')  # 관람등급
        for a in info2[-1].select('a'):
            age = a.text
            break

        video = html.select('div.video div.title_area em')  # 비디오 갯수
        nvideo = video[0].text

    return n_score, s_score, genre, time, age, nvideo

def movie_info_sheet(start, end):
    movies = list(kobis['영화명'])[start:end]  # 10개씩만 출력 가능한듯
    movie_ids = []  # 영화 고유값 list

    for movie in movies:
        id = search_movie(client_id, client_secret, movie)
        movie_ids.append(id)

    # 시트에 넣을 값 지정
    for id in movie_ids:
        name = kobis.loc[start, '영화명']
        year = kobis.loc[start, '개봉year']
        month = kobis.loc[start, '개봉month']
        day = kobis.loc[start, '개봉weekday']
        country = kobis.loc[start, '대표국적']
        n_score, s_score, genre, time, age, nvideo = movie_info(id)  # 네이버 영화에서 정보 크롤링
        views = kobis.loc[start, '관객수']

        sheet.append([name, year, month, day, country,
        n_score, s_score, nvideo,genre, age, time, views])

        start+=1
    return wb.save('naver_movie.xlsx')


wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(['영화명', '개봉y', '개봉m', '개봉d', '국가', '네티즌', '기자', '동영상',
              '장르', '관람등급', '상영시간', '관람객수'])

client_id = ""
client_secret = ""

# 2018년 관객수 상위 300 영화 데이터 적재
kobis = pd.read_csv('kobis_300.csv')


movie_info_sheet(0, 10)
movie_info_sheet(10, 20)
movie_info_sheet(20, 30)
movie_info_sheet(30, 40)
movie_info_sheet(40, 50)
movie_info_sheet(50, 60)
movie_info_sheet(60, 70)
