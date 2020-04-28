import pandas as pd
import numpy as np
from datetime import datetime

# 데이터 적재
kobis = pd.read_excel('/Users/yuheunkim/Downloads/KOBIS_연도별박스오피스_2020-04-25.xlsx')

# 컬럼값 지정, 인덱스 재설정
kobis.columns = np.array(kobis.loc[3])
kobis = kobis.loc[4:].reset_index().drop('index', axis=1)

# 2018 관객수 상위 300 영화
kobis['개봉일'] = pd.to_datetime(kobis.개봉일)
kobis_300 = kobis[kobis['개봉일'].dt.strftime('%Y') == '2018'][:300]
kobis_300 = kobis_300.reset_index().drop('index', axis=1)

# 컬럼값 지정
hw_columns = ['영화명', '개봉일', '관객수', '대표국적']
kobis_300 = kobis_300[hw_columns]

# 날짜 따로 빼기
kobis_300['개봉year'] = kobis_300['개봉일'].dt.strftime('%Y')
kobis_300['개봉month'] = kobis_300['개봉일'].dt.strftime('%m')
kobis_300['개봉weekday'] = kobis_300['개봉일'].dt.strftime('%w')  # 0: sunday
kobis_300.drop('개봉일', axis=1, inplace=True)

# 관객수로 sorting
kobis_300 = kobis_300.sort_values(by='관객수', ascending=False)
kobis_300.to_csv('kobis_300.csv')  # csv로 저장
