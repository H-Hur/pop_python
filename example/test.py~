#jsonUrl = 'http://kosis.kr/openapi/statisticsData.do?method=getList&apiKey=OGM2N2RlYTNmMjQ2MGY3NWEyYzg5MGU2NmJjYTI4Zjg=&format=json&jsonVD=Y&userStatsId=ox8eb1f9/101/DT_1B040A3/2/1/20200505230018&prdSe=M&newEstPrdCnt=1'

#apiKey=NDc0YWJhYjY2MGJkMzQxYjM4YjkzMjljNDgxMmRkY2Q=
url=\
'https://kosis.kr/openapi/statisticsBigData.do?method=getList&apiKey=NDc0YWJhYjY2MGJkMzQxYjM4YjkzMjljNDgxMmRkY2Q=&format=xls&userStatsId=gjgusdh/101/DT_1B040M1/3/1/20220220214135&prdSe=Y&startPrdDe=2021&endPrdDe=2021'
#https://kosis.kr/openapi/statisticsData.do?method=getList&apiKey=OGM2N2RlYTNmMjQ2MGY3NWEyYzg5MGU2NmJjYTI4Zjg=&format=json&jsonVD=Y&userStatsId=ox8eb1f9/101/DT_1B040A3/2/1/20200505230018&prdSe=M&newEstPrdCnt=1

import pandas as pd
import requests

try:
    response = requests.get(url)
    if response.status_code == 200:
#        population = pd.read_json(jsonUrl)
        population = pd.read_xls(jsonUrl)
    else:
        response.close()
except Exception as e:
    print(e)

population.head()
#print(population[{'PRD_DE': '202201', 'C1_OBJ_NM':'노원구'}]r
