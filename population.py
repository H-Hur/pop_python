import pandas as pd

# in_file: xls file
# in_sex: 성별 구분(현재 기능 없음, '전체' 반드시 필요)
# in_area_metro: 전국/광역시도 분류, in_area_base: 시군구
# age: 연령, 199 = 전체 연령, 100 = 100세 이상, 그 외 실제 연령(만 나이)
def read_pop_city_xls(in_file,in_year,in_sheet,in_area_metro=None):
    in_sex = '계' # 성별 조건은 엑셀 행 범위(65535) 최대에 걸려서 일단 비활성
    col_area='행정구역(시군구)별'
    col_sex='성별'
    col_age='연령별'
    col_pop=str(in_year)+' 년'
    print('reading',in_year,'년')
# xlsx 파일 읽기
    xl_read = pd.read_excel(in_file,sheet_name=in_sheet)
    xl = xl_read.loc[:,[col_area,col_sex,col_age,col_pop]]
# DataFrame 첫 번째 열에 광역시도 열 추가
    col_metro='광역시도'
    xl.insert(0, col_metro, '')
# 현행 광역단체 구분 정보
    metro = ['전국','서울특별시','부산광역시','대전광역시',\
    '대구광역시','광주광역시','인천광역시','울산광역시', \
    '경기도','강원도','충청북도','충청남도',\
    '경상북도','경상남도','전라북도','전라남도',\
    '세종특별자치시','제주특별자치도']
# 행정구역(시군구별) 열에서 세분화 된 지역 존재로 삭제할 영역
    base_large = ['수원시','성남시','고양시','안양시',\
    '안산시','용인시','천안시','전주시','포항시',\
    '통합창원시',\
    '상당구','흥덕구','청원구','서원구'] #청주는 청주+4개구로 자료가 돼 있어 구를 삭
# 광역시도 정보를 입력
# 행정구역(시군구별) 열에 광역시도 정보가 있으면 이를 광역시도에 복사
# 행정구역(시군구별) 열에 시군구 정보가 있으면 이전 광역시도 정보를 가져옴
    current_area=' '
    for i, area in enumerate(xl[col_area]):
        if area in metro and area != current_area:
            if current_area == in_area_metro: out = 'out'
            for j, area2 in enumerate(metro):
                if area == area2: 
                    current_area = area
        xl.at[i,col_metro] = current_area
# 행정구역(시군구별) 광역시도 정보가 있는 행 삭제
    for metro_ch in metro:
        metro_del = xl[xl[col_area] == metro_ch].index
        xl.drop(metro_del, inplace = True)
    for area_ch in base_large:
        area_del = xl[xl[col_area] == area_ch].index
        xl.drop(area_del, inplace = True)
#    print(xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '인천광역시'), col_area])

    xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '인천광역시'), col_area] = '인천광역시서구'
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '인천광역시'), col_area] = '인천광역시미추홀구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '인천광역시'), col_area] = '인천광역시동구'
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '인천광역시'), col_area] = '인천광역시중구'
    xl.loc[(xl[col_area] == '미추홀구') & (xl[col_metro] == '인천광역시'), col_area] = '인천광역시미추홀구'
# ~2017 남구, 2018~ 미추홀구                                      

    xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '광주광역시'), col_area] = '광주광역시서구'
    xl.loc[(xl[col_area] == '북구') & (xl[col_metro] == '광주광역시'), col_area] = '광주광역시북구'
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '광주광역시'), col_area] = '광주광역시남구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '광주광역시'), col_area] = '광주광역시동구'
                                      
    xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '대전광역시'), col_area] = '대전광역시서구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '대전광역시'), col_area] = '대전광역시동구'
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '대전광역시'), col_area] = '대전광역시중구'
                                      
    xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '부산광역시'), col_area] = '부산광역시서구'
    xl.loc[(xl[col_area] == '북구') & (xl[col_metro] == '부산광역시'), col_area] = '부산광역시북구'
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '부산광역시'), col_area] = '부산광역시남구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '부산광역시'), col_area] = '부산광역시동구'
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '부산광역시'), col_area] = '부산광역시중구'

    xl.loc[(xl[col_area] == '서구') & (xl[col_metro] == '대구광역시'), col_area] = '대구광역시서구'
    xl.loc[(xl[col_area] == '북구') & (xl[col_metro] == '대구광역시'), col_area] = '대구광역시북구'
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '대구광역시'), col_area] = '대구광역시남구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '대구광역시'), col_area] = '대구광역시동구'
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '대구광역시'), col_area] = '대구광역시중구'                                      
    xl.loc[(xl[col_area] == '북구') & (xl[col_metro] == '울산광역시'), col_area] = '울산광역시북구'
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '울산광역시'), col_area] = '울산광역시남구'
    xl.loc[(xl[col_area] == '동구') & (xl[col_metro] == '울산광역시'), col_area] = '울산광역시동구'
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '울산광역시'), col_area] = '울산광역시중구'
                                      
    xl.loc[(xl[col_area] == '중구') & (xl[col_metro] == '서울특별시'), col_area] = '서울특별시중구'
                                      
    xl.loc[(xl[col_area] == '남구') & (xl[col_metro] == '경상북도'), col_area] = '포항시남구'
    xl.loc[(xl[col_area] == '북구') & (xl[col_metro] == '경상북도'), col_area] = '포항시북구'

    xl.loc[xl[col_area] == '장안구', col_area] = '수원시장안구'
    xl.loc[xl[col_area] == '권선구', col_area] = '수원시권선구'
    xl.loc[xl[col_area] == '팔달구', col_area] = '수원시팔달구'
    xl.loc[xl[col_area] == '영통구', col_area] = '수원시영통구'

    xl.loc[xl[col_area] == '수정구', col_area] = '성남시수정구'
    xl.loc[xl[col_area] == '중원구', col_area] = '성남시중원구'
    xl.loc[xl[col_area] == '분당구', col_area] = '성남시분당구'

    xl.loc[xl[col_area] == '덕양구', col_area] = '고양시덕양구'
    xl.loc[xl[col_area] == '일산동구', col_area] = '고양시일산동구'
    xl.loc[xl[col_area] == '일산서구', col_area] = '고양시일산서구'

    xl.loc[xl[col_area] == '만안구', col_area] = '안양시만안구'
    xl.loc[xl[col_area] == '동안구', col_area] = '안양시동안구'
    xl.loc[xl[col_area] == '상록구', col_area] = '안산시상록구'
    xl.loc[xl[col_area] == '단원구', col_area] = '안산시단원구'

    xl.loc[xl[col_area] == '처인구', col_area] = '용인시처인구'
    xl.loc[xl[col_area] == '기흥구', col_area] = '용인시기흥구'
    xl.loc[xl[col_area] == '수지구', col_area] = '용인시수지구'

#    xl.loc[xl[col_area] == '상당구', col_area] = '청주시상당구'
#    xl.loc[xl[col_area] == '서원구', col_area] = '청주시서원구'
#    xl.loc[xl[col_area] == '흥덕구', col_area] = '청주시흥덕구'
#    xl.loc[xl[col_area] == '청원구', col_area] = '청주시청원구'

    xl.loc[xl[col_area] == '동남구', col_area] = '천안시동남구'
    xl.loc[xl[col_area] == '서북구', col_area] = '천안시서북구'

    xl.loc[xl[col_area] == '완산구', col_area] = '전주시완산구'
    xl.loc[xl[col_area] == '덕진구', col_area] = '전주시덕진구'

    xl.loc[xl[col_area] == '의창구'    , col_area] = '창원시의창구'
    xl.loc[xl[col_area] == '성산구'    , col_area] = '창원시성산구'
    xl.loc[xl[col_area] == '마산합포구', col_area] = '창원시마산합포구'
    xl.loc[xl[col_area] == '마산회원구', col_area] = '창원시마산회원구'
    xl.loc[xl[col_area] == '진해구'    , col_area] = '창원시진해구'

#    xl.loc[xl[col_area] == '오정구', col_area] = '부천시오정구'
#    xl.loc[xl[col_area] == '소사구', col_area] = '부천시소사구'
#    xl.loc[xl[col_area] == '원미구', col_area] = '부천시원미구'

    xl.loc[xl[col_area] == '세종특별자치시', col_area] = '세종시'


# 연령정보 수정
# 100세 이상 -> 100, '계' 삭제
# '세 이상' 포함하는 행 삭제
# '세' 글자 삭제
#  연령별 열 데이터 타입을 Flat으로 변경
    xl.loc[xl[col_age] == '100세 이상', col_age] = '100'
    del_age = xl[xl[col_age].str.contains('세 이상')].index
    xl.drop(del_age, inplace=True)
    del_age = xl[xl[col_age].str.contains('계')].index
    xl.drop(del_age, inplace=True)
    xl[col_age]=xl[col_age].str.replace(pat='세',repl='')
    xl = xl.astype({col_age:'int16'})

# DataFrame에서 성별, 광역시도, 시군구 조건으로 추출
# 만일 시군구 조건이 없으면 성별 조건으로만 추출
    sex = xl[col_sex] == in_sex
    metro = xl[col_metro] == in_area_metro
    if in_area_metro: xl = xl[metro & sex]
    else: xl = xl[sex]
# 추출한 데이터 프레임의 열 번호 리셋
    xl.reset_index(inplace=True, drop=True)
    xl.rename(columns = {col_pop:'pop'},inplace=True)
    print('시군구 수:',xl.shape[0])
    return xl
