import win32com.client #보안문서를 열 수 있는 모듈
import xlswriter
from datetime import datetime
#datetime.now().strftime("%Y%m%d") #하면 오늘날짜

def pos(ws, i, h):
    today = int(datetime.now().strftime("%Y%m%d"))
    date = []
    position = []
    global copy_date
    copy_date = []
    df = import
    
    #값을 넣기위한 엑셀의 좌표를 얻기 위해 반복문 해당 엑셀의 행렬이 어디인지
    while i:
        print(i)
        date__ = ws.Range(xlswriter.utility.xl_xol_to_name(i) + h).value
        #숫자를 알파벳으로 변환 , 
        if date__ == None:  #값이 없는 경우 건너뛰기
            pass
        else:  #값이 있다면 해당 년월에 맞는 위치에 값을 넣어야 하기 때문에
            a_temp = date__.strftime("%Y%M%d")
            if int(a_temp) >= 20200601 and int(a_temp) < today:
                date.append(a_temp)
                position.append(xlswriter.utility.xl_xol_to_name(i))
            if int(date__.strftime("%y%m%d")) == today:
                break
            # 해당 위치의 엑셀 알파벳(col)을 append 하고 밑에서 return
        i = i+1
        if(i == df+40):
            break
    copy_date = date
    return position


def pos_(ws, i, n):
    start_day = int(datetime.now(0.strftime("%y%m%d") + "01")) #영업일 기준 시작일
    today = int(datetime.now().strftime("%Y%m%d"))
    date = []
    position = []
    df = i
    flag = True
    while 1:
        date__ = ws.Range(xlswriter.utility.xl_xol_to_name(i) + n).value  
        if type(date__) ! = int:
            if type(date__) ==float:
                date__ = int(date__)
            else:
                pass
        # 년월의 값을 가져오기때문에 int인지 확인, float인 값들도 있기때문에 타압을 int로 통합시켜준다.

        if str(date__) == str(start_dady):
            while 1:
                date__ = ws.Range(xlswriter.utility.xl_xol_to_name(i) + n).value
                a_temp = date__
                if str(date__) != '2':
                    flag = False
                    break                

                if str(int(a_temp)) in copy_date:
                    date.append(a_temp)
                    position.append(xlswriter.utility.xl_col_to_name(i))
                # 값을 복사하는 엑셀과 붙여넣는 엑셀이 있음 붙여넣을 엑셀의 알파벳을 미리 copy_date 리스트에 담아두고 복사해온 엑셀의 값(날짜)이 리스트 안에 있다면 넣어도 좋은 값이라 판단
                if int(date__) == today-1:
                    flag = False
                    break

                i = i+1

        if flag == False:
            break
        i = i+1
        if(i == df+40):
            print("어디서부터 반복문을 시작할지 정하는 구문 이 경우 start_number를 바꿔보셈")
            break
    return position



def input_(col_start_copy, col_end_copy , row_start_copy, row_end_copy , col_start , col_end, row_start, row_end, ws_copy, ws_paste):
    for i in range(0, len(col_end)):
        copy = []
        paste_ = []
        for j in range(0, len(row_start)):
            copy_.append(ws_copy.Range(col_start_copy[i] + row_start_copy[j] +  ':' + col_end_copy[i] + row_end_copy[j]))
            paste_.append(ws_paste.Range(col_start[i] . row_start[j] +  ':' + col_end[i] + row_end[j]))

        for r in range(0, len(copy_)):
            for k in range(0, len(copy_[r])):
                paste_[r][k].value = copy_[r][k].value
            # 해당 범위에 복사해온 값을 값복사 한다.
            # 더 편한 방법을 찾았다 밑에서 나옴


def family(n):
    path_paste = 'D:\\sy1909\\download\\ 폴더명 \\ 파일이름.xlsx'
    excel_file_paste = excel.Workbooks.Open(pate_paste) # 해당경로의 엑셀파일 오픈
    ws_paste = excel_file_paste.Worksheets('붙여넣을_시트이름')
    ws_copy = excel_file_copy.Worksheets('복사할_시트이름') 

    pos_copy = pos_(ws_copy , 20 , '51')
    #pos_copy = pos_(열린 시트의 변수명 , 시작 컬럼 위치(1=A,2=B) , 행의 수 str형식으로)
    #복사할 곳의 열 좌표를 구하기 위함
    col_start_copy = pos_copy
    col_end_copy = pos_copy
    row_start_copy = ['52', '60', '65']
    row_end_copy =  ['58', '63', '66']
    # 붙여넣을 엑셀의 열,행 범위 위아래로 52-58  60-63 65-66 식으로 값이 들어감 열은 위에서 구한값

    col_start = pos_paste
    col_end = pos_paste
    row_start = ['27', '46', '56']
    row_end =   ['33', '49', '57'] 
    # 복사할 엑셀의 열,행 범위들 
    # 간소화 할 수 있음

    input_(col_start_copy, col_end_copy , row_start_copy, row_end_copy , col_start , col_end, row_start, row_end, ws_copy, ws_paste)
    # 구한 값을 토대로 값을 입력하는 함수


def open_dudtlf(): #해당 파일을 여는 함수
    global excel
    excel = win32com.client.Dispatch("excel.Application")
    #엑셀이라는 프로그램을 할당한다.
    excel.Visible = True 
    #실행되면 엑셀을 실제로 띄워서 보여준다. 속도를 높이려면 닫아놓는 방법도 있음
    global path_copy
    path_copy = '해당 엑셀파일경로명\\파일명.xlsx'
    global excel_file_copy
    excel_file_copy = excel.Workbooks.Open(path_copy)

    #위의 변수들을 적절히 활용해서 여러개의 엑셀 파일과 시트들을 열고 닫고 수정할 수 있음

def ge(cell , workday):
    ws_copy = excel_file_ge.Worksheets('시트이름')
    #새로운 엑셀을 오픈하고 해당 시트 불러오기
    ws_1.Range('A1' , ws_1.Cells(cell[0] , cell[1])).Copy()
    #특정 셀좌표에서 이런식으로 cell 로 불러온 좌표 사용가능
    ws_copy.Range("E6").PasteSpecial(Paste = -4163)
    #복사한값을 해당 좌표를 시작점으로 붙여넣는 함수 -4163은 값복사
    #   -4123 수식복사
    #   PasteSpecial 관련 자료 찾아봐도 될듯

    for border_id in range(7,13):
        ws_copy.Range("E6" , "ck" + str(cell[0]+5)).Borders(border_id).LineStyle = 1 
        # linesyle은 어떤 테두리 스타일을 쓸 것인가
        ws_copy.Range("E6" , "ck" + str(cell[0]+5)).Borders(border_id).weight = 2
        # 두께 지정 2가 보통 두께인듯
    #range(7,13)은 세로 가로 border_id에 대한 값으로 세로 가로 중간 등 테두리를 설정해줌 반복돌려서 전부 하면 전체 테두리 설정이 됨

    pivotCount = ws_copy.PivotTables().Count # 피벗테이블 개수를 구하는 듯 새로고침 하기 위함
    for j in range(pivotCount , pivotCount+1):
        ws_copy.PivotTables(j).PivotCache().Refresh() # 피벗테이블 새로고침 지금은 1개만


def open_raw(path):
    global path_paste_daily
    path_paste_daily = path
    global excel_file_raw
    excel_file_raw = excel.Workbooks.Open(path_paste_daily)
    global ws_1ws_1 = excel_file_raw.Worksheets('시트이름')
    cell = []
    cell.append(ws_1.UsedRange.Rows.Count)
    cell.append(ws_1.UsedRange.Columns.Count)
    # 해당 엑셀파일에 이용되고있는(값이 들어가있는) row와 columb 을 받아온다.

    return cell    




if __name__ == '__main__':
    open_dudtlf() # 먼저 엑셀파일과 등등을 열어놓음
    start_number = 222
    family(start_number)
    open_daily()