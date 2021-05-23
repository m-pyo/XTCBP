from numpy import NaN
import pandas as pd
import os
import glob
import datetime

from distutils.dir_util import copy_tree

#대상파일 위치
XLSX_PATH = './xlsxFiles'

#변환 파일 관련
RUN_DATE = datetime.date.today().strftime('%y%m%d')
CSV_PATH = './' + RUN_DATE + '/Original/csv'
IMAGE_PATH = './' + RUN_DATE + '/Original/Images'

USED_CAR_FOLDER = 'UsedCar'
NEW_CAR_FOLDER = 'NewCar'
GOONET_FOLDER = 'Goonet'

USED_CAR_TYPE = 1
NEW_CAR_TYPE = 2
GOONET_TYPE = 3

CSV_STORE_FILE_NAME = '_store'
CSV_PLAN_FILE_NAME = '_plan'
CSV_CAR_FILE_NAME = '_car'

SKIP_ROWS = 2  # 차량 정보 취득시 생략하는 라인수

#설정파일관련
SET_FILE = './変換設定.xlsx'

def getSetData() -> dict:
    """설정파일의 내용을 취득
    
    Returns:
       폴더명:{타입: 회사ID}형태의 딕셔너리를 반환
    """
    result = {}
    setData = pd.read_excel(SET_FILE, header = 1, usecols = [1,2,3], engine='openpyxl').values.tolist()
    
    for data in setData:
        key = data[0]
        type = data[1]
        company_id = data[2]
        
        if key in result: 
            result[key][type] = company_id
        else:
            result[key] = {type: company_id}
    
    return result

def getFolderList() -> list:
    """폴더 리스트 취득
    
    Returns:
       대상폴더의 폴더내역을 반환  
    """
    return os.listdir(XLSX_PATH)

def getXlsxFileList(folderName: str) -> list:
    """xlsx 파일 리스트를 취득
    
    Args:
        folderName: 폴더명
        
    Returns:
       .xlsx 형식의 파일 리스트   
    """
    
    path = f'{XLSX_PATH}/{folderName}'
    
    try:
        result = [fileName for fileName in os.listdir(path) if fileName.endswith('.xlsx')] 
    except:
        print('처리실패:' +  folderName)
        result = []
        
    return result 

def getFilePath(folderName:str, extension:str = "" ) -> list:
    """파일 패스 취득(이미지에서 활용)

    Args:
        folderName: 파일명
        extension: 확장자명
        
    Returns:
        unix주소형태의 파일 목록반환
    """
    
    path = f'{XLSX_PATH}/{folderName}/'
    return [path for path in glob.glob(path + extension, recursive = True) if os.path.isfile(path)] 


def getSheetList(folderName: str, fileName: str) -> list:
    """xlsx의 시트 리스트 취득(차후 시트명으로 판단하게 될시 필요함)
    
    Args:
        folderName: 폴더명
        fileName: 취득할 xlsx파일명(.xlsx포함)
    
    Returns: 
        시트명을 리스트로 반환
    """
    
    folderPath = f'{XLSX_PATH}/{folderName}'
    xlsxPath = f'{folderPath}/{fileName}'
    
    wb = pd.read_excel(xlsxPath, sheet_name = None, engine='openpyxl')
    return [*wb]

def getTypeToFolder(fileType: int) -> str:
    """저장될 폴더 타입을 반환
    
    Args: 
        fileType: 파일 타입(1,2,3)
        
    Returns:
        폴더명 반환
    """
    
    if fileType == 1:
        result = USED_CAR_FOLDER
    elif fileType == 2:
        result = NEW_CAR_FOLDER
    elif fileType == 3:
        result = GOONET_FOLDER
    else:
        result = False
        
    return result

def getCompanyIds(folderName: str) -> list:
    """회사ID 정보 취득 
    
    Args:
        forderName: 파일명
        
    Returns:
        [{타입:회사ID}] 의 형태로 회사ID반환
    """
    try:
        setData = getSetData()    
        result = setData[folderName]
    except:
        result = []
        
    return result


def createOriginalDir() -> bool:   
    """변환후 저장 파일 경로 생성
    """     
    
    createDirs = [
        CSV_PATH + '/' + USED_CAR_FOLDER,
        CSV_PATH + '/' + NEW_CAR_FOLDER,
        CSV_PATH + '/' + GOONET_FOLDER,
        IMAGE_PATH
    ]
    
    try: 
        for path in createDirs:
            #CSV 디렉토리 생성
            os.makedirs(path, exist_ok=True) #옵션으로 이미 존재할때는 넘어감
        return True
    except:
        return False
  
 
def getXlsxData(path: str) -> dict:
    """xlsx 데이터 취득
    
    Args: 
        파일 경로
    
    Returns:
        취득된 데이터 정보(DataFrame형태)
    """
    
    wb = pd.read_excel(path, header = None, sheet_name = None, engine='openpyxl')
    return wb

def getXlsxFilePath(fdn: str, fn: str) -> str:
    """xlsx 파일 경로 반환
    
    Args:
        fdn: 폴더명
        fn: 파일명
        
    Returns:
        파일경로
    """
    
    return f'{XLSX_PATH}/{fdn}/{fn}'

def getXlsxFolderPath(fdn: str) -> str:
    """폴더 경로 반환
    
    Args:
        fdn: 폴더명
        
    Returns:
        파일경로
    """
    
    return f'{XLSX_PATH}/{fdn}'

def getXlsxFilePath(fdn: str, fn: str) -> str:
    """xlsx 파일 경로 반환
    
    Args:
        fdn: 폴더명
        fn: 파일명
        
    Returns:
        파일경로
    """
    
    return f'{XLSX_PATH}/{fdn}/{fn}'

def getCsvFilePath(fdn: str,fn: str) -> str:
    """폴더 경로 반환
    
    Args:
        fdn: 폴더명
        
    Returns:
        파일경로
    """
    
    return f'{CSV_PATH}/{fdn}/{fn}'


def beforeXlsxCheck(fileList: list, companyIds: dict) -> bool:
    """변환전 xlsx파일 설정과 xlsx파일 개수, 논리확인
    
    Args:
        fileList: 파일 리스트
        companyIds: 회사 아이디 리스트
    
    Returns:
        에러여부 반환, 에러가 있는경우 True
    """
    
    if len(companyIds) == 0:
        print('処理失敗[未登録会社]：' + folderName + '/' + ','.join(fileList[:])) 
        return True
    
    if len(fileList) == 0 or len(fileList) > 3:
        print('処理失敗[xlsxファイル確認必要]：' + folderName + '/' + ','.join(fileList[:])) 
        return True
        
    elif len(fileList) != len(companyIds) :
        print('処理失敗[xlsxファイル件数相違]：' + folderName + '/' + ','.join(fileList[:])) 
        return True

    return False


# CSV파일생성
def createCsv(xlsxPath: str, createDir: str, companyId: str) -> bool:        
    try: 
        wb = getXlsxData(xlsxPath)
        sheetList = [*wb]
        
        for i in range(1,4):
            #시트 순서대로 처리(시트명 고정시 시트명으로 변경)
            if i == 1:
                csvName = CSV_STORE_FILE_NAME
            elif i == 2:
                csvName = CSV_PLAN_FILE_NAME
            elif i == 3:
                csvName = CSV_CAR_FILE_NAME    
                
            wb[sheetList[i]].to_csv(f'{CSV_PATH}/{createDir}/{companyId}{csvName}.csv', header = None,index = False, encoding = 'utf-8-sig')

        return True
    except:
        return False


def getCarIds(fileType: str,companyId:str) -> list:
    
    result = {}
    
    path = f'{CSV_PATH}/{getTypeToFolder(fileType)}/{companyId}_car.csv'
    csvData = pd.read_csv(path, index_col = None, header = None)
    
    #점포, 차량id 취득(nan 값 제외)
    carIdList = csvData[[1,3]].dropna().values.tolist()
    
    for item in carIdList[SKIP_ROWS:]:
        clientId = item[1] if fileType == USED_CAR_TYPE else companyId
        carId = item[0]
        
        result[carId] = clientId
        
    return result

def imageCopy(folderName, companyId, fileType):
    carIds = getCarIds(fileType, companyId)
    for carId, client in carIds.items():
        imageList = getFilePath(folderName,f'**/{carId}/*')
        
        if len(imageList) < 1:
            continue
        
        pathLists = ['/'.join(item.split('/')[:-1]) for item in imageList]
        pathListSet = set(pathLists)
        if len(pathListSet) > 1:
            print('処理失敗[イメージパス修復（車輌ID）]：' + ' '.join(pathListSet)) 
            continue
        
        copy_tree(pathLists[0], f'{IMAGE_PATH}/{companyId}/{client}/{carId}')
    
# 메인로직 
if __name__ == "__main__":
    
    print('start')
    createOriginalDir()
    
    for folderName in getFolderList():
        fileList = getXlsxFileList(folderName)
        companyIds = getCompanyIds(folderName)
        
        if beforeXlsxCheck(fileList,companyIds):
            continue
        
        try:
            if len(fileList) == 1:
                fileType =list(companyIds.keys())[0]
                fileName = fileList[0]
                companyId = companyIds[fileType]
                
                xlsxPath = getXlsxFilePath(folderName,fileName)
                
                createCsv(xlsxPath, getTypeToFolder(fileType), companyId)
                
                #게재점의 경우 종료
                if (fileType == GOONET_TYPE):
                    continue
                
                imageCopy(folderName, companyId, fileType)
               
                    
            else:
                for fileName in fileList:
                    if '中古' in fileName:
                        fileType = USED_CAR_TYPE
                        companyId = companyIds[fileType]
                        xlsxPath = getXlsxFilePath(folderName,fileName)
                    elif '新車' in fileName:
                        fileType = NEW_CAR_TYPE
                        companyId = companyIds[fileType]
                        xlsxPath = getXlsxFilePath(folderName,fileName)
                    else:
                        print('処理失敗[xlsxファイル名中古新車なし]：' + fileName) 
                        continue 
                    
                    createCsv(xlsxPath, getTypeToFolder(fileType), companyId)
                    imageCopy(folderName, companyId, fileType)
               
        except:
            print('処理失敗[設定・その他エラー]：' + folderName) 
                    
    