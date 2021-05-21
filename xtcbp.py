import csv
import pandas as pd
import os
import glob
import datetime

#대상파일 위치
XLSX_PATH = './xlsxFiles'

#변환 파일 관련
CSV_PATH = './' + datetime.date.today().strftime('%y%m%d') + '/Original/csv'
IMAGE_PATH = './' + datetime.date.today().strftime('%y%m%d') + '/Original/images'

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
    return [path for path in glob.glob(path + extension) if os.path.isfile(path)] 

# 시트 리스트 취득
def getSheetList(folderName: str, fileName: str) -> list:
    folderPath = f'{XLSX_PATH}/{folderName}'
    xlsxPath = f'{folderPath}/{fileName}'
    
    wb = pd.read_excel(xlsxPath, sheet_name = None, engine='openpyxl')
    createCsv(CSV_PATH,folderName,wb)    
    return [*wb]

# 회사ID 정보 취득 
def getCompanyId(folderName: str, fileName: str, fileType: int = 0) -> str:
    setData = getSetData()    
    if fileType != 0 :
        findData = fileName.replace('.xlsx')   
    else:
        result = setData[folderName][0]
     
    return result

# CSV파일생성
def createCsv(path: str, fileName: str, data: list) -> None:
    def createDir():
        dir = f'{CSV_PATH}/test/123'
        os.makedirs(dir, exist_ok=True) #옵션으로 이미 존재할때는 넘어감
    createDir()
    # data['Sheet1'].to_csv(f'{path}/{fileName}.csv', index=False, encoding='utf-8-sig')    



# 메인로직 
if __name__ == "__main__":
    
    for folderName in getFolderList():
        fileList = getXlsxFileList(folderName)
        try:
            if len(fileList) == 0:
                print('処理失敗[xlsxファイルなし]：' + folderName + '/' + fileName) 
                continue
            elif len(fileList) == 1:
                fileName = fileList[0]
                companyId = getCompanyId(folderName, fileName)
            else:
                for fileName in fileList:
                    fileType = fileName
                    companyId = getCompanyId(folderName, fileName)
            
            
               
        except:
            print('処理失敗[設定・その他エラー]：' + folderName + '/' + fileName) 
            
        
        
            
        
        
        # sheetList = getSheetList(folderName,fileList[0])        
        
    