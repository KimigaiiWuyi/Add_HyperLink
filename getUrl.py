import openpyxl,json,sys,requests,yaml,os

FILE_PATH = os.path.dirname(__file__)

with open('Config.yaml', encoding='UTF-8') as yaml_file:
    token = yaml.safe_load(yaml_file)

ck = token["Cookies"]

if len(sys.argv) > 1:
    url = sys.argv[1]
else:
    url = input("input your excel name: ")

wb=openpyxl.load_workbook(os.path.join(FILE_PATH,url))
sheet=wb.worksheets[0]
M_url = "https://quantum.63yx.com/index.php?c=api-AdsysMaterial&a=transfer&sec="

def GetURL(name):
    req = requests.get(
            url="https://quantum.37wan.com/index.php?c=adsys-AdsysMaterial&a=list&search_name=" + str(name),
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36',
                "Cookie": ck})
    data = json.loads(req.text)
    return data
        
for row in range(2,sheet.max_row+1):  
    for column in "B":  #Here you can add or reduce the columns
        print(row)
        cell_name = "{}{}".format(column, row)
        raw_data = GetURL(sheet[cell_name].value)
        for i in raw_data["list"]["data"]:
            if i["NAME"] != sheet[cell_name].value:
                raw_data["list"]["data"].remove(i)
        try:
            if raw_data["list"]["data"][0]["MAX_SOURCE"]["URL"].find("http://") != -1:
                #sheet.cell(row,2).value='=HYPERLINK("%s","%s")' % (raw_data["list"]["data"][0]["MAX_SOURCE"]["URL"], sheet[cell_name].value)
                sheet.cell(row,2).hyperlink = raw_data["list"]["data"][0]["MAX_SOURCE"]["URL"]
                sheet.cell(row,2).value = sheet[cell_name].value
            else:
                sheet.cell(row,2).hyperlink = M_url + raw_data["list"]["data"][0]["MAX_SOURCE"]["ID"]
                sheet.cell(row,2).value = sheet[cell_name].value
            sheet.cell(row,2).style = "Hyperlink"
        except:
            print("error")
        #if row >= 100:
        #    break
    else:
        continue
    break
wb.save('文件名称.xlsx')
