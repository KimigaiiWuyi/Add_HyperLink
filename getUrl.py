import openpyxl,json,sys,requests,yaml,os

FILE_PATH = os.path.dirname(__file__)

with open('Config.yaml', encoding='UTF-8') as yaml_file:
    token = yaml.safe_load(yaml_file)

ck = token["Cookies"]

if len(sys.argv) == 3:
    url = sys.argv[1]
    column_input = sys.argv[2]
else:
    url = input("input your excel name: ")
    column_input = input("输入你数据所在的列名（例如：A、B）: ")

column_input = column_input.upper()
wb=openpyxl.load_workbook(os.path.join(FILE_PATH,url))
sheet=wb.active
M_url = "https://quantum.63yx.com/index.php?c=api-AdsysMaterial&a=transfer&sec="

def GetURL(name):
    req = requests.get(
            url="https://quantum.37wan.com/index.php?c=adsys-AdsysMaterial&a=list&search_name=" + str(name),
            headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36',
                "Cookie": ck})
    data = json.loads(req.text)
    return data
        
for cell in sheet[column_input]:
    if cell.value:
        print("正在为【" + cell.value + "】添加超链接……")
        raw_data = GetURL(cell.value)
        for i in raw_data["list"]["data"]:
            if i["NAME"] != cell.value:
                raw_data["list"]["data"].remove(i)
        try:
            if raw_data["list"]["data"][0]["MAX_SOURCE"]["URL"].find("http://") != -1:
                cell.hyperlink = raw_data["list"]["data"][0]["MAX_SOURCE"]["URL"]
                cell.value = cell.value
            else:
                cell.hyperlink = M_url + raw_data["list"]["data"][0]["MAX_SOURCE"]["ID"]
                cell.value = cell.value
            cell.style = "Hyperlink"
        except:
            print("未找到相应数据。")
    else:
        print("该行数据为空。")

wb.save('{}_带链接.xlsx'.format(url[:-5]))
