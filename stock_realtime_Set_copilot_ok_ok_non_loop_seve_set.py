import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time
import xlwings as xw
from pathlib import Path
from selenium.webdriver.chrome.service import Service

# ตั้งค่า driver สำหรับ Chrome
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # ใช้งานในโหมด headless
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# ไปยังหน้าหลักของ SET
driver.get('https://www.set.or.th/en/home')

# ดึงข้อมูลหน้าเว็บ
data = driver.page_source

# อ่านข้อมูลตารางจากหน้าเว็บ
data_df = pd.read_html(data)[0]

# ทำความสะอาดชื่อคอลัมน์
data_df.columns = [c.replace(' (Click to sort Ascending)', '') for c in data_df.columns]

# ลบแถวแรกที่ไม่จำเป็น
data_df.drop([0], inplace=True)

# ตั้งค่า index
data_df.set_index('Index', inplace=True)

# ตรวจสอบว่าไดเรกทอรีที่ต้องการบันทึกไฟล์มีอยู่จริง
output_dir = Path(r'C:/Users/plaifa/Downloads/Python/Data')
output_dir.mkdir(parents=True, exist_ok=True)

# ฟังก์ชันสำหรับดึงข้อมูลหุ้น
def get_stock_data(stock, driver):
    stock = stock.split()[0]
    url = f'https://www.set.or.th/en/market/index/{stock}/overview'
    driver.get(url)
    stock_data = driver.page_source
    a_df = pd.read_html(stock_data)[1]
    a_df.columns = [c.replace(' (Click to sort Ascending)', '') for c in a_df.columns]
    return a_df

# ดึงข้อมูลหุ้นทั้งหมดเพียงครั้งเดียว
all_stock_dict = {}
for stock in data_df.index:
    all_stock_dict[stock] = get_stock_data(stock, driver)
    all_stock_dict[stock].to_csv(output_dir / f'{stock}.csv', index=False)
    print(all_stock_dict[stock])

# ปิด driver เมื่อทำงานเสร็จ
driver.quit()

#### Save as Set.csv  #############
# กำหนด path ของไฟล์ CSV
path = 'C:/Users/plaifa/Downloads/Python/Data/'

# อ่านข้อมูลจากแต่ละไฟล์ CSV
set100_df = pd.read_csv(path + 'SET100.csv')
sset_df = pd.read_csv(path + 'sSET.csv')
sethd_df = pd.read_csv(path + 'SETHD.csv')
setclmv_df = pd.read_csv(path + 'SETCLMV.csv')
setwb_df = pd.read_csv(path + 'SETWB.csv')
setesg_df = pd.read_csv(path + 'SETESG.csv')
set10ff_df = pd.read_csv(path + 'SET100FF.csv')
set50ff_df = pd.read_csv(path + 'SET50FF.csv')
set50_df = pd.read_csv(path + 'SET50.csv')


# รวมข้อมูลทั้งสาม DataFrame
combined_df = pd.concat([set100_df, sset_df, sethd_df])

# บันทึก DataFrame ที่รวมแล้วลงในไฟล์ใหม่ชื่อ set.csv ที่ path เดิม
combined_df.to_csv(path + 'set.csv', index=False)

print("ไฟล์รวม set.csv ถูกสร้างและเก็บไว้ที่ C:/Users/Thanimwas/Downloads/Python/Data/ เรียบร้อยแล้ว")
