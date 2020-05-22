from selenium import webdriver
from xlutils.copy import copy
import xlrd
from xlrd import *

Work_Order_Path = ("/Users/arafat.safdar/Desktop/Airtable_workorder.xlsx")
Save_Path = ('/Users/arafat.safdar/Desktop/AirTable_0123.xls')
Webdriver_Path = ('/Users/arafat.safdar/Downloads/chromedriver')

r = 0
w = copy(open_workbook(Work_Order_Path))
wb = xlrd.open_workbook(Work_Order_Path)
sheet = wb.sheet_by_index(0)
len_count = len(sheet.col_values(0))
print(len_count)

while r < len_count:
    wb = xlrd.open_workbook(Work_Order_Path)
    sheet = wb.sheet_by_index(0)
    checkcellvalue = sheet.cell_value(r, 0)
    print('xxxx____________xxxx____________xxxx____________xxxx____________xxxx____________xxxx')

    if checkcellvalue == '':
        r = r + 1
        print('no work order number')
        print(r)
        continue

    order = str(int(sheet.cell_value(r, 0)))

    driver = webdriver.Chrome(Webdriver_Path)
    driver.implicitly_wait(10)

    link = 'http://10.23.112.40/channels?name='

    url = link + order

    print(url)
    driver.get(url)
    driver.set_window_size(200,200)

    xyz = driver.find_element_by_xpath(
        '//*[@ng-repeat="channel in channels track by trackingKeys(channel)"]/tr[2]/td/div/div/div[2]/div/div/lanel/input').get_attribute(
        'value')
    abc = driver.find_element_by_xpath(
        '//*[@ng-repeat="channel in channels track by trackingKeys(channel)"]/tr[2]/td/div/div/div[4]/div/div/lanel/input').get_attribute(
        'value')
    if abc == 'eth2':
        print(xyz)
        print(abc)
        w.get_sheet(0).write(r, 2, xyz)
        if xyz[48] == '/':
            w.get_sheet(0).write(r, 7, xyz[42:47])
            w.get_sheet(0).write(r, 8, 'https://kayoeventshlsts.akamaized.net/hls/live/' + xyz[42:] + '.m3u8')

        else:
            w.get_sheet(0).write(r, 7, abc[43:50])
            w.get_sheet(0).write(r, 8, 'https://kayoeventshlsts.akamaized.net/hls/live/' + xyz[43:] + '.m3u8')
    else:
        print(xyz)
        print(abc)
        w.get_sheet(0).write(r, 1, xyz)
        w.get_sheet(0).write(r, 2, abc)
        if xyz[55] == '/':
            w.get_sheet(0).write(r, 5, xyz[48:55])
            w.get_sheet(0).write(r, 6, 'https://kayoeventshls.akamaized.net/cmaf/live/' + xyz[48:]+ '.m3u8' )

        else:
            w.get_sheet(0).write(r, 5, xyz[47:53])
            w.get_sheet(0).write(r, 6, 'https://kayoeventshls.akamaized.net/cmaf/live/' + xyz[47:] + '.m3u8')

        if abc[48] == '/':
            w.get_sheet(0).write(r, 7, abc[42:47])
            w.get_sheet(0).write(r, 8, 'https://kayoeventshlsts.akamaized.net/hls/live/' + abc[42:] + '.m3u8')

        else:
            w.get_sheet(0).write(r, 7, abc[43:50])
            w.get_sheet(0).write(r, 8, 'https://kayoeventshlsts.akamaized.net/hls/live/' + abc[43:] + '.m3u8')



    r = r + 1
    print(order)
    print(r)
    print('xxxx____________xxxx____________xxxx____________xxxx____________xxxx____________xxxx')
    driver.close()

w.save(Save_Path)
driver.quit()
print("Complete")

