from selenium import webdriver
import time
import xlwt

driver = webdriver.Chrome()

def pagescroll():
    jslist=['window.scrollTo(0,1200);','window.scrollTo(0,2400);','window.scrollTo(0,3600);''window.scrollTo(0,0);']
    for js in jslist:
        driver.execute_script(js)
        time.sleep(0.5)

def get_data():
    # # driver = webdriver.Chrome()
    # driver.get(url)
    # pagescroll()
    # exhlist = driver.find_elements_by_class_name(cn)
    # for exh in exhlist:
    #     print(exh.text)
    # # driver.quit()

    exhurl='https://www.nationalmuseum.sg/our-exhibitions/exhibition-list'
    prgurl='https://www.nationalmuseum.sg/our-programmes/programmes-list'
    radurl='https://www.nationalmuseum.sg/retail-and-dinning/retail-and-dining-list'
    spturl='https://www.nationalmuseum.sg/support-us/support-us-list'
    list_url=[exhurl, prgurl, radurl]
    list_button=['exhibitions', 'programmes','retail_and_dining','support']
    exhibitions=[]
    programmes=[]
    retail_and_dining=[]
    support=[]
    dict={"exhibitions":exhibitions, "programmes":programmes, "retail_and_dining":retail_and_dining, "support":support}

    for i in range(0,2):
        driver.get(list_url[i])
        pagescroll()

        element_name=driver.find_elements_by_class_name('grid_item--title')
        element_mix=driver.find_elements_by_class_name('grid_item--location')
        element_time=driver.find_elements_by_class_name('grid_item--date')
        element_content=driver.find_elements_by_class_name('grid_item--content')
        element_image=driver.find_elements_by_class_name('carousel--img')
        element_url=driver.find_elements_by_css_selector("[class='listing--grid_item scroll-watch-in-view scroll-watch-ignore']")

        mixlen = len(element_mix)
        number = len(element_content)

        name_list=[]
        location_list=[]
        price_list=[]
        time_list=[]
        content_list=[]
        image_list=[]
        url_list = ['https://www.nationalmuseum.sg/our-exhibitions/exhibition-list/art-of-the-rehearsal?sc_lang=en']

        for image in element_image:
            print(image.get_attribute('src'))
            image_list.append(image.get_attribute('src'))
        for name in element_name:
            print(name.text)
            name_list.append(name.text)
        for k in range(0,mixlen):
            print(element_mix[k].text)
            if k%2==0:
                location_list.append(element_mix[k].text)
            else:
                price_list.append(element_mix[k].text)
        for tm in element_time:
            print(tm.text)
            time_list.append(tm.text)
        for content in element_content:
            print(content.text)
            content_list.append(content.text)
        for url in element_url:
            print(url.get_attribute('href'))
            url_list.append(url.get_attribute('href'))

        for j in range(0,number):
            dict[list_button[i]].append(name_list[j])

            dict[list_button[i]].append(location_list[j])
            dict[list_button[i]].append(price_list[j])

            dict[list_button[i]].append(time_list[j])
            dict[list_button[i]].append(content_list[j])
            dict[list_button[i]].append(image_list[j])
            dict[list_button[i]].append(url_list[j])

    driver.get(radurl)
    pagescroll()
    rad_name=driver.find_elements_by_class_name('grid_item--title')
    rad_time = driver.find_elements_by_class_name('grid_item--date')
    rad_content = driver.find_elements_by_class_name('grid_item--content')
    rad_image=driver.find_elements_by_class_name('carousel--img')
    rad_url=driver.find_elements_by_css_selector("[class='listing--grid_item scroll-watch-in-view scroll-watch-ignore']")
    rad_name_list=[]
    rad_time_list = []
    rad_content_list=[]
    rad_image_list=[]
    rad_url_list=[]
    for name in rad_name:
        print(name.text)
        rad_name_list.append(name.text)
    for tm in rad_time:
        print(tm.text)
        rad_time_list.append(tm.text)
    for content in rad_content:
        print(content.text)
        rad_content_list.append(content.text)
    for image in rad_image:
        print(image.get_attribute('src'))
        rad_image_list.append(image.get_attribute('src'))
    for url in rad_url:
        print(url.get_attribute('href'))
        rad_url_list.append(url.get_attribute('href'))
    number = len(rad_name)
    for j in range(0,number):
        dict['retail_and_dining'].append(rad_name_list[j])
        dict['retail_and_dining'].append(0)
        dict['retail_and_dining'].append(0)
        dict['retail_and_dining'].append(rad_time_list[j])
        dict['retail_and_dining'].append(rad_content_list[j])
        dict['retail_and_dining'].append(rad_image_list[j])
        dict['retail_and_dining'].append(rad_url_list[j])

    driver.get(spturl)
    pagescroll()
    spt_name=driver.find_elements_by_class_name('grid_item--title')
    spt_content = driver.find_elements_by_class_name('grid_item--content')
    spt_image=driver.find_elements_by_class_name('carousel--img')
    spt_url=driver.find_elements_by_css_selector("[class='listing--grid_item scroll-watch-in-view scroll-watch-ignore']")
    spt_name_list=[]
    spt_content_list=[]
    spt_image_list=[]
    spt_url_list=[]
    for name in spt_name:
        print(name.text)
        spt_name_list.append(name.text)
    for content in spt_content:
        print(content.text)
        spt_content_list.append(content.text)
    for image in spt_image:
        print(image.get_attribute('src'))
        spt_image_list.append(image.get_attribute('src'))
    for url in spt_url:
        print(url.get_attribute('href'))
        spt_url_list.append(url.get_attribute('href'))
    number = len(spt_name)
    for j in range(0,number):
        dict['support'].append(spt_name_list[j])
        dict['support'].append(0)
        dict['support'].append(0)
        dict['support'].append(0)
        dict['support'].append(spt_content_list[j])
        dict['support'].append(spt_image_list[j])
        dict['support'].append(spt_url_list[j])


    print(dict)
    return dict

def save_data(dict):
    wb=xlwt.Workbook(encoding='utf-8')
    for d in dict:
        sheet=wb.add_sheet(d,cell_overwrite_ok=True)
        headlist=[d,'location', 'price', 'time', 'content','image','url']
        row=0
        col=0
        for head in headlist:
            sheet.write(col,row,head)
            row+=1
        i=0
        for data in dict[d]:
            if(i%7==0):
                col+=1
            sheet.write(col,i%7,data)
            i+=1
    wb.save('museum_data_raw.xls')

def run():
    dict=get_data()
    save_data(dict)


run()
driver.quit()