import os
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome  import ChromeDriverManager



ROOT_DIR = os.path.dirname(__file__)
data_path = os.path.join(ROOT_DIR, "data")
os.makedirs(data_path, exist_ok=True)
#driver = webdriver.Chrome('./chromedriver')
driver=webdriver.Chrome(service=Service(ChromeDriverManager().install()))

def get_element(element_list):
    element = None
    usage=""
    for elements in element_list:
        if element !=None:
            break
        print(f"Trying to set: {elements} Using:")
        try:
            usage+="CSS SELECTOR, "
            element = driver.find_element(By.CSS_SELECTOR,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="XPATH, "
            element = driver.find_element(By.XPATH,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="ID, " 
            element = driver.find_element(By.ID,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="NAME, "
            element = driver.find_element(By.NAME,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="LINK, "
            element = driver.find_element(By.LINK_TEXT,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="PARTIAL LINK, "
            element = driver.find_element(By.PARTIAL_LINK_TEXT,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="TAG, "
            element = driver.find_element(By.TAG_NAME,f"{elements}")
            break
        except:
            element=None
        try:
            usage+="CLASS.\n"
            element = driver.find_element(By.CLASS_NAME,f"{elements}")
            break
        except:
            element=None
    print(f"{usage}")
    if element == None:
        print(f"Cant find any element on {element_list}")
    else:
        print(f"Selected:: {element}")
    return element



def get_instructions(user,password,period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp):
    return f'''main\n
goto|https://efow-dev1.fa.us2.oraclecloud.com\n
wait|efow-dev1\n
send|{user}??[aria-label='User ID'],,#userid\n
send|{password}??[aria-label='Password'],,#password\n
click|[aria-label='Sign In'],,#btnActive\n
wait|fscmUI/faces\n
goto|https://efow-dev1.fa.us2.oraclecloud.com\n
sleep|3\n
click|[aria-label='Home'],,#pt1\\:_UIShome\\:\\:icon > g:nth-child(5) > path\n
click|#itemNode_financial_reporting_financial_reporting > svg > path.svg-outline,,aria-label='Financial Reporting Center'\n
sleep|3\n
send|R-002??[aria-label='Enter search terms'],,#pt1\\:_FOr1\\:1\\:_FONSr2\\:0\\:_FOTsr1\\:0\\:SP1\\:s12\\:it3\\:\\:content\n
click|#pt1\\:_FOr1\\:1\\:_FONSr2\\:0\\:_FOTsr1\\:0\\:SP1\\:s12\\:cil1\\:\\:icon,,[aria-label='img[alt=\"Search\"]'],,img[alt='Search'],,Search\n
sleep|4\n
click|#pt1\\:_FOr1\\:1\\:_FONSr2\\:0\\:_FOTsr1\\:0\\:SP1\\:t1\\:0\\:gil1,,//*[@id="pt1:_FOr1:1:_FONSr2:0:_FOTsr1:0:SP1:t1:0:gil1"],,#pt1:_FOr1:1:_FONSr2:0:_FOTsr1:0:SP1:t1:0:gil1\n
sleep|15\n
window|1\n
iframe|DialogContentBodyContent\n
send|,{period}??aria-label=AccountingPeriod[role="textbox"],,#textarea0,,fld0,,bodyFonts,,virtual\n
parent|1\n
click|#DialogContentFooterContent > tbody > tr:nth-child(1) > td:nth-child(5) > button > u,,[aria-label='OK[role=\"button\"]'],,role='button',,button,,OK,,primaryActive\n
default_content|1\n
main\n
sleep|5\n
space|3\n
get_list|iframe\n
switch_to_element_number|iframe??1\n
get_list|frame\n
switch_to_element_number|frame??0\n
get_list|frame\n
switch_to_element_number|frame??0\n
click|'//select/option=[@id=previewpov]',,//select/option=[@id=previewpov],,previewpov,,#previopov,,//select/option=[@id=#previewpov],,'//select/option=[#@id=previewpov]'\n
window|1\n
iframe|DialogContentBodyContent\n
remove_text|[id=\"dimText0\\+0\"],,[id=\'dimText0\\+0\'],,[id=dimText0 +0],,[id=dimText0\\+0],,[@id=\"dimText0\\+0\"],,dimText0+0,,#dimText0+0,,[@id=dimText0+0]\n
send|{ledger}??[id=\"dimText0\\+0\"],,[id=\'dimText0\\+0\'],,[id=dimText0 +0],,[id=dimText0\\+0],,[@id=\"dimText0\\+0\"],,dimText0+0,,#dimText0+0,,[@id=dimText0+0]\n
remove_text|[id=\"dimText0\\+1\"],,[id=\'dimText0\\+1\'],,[id=dimText0 +1],,[id=dimText0\\+1],,[@id=\"dimText0\\+1\"],,dimText0+1,,#dimText0+1,,[@id=dimText0+1]\n
send|{entity}??[id=\"dimText0\\+1\"],,[id=\'dimText0\\+1\'],,[id=dimText0 +1],,[id=dimText0\\+1],,[@id=\"dimText0\\+1\"],,dimText0+1,,#dimText0+1,,[@id=dimText0+1]\n
remove_text|[id=\"dimText0\\+2\"],,[id=\'dimText0\\+2\'],,[id=dimText0 +2],,[id=dimText0\\+2],,[@id=\"dimText0\\+2\"],,dimText0+2,,#dimText0+2,,[@id=dimText0+2]\n
send|{dpt}??[id=\"dimText0\\+2\"],,[id=\'dimText0\\+2\'],,[id=dimText0 +2],,[id=dimText0\\+2],,[@id=\"dimText0\\+2\"],,dimText0+2,,#dimText0+2,,[@id=dimText0+2]\n
remove_text|[id=\"dimText0\\+3\"],,[id=\'dimText0\\+3\'],,[id=dimText0 +3],,[id=dimText0\\+3],,[@id=\"dimText0\\+3\"],,dimText0+3,,#dimText0+3,,[@id=dimText0+3]\n
send|{product}??[id=\"dimText0\\+3\"],,[id=\'dimText0\\+3\'],,[id=dimText0 +3],,[id=dimText0\\+3],,[@id=\"dimText0\\+3\"],,dimText0+3,,#dimText0+3,,[@id=dimText0+3]\n
remove_text|[id=\"dimText0\\+4\"],,[id=\'dimText0\\+4\'],,[id=dimText0 +4],,[id=dimText0\\+4],,[@id=\"dimText0\\+4\"],,dimText0+4,,#dimText0+4,,[@id=dimText0+4]\n
send|{project}??[id=\"dimText0\\+4\"],,[id=\'dimText0\\+4\'],,[id=dimText0 +4],,[id=dimText0\\+4],,[@id=\"dimText0\\+4\"],,dimText0+4,,#dimText0+4,,[@id=dimText0+4]\n
remove_text|[id=\"dimText0\\+5\"],,[id=\'dimText0\\+5\'],,[id=dimText0 +5],,[id=dimText0\\+5],,[@id=\"dimText0\\+5\"],,dimText0+5,,#dimText0+5,,[@id=dimText0+5]\n
send|{ico}??[id=\"dimText0\\+5\"],,[id=\'dimText0\\+5\'],,[id=dimText0 +5],,[id=dimText0\\+5],,[@id=\"dimText0\\+5\"],,dimText0+5,,#dimText0+5,,[@id=dimText0+5]\n
remove_text|[id=\"dimText0\\+6\"],,[id=\'dimText0\\+6\'],,[id=dimText0 +6],,[id=dimText0\\+6],,[@id=\"dimText0\\+6\"],,dimText0+6,,#dimText0+6,,[@id=dimText0+6]\n
send|{future}??[id=\"dimText0\\+6\"],,[id=\'dimText0\\+6\'],,[id=dimText0 +6],,[id=dimText0\\+6],,[@id=\"dimText0\\+6\"],,dimText0+6,,#dimText0+6,,[@id=dimText0+6]\n
remove_text|[id=\"dimText0\\+7\"],,[id=\'dimText0\\+7\'],,[id=dimText0 +7],,[id=dimText0\\+7],,[@id=\"dimText0\\+7\"],,dimText0+7,,#dimText0+7,,[@id=dimText0+7]\n
send|{currency_type}??[id=\"dimText0\\+7\"],,[id=\'dimText0\\+7\'],,[id=dimText0 +7],,[id=dimText0\\+7],,[@id=\"dimText0\\+7\"],,dimText0+7,,#dimText0+7,,[@id=dimText0+7]\n
remove_text|[id=\"dimText0\\+8\"],,[id=\'dimText0\\+8\'],,[id=dimText0 +8],,[id=dimText0\\+8],,[@id=\"dimText0\\+8\"],,dimText0+8,,#dimText0+8,,[@id=dimText0+8]\n
send|{currency}??[id=\"dimText0\\+8\"],,[id=\'dimText0\\+8\'],,[id=dimText0 +8],,[id=dimText0\\+8],,[@id=\"dimText0\\+8\"],,dimText0+8,,#dimText0+8,,[@id=dimText0+8]\n
remove_text|[id=\"dimText0\\+9\"],,[id=\'dimText0\\+9\'],,[id=dimText0 +9],,[id=dimText0\\+9],,[@id=\"dimText0\\+9\"],,dimText0+9,,#dimText0+9,,[@id=dimText0+9]\n
send|{tp}??[id=\"dimText0\\+9\"],,[id=\'dimText0\\+9\'],,[id=dimText0 +9],,[id=dimText0\\+9],,[@id=\"dimText0\\+9\"],,dimText0+9,,#dimText0+9,,[@id=dimText0+9]\n
parent|1\n
click|[arialabel='OK[role=\"button\"]'],,#DialogContentFooterContent > tbody > tr:nth-child(1) > td:nth-child(5) > button\n
default_content|1\n
main\n
get_list|iframe\n
switch_to_element_number|iframe??1\n
get_list|frame\n
switch_to_element_number|frame??0\n
get_list|frame\n
switch_to_element_number|frame??0\n
click|'//select/option=[@id=excel]',,//select/option=[@id=excel],,excel,,#excel,,//select/option=[@id=#excel],'//select/option=[#@id=excel]',\n
window|1\n
close|0\n
#default_content|1\n
#parent|1\n
main\n
goto|https://efow-dev1.fa.us2.oraclecloud.com\n
sleep|3\n
click|"#pt1\\:_UIScmil2u",,"pt1\\:_UIScmil2u",,#pt1\\:_UIScmil2u,,pt1\\:_UIScmil2u,,pt1:_UIScmil2u,,#pt1:_UIScmil2u\n
click|"#pt1\\:_UISlg1",,"pt1\\:_UISlg1",,#pt1\\:_UISlg1,,pt1\\:_UISlg1,,#pt1:_UISlg1,,pt1:_UISlg1\n
click|"#Confirm",,"COnfirm",,#Confirm,,Confirm\n
goto|https://efow-dev1.fa.us2.oraclecloud.com\n
sleep|3\n
close|0\n
'''



def get_list_of_element(element_list):
    element = []
    print(f"Trying to get element list: {element_list} using:")
    usage=""
    for elements in element_list:
        if element != []:
            break
        usage+="TAG, "
        try:
            element = driver.find_elements(By.TAG_NAME,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="CLASS, "
        try:
            element = driver.find_elements(By.CLASS_NAME,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="NAME, "
        try:
            element = driver.find_elements(By.NAME,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="PARTIAL LINK, "
        try:
            element = driver.find_elements(By.PARTIAL_LINK_TEXT,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="LINK, "
        try:
            element = driver.find_elements(By.LINK_TEXT,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="ID, "
        try:
            element = driver.find_elements(By.ID,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="SELECTOR, "
        try:
            element = driver.find_elements(By.CSS_SELECTOR,f"{elements}")
            if element != []:
                break
        except:
            element=[]
        usage+="XPATH, \n"
        try:
            element = driver.find_elements(By.XPATH,f"{elements}")
            if element != []:
                break
        except:
            element=[]
    print(f"{usage}")    
    if element ==[]:
        element=None
    if element == None:
        print(f"Cant find any element on {element_list}")
    else:
        counter = -1
        for webelement in element:
            counter +=1
            print(f"{counter} :: {webelement.get_attribute('name')}")
            print(f"{counter} :: {webelement.get_attribute('id')}")
        #print(f"Selected:: {element}")
    return element

    
def wait_for(url):
    while True:
       if url in driver.current_url:
            sleep(1)
            break

def execute(user,password,period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp):
    main_page = driver.current_window_handle
    instructions=get_instructions(user,password,period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp).splitlines()
    element=None
    for instruction in instructions:
        print(f"Instruction :: {instruction}")
        if instruction.strip() =="":
            print("\n\nJust separator \n\n")
        elif instruction[0] == "#":
            print("Just a comment")
        else:
            try:
                action, element_list = instruction.split("|")
            except:
                #driver.switch_to.default_content()
                driver.switch_to.window(main_page)
            else:
                print(f"Action :: {action}")
                print(f"Element :: {element_list}")
                if action.upper() == "CLICK":
                    element_list = element_list.split(",,")
                    element= get_element(element_list)
                    if element != None:
                        element.click() 
                        sleep(5)
                    else:
                        print("Null element on click")
                        break
                if action.upper()=="GOTO":
                     driver.get(f"{element_list}")
                if action.upper() == "SEND":
                    words , element_list = element_list.split("??")
                    element_list = element_list.split(",,")
                    element= get_element(element_list)
                    if element != None:
                        element.click()
                        element.send_keys(f"{words}")
                        sleep(2)
                    else:
                        print("Null element on send")
                        break
                if action.upper()=="WAIT":
                    wait_for(f"{element_list}")
                if action.upper() == "FRAME":
                    element_list = element_list.split(",,")
                    element= get_element(element_list)
                    print(f"Element:: {element}")
                    for element in element_list:
                        try:
                            print(f"Trying to change to frame to {element.get_attribute('name')}")
                            driver.switch_to.frame(f"{element}")
                        except:
                            element=None
                    if element == None:
                        print(f"Cant find any frame with name in the namelist {element_list}")
                        break
                if action.upper() =="SLEEP":
                    sleep(int(f"{element_list}"))
                if action.upper() =="WINDOW":
                    for window_handle in driver.window_handles:
                        if window_handle != main_page:
                            print(f"Changing to {window_handle} from main window {main_page}")
                            driver.switch_to.window(window_handle)
                if action.upper() == "IFRAME":
                    element_list=["iframe"]
                    element = get_element(element_list)
                    print(element.get_attribute("name"))
                    driver.switch_to.frame(element)
                if action.upper() == "CLOSE":
                    driver.close()
                if action.upper() == "PARENT":
                    driver.switch_to.parent_frame()
                if action.upper() == "DEFAUL_CONTENT":
                    driver.switch_to.default_content()
                if action.upper() == "SEND_TO":
                    word,number = element_list.split("??")
                    selected = element[int(number)-1]
                    selected.send(word)
                if action.upper() == "CLICK_TO":
                    selected = element[int(element_list)-1]
                    selected.click()
                if action.upper() =="GET_LIST":
                    element = get_list_of_element(element_list)
                if action.upper() == "SWITCH_TO_ELEMENT_NUMBER":
                    frametag , number = element_list.split("??")
                    frame=driver.find_elements(By.TAG_NAME,frametag)[int(number)]
                    driver.switch_to.frame(frame)
                if action.upper() == "SPACE":
                    lines= "\n" * int(element_list)
                    print(lines)
                if action.upper() =="REMOVE_TEXT":
                    element_list = element_list.split(",,")
                    element = get_element(element_list)
                    element.send_keys(Keys.CONTROL,'a')
                    element.send_keys(Keys.BACKSPACE)
               
                                        
        
def Do_the_job(excel_file):
    frs_df = pd.read_excel(f"{ROOT_DIR}\\{excel_file}", skiprows=4)
    columns = pd.read_excel(f"{ROOT_DIR}\\{excel_file}", skiprows=4).columns
    for index, row in frs_df.iterrows():
        print(row)
        period=row['Accounting Period']
        ledger=row['Ledger']
        entity=row['Entity']
        dpt=row['Dept']
        product=row['Product']
        project=row['Project']
        ico=row['ICO']
        future=row['GWFuture']
        currency_type=row['Scenario']
        currency=row['Currency']
        tp=row['Currency Type']
        #print(period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp)
        #print(get_instructions("Arturo","Roa",period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp).splitlines())
        user="vsethi"
        password="Welcome1"
        execute(user,password,period,ledger,entity,dpt,product,project,ico,future,currency_type,currency,tp):
        break
    
            
        
