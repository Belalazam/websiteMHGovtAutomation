import sys
import time
import traceback
import openpyxl
import threading
from tkinter import *
from tkinter import filedialog
from selenium import webdriver
from datetime import datetime
from cryptography.fernet import Fernet
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


#globla variables
listOfQuery = []
ii = -1

#loading the excel file
def getDataFromExcel(specificFilePath):
    wb = openpyxl.load_workbook(specificFilePath)
    sheet = wb.active
    data = []
    memory = {}
    with open("./memo/failure.txt",'r') as f:
        for line in f:
            line = line.strip()
            memory[str(line)] = True   
    with open("./memo/memory.txt",'r') as f:
        for line in f:
            line = line.strip()
            memory[str(line)] = True
    for i in range(3,sheet.max_row+1):
        temp = []
        tempz = str(sheet.cell(row=i,column=10).value)
        if(memory.get(tempz) != None):
            continue
        temp.append(str(sheet.cell(row=i,column=1).value))
        temp.append(str(sheet.cell(row=i,column=2).value))
        temp.append(str(sheet.cell(row=i,column=3).value))
        temp.append(str(sheet.cell(row=i,column=4).value))
        temp.append(str(sheet.cell(row=i,column=5).value))
        temp.append(str(sheet.cell(row=i,column=6).value))
        temp.append(str(sheet.cell(row=i,column=7).value))
        temp.append(str(sheet.cell(row=i,column=8).value))
        temp.append(str(sheet.cell(row=i,column=9).value))
        temp.append(str(sheet.cell(row=i,column=10).value))
        temp.append(str(sheet.cell(row=i,column=11).value))
        temp.append(str(sheet.cell(row=i,column=12).value))
        temp.append(str(sheet.cell(row=i,column=13).value))
        temp.append(str(sheet.cell(row=i,column=14).value))
        temp.append(str(sheet.cell(row=i,column=15).value))
        temp.append(str(sheet.cell(row=i,column=16).value))
        temp.append(str(sheet.cell(row=i,column=17).value))
        data.append(temp)
    return data

#filling the output areas
######################################################
def fillOutputArea(outputBox,query,mode):
    if(mode == 0):
        outputBox.insert(END,query+'\n')
    else :
        outputBox.insert(END,query)
######################################################


#file Loading from desktop
def loadExcelFile():
     global filePath
     latestTime.set("Wait Loading....")
     load_button.config(state=DISABLED,bg='red')
     filePath = filedialog.askopenfilename()
     global listOfQuery
     listOfQuery = getDataFromExcel(filePath)
     var = "Last Updated: " + str(time.strftime("%H:%M:%S", time.localtime()))
     latestTime.set(var)
     load_button.config(state=ACTIVE,bg=orig_color)

####################################################################

def getOnlyValue(specificFilePath):
    with open(specificFilePath,'r') as f_in:
        z = f_in.readline()
        return str(z).strip()
#########################################################################################################################################
def getElement(driver,byWhat,strng,waitTime):
    wait = WebDriverWait(driver, waitTime)
    return wait.until(EC.presence_of_element_located((byWhat,strng)))
#####################################################################################################
def getElements(driver, byWhat, strng, waitTime):
    wait = WebDriverWait(driver, waitTime)
    return wait.until(EC.presence_of_all_elements_located((byWhat, strng)))
#################################################################################################################
def operateElement(value,element,keysValue,waitTime):
    start_time = time.time()
    if(value == "click"):
        while time.time() - start_time < waitTime:
            try:
                element.click()
                break  
            except:
                time.sleep(0.5)
    elif(value == "send_keys"):
        while time.time() - start_time < waitTime:
            try:
                element.send_keys(keysValue)
                break
            except:
                time.sleep(0.5)

def decrypt_string(encrypted_string):
    key = 'ECPHuqGMo6QE2tcLElUX2GBmvOngpzFTbPAO09KMqdo='
    f = Fernet(key)
    decrypted = f.decrypt(encrypted_string)
    return decrypted.decode()
#########################################################################################################################################

globalPresence = {}

with open("./memo/memory.txt", "r") as r:
    lines = r.read().splitlines()
    for line in lines:
        globalPresence[int(line)]

def logic():
    try:
        username = username_entry.get()
        password = password_entry.get()

        # #check if username valid 
        # with open('./appData/allowed_users.txt','r') as f_in:
        #     text = f_in.read()
        # text = decrypt_string(text)
        # if(username not in text):
        #     labelAuthText.set("Unauthorized")
        #     with open('./memo/log.txt','a') as f:
        #             f.write(f'unauthorized' + '\n')
        #     return
        # else:
        #     labelAuthText.set("Authorized")

        service = Service('chromedriver.exe')
        driver = webdriver.Chrome(service=service) 
        driver.get("https://agcensus.gov.in/AgriCensus/Agri_2122.jsp")
        driver.maximize_window()
        ##############################################################################################
        element = getElement(driver,By.ID,"state_list",10)
        select = Select(element)
        select.select_by_visible_text("15 Maharashtra")
        element = getElement(driver,By.ID,"user_id",10)
        operateElement("send_keys",element,username,10)
        element = getElement(driver,By.ID,"password",10)
        operateElement("send_keys",element,password,10)
        element = getElement(driver,By.ID,"captcha",10)
        text = element.text
        words = text.split()
        result = words[-1]

        element = getElement(driver,By.ID,"textbox",30)
        operateElement("send_keys",element,result,30)
        element = getElement(driver,By.ID,"Procced",30)
        operateElement("click",element,"",30)
 
        element = getElements(driver, By.CSS_SELECTOR,"span.d-lg-flex.d-sm-inline-block.ms-lg-0.ms-3", 30)
        operateElement("click",element[2],"",30)

        element = getElements(driver, By.CSS_SELECTOR,"a.dropdown-item.fw-bold[onclick^='Schedule_H_Entry']", 30)
        operateElement("click",element[0],"",30)
     

        with open("./memo/log.txt", "w") as f:
            f.write("")


        

 
        global ii
        sr_number  = 0

        time.sleep(3)
        driver.execute_script("document.body.style.zoom='75%'")
        while(ii<len(listOfQuery)):
            try:
        
                ii = ii+1
                village_index = listOfQuery[ii][0]
                sr_number = listOfQuery[ii][1]
                area_opr = listOfQuery[ii][4]
                irri = listOfQuery[ii][5]
                non_irri = listOfQuery[ii][6]
                crop_code_1 = listOfQuery[ii][7]
                crop_irri_1 = listOfQuery[ii][8]
                crop_un_irri_1 = listOfQuery[ii][9]
                crop_code_2 = listOfQuery[ii][10]
                crop_irri_2 = listOfQuery[ii][11]
                crop_un_irri_2 = listOfQuery[ii][12]
                crop_code_3 = listOfQuery[ii][13]
                crop_irri_3 = listOfQuery[ii][14]
                crop_un_irri_3 = listOfQuery[ii][15]
                souce_of_irri =  listOfQuery[ii][16]


                if irri == 'None':
                    irri = 0.0
                if non_irri == 'None':
                    non_irri = 0.0



                totalTypes = 0
                if crop_code_1 != 'None' and len(crop_code_1) > 1: 
                        if crop_irri_1 == 'None':
                            crop_irri_1 = 0
                        if crop_un_irri_1 == 'None':
                            crop_un_irri_1 = 0
                        totalTypes += 1
                if crop_code_2 != 'None' and len(crop_code_2) > 1:
                        if crop_irri_2 == 'None':
                            crop_irri_2 = 0
                        if crop_un_irri_2 == 'None':
                            crop_un_irri_2 = 0
                        totalTypes += 1
                if crop_code_3 != 'None' and len(crop_code_3) > 1:
                        if crop_irri_3 == 'None':
                            crop_irri_3 = 0
                        if crop_un_irri_3 == 'None':
                            crop_un_irri_3 = 0
                        totalTypes += 1



                element = getElement(driver,By.CSS_SELECTOR, "select[name='vlg_list']", 10)
                select = Select(element)
                select.select_by_index(village_index)

                element = getElement(driver,By.CSS_SELECTOR, "input[name='snpl_1']", 10)
                element.send_keys(Keys.BACKSPACE*15)
                operateElement("send_keys", element, sr_number, 10)
                operateElement("send_keys",element,"" + Keys.TAB, 10)

                time.sleep(3)

                element = getElement(driver,By.CSS_SELECTOR, "input[name='field_05']", 10)
                element.send_keys(Keys.BACKSPACE*15)
                element.send_keys(str(float(irri)))

                

                element = getElement(driver,By.CSS_SELECTOR, "input[name='field_06']", 10)
                element.send_keys(Keys.BACKSPACE*15)
                element.send_keys(str(float(non_irri)))

                operateElement("send_keys",element,"" + Keys.TAB, 10)

                time.sleep(1)


                
                element = getElement(driver, By.CSS_SELECTOR, "input[name='tot_crops']", 10)
                element.send_keys(Keys.BACKSPACE*15)
                element.send_keys(totalTypes)
                operateElement("send_keys",element,"" + Keys.TAB, 10)


                if totalTypes > 0:
                    element = getElement(driver, By.CSS_SELECTOR, "input#cr_code_10", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_code_1)

                    element = getElement(driver, By.CSS_SELECTOR, "input#irri_ar_10", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_irri_1)

                    element = getElement(driver, By.CSS_SELECTOR, "input#unirri_ar_10", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_un_irri_1)
                
                if totalTypes > 1:
                    element = getElement(driver, By.CSS_SELECTOR, "input#cr_code_11", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_code_2)

                    element = getElement(driver, By.CSS_SELECTOR, "input#irri_ar_11", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_irri_2)

                    element = getElement(driver, By.CSS_SELECTOR, "input#unirri_ar_11", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_un_irri_2)

                if totalTypes > 2:
                    element = getElement(driver, By.CSS_SELECTOR, "input#cr_code_12", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_code_3)

                    element = getElement(driver, By.CSS_SELECTOR, "input#irri_ar_12", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_irri_3)

                    element = getElement(driver, By.CSS_SELECTOR, "input#unirri_ar_12", 10)
                    element.send_keys(Keys.BACKSPACE*15)
                    element.send_keys(crop_un_irri_3)

                
                element = getElement(driver, By.CSS_SELECTOR, "select[name='source_irr']", 10)
                select = Select(element)
                select.select_by_index(int(souce_of_irri))
                

                element = getElement(driver, By.CSS_SELECTOR, "input.btn_1[name='Save']", 10)
                operateElement("click",element,"",30)

                element = getElement(driver, By.ID, "swal2-title", 10)

                if str(element.text).strip() == 'Record saved successfully':
                    with open('./memo/log.txt','a') as f:
                        f.write(f'{sr_number} done at {datetime.now()}' + '\n')
                    with open('./memo/success.txt','a') as f:
                        f.write(f'{sr_number}' + '\n')
                    with open('./memo/memory.txt','a') as f:
                        f.write(f'{sr_number}' + '\n')
                    
                    fillOutputArea(outputArea2,sr_number,0)
                    fillOutputArea(outputArea3,sr_number,0)
                
                else:
                    with open('./memo/failure.txt','a') as f:
                        f.write(f'{sr_number} with error {element.text}' + '\n')
                    fillOutputArea(outputArea1,sr_number,0)

    
                element = getElement(driver, By.CSS_SELECTOR, "button.swal2-confirm", 10)
                operateElement("click",element,"",30)  

                
               
                    
                
            except Exception as e:
                print(e) 
                traceback.print_exc()
                print("hello world999")
                with open('./memo/log.txt','a') as f:
                    f.write(f'failed for {sr_number} with error {e} at {datetime.now()}' + '\n')
                with open('./memo/failure.txt','a') as f:
                    f.write(f'{sr_number}' + '\n')
                fillOutputArea(outputArea1,sr_number,0)
                driver.close()
                driver.quit()
                return
            time.sleep(1)
    except Exception as e:
        print(e)
        traceback.print_exc()
        with open('./memo/log.txt','a') as f:
            f.write(f'failed  with error {e} at {datetime.now()}' + '\n')
        driver.close()
        driver.quit()
        return


def logicCaller():
    global ii
    while ii < len(listOfQuery):
        logic()
    programDone()
    

def programDone():
    labelDoneText.set("Done")
    labelDone.config(bg='green')
    

def startProgram():
    startButton.config(state=DISABLED,bg='LIGHT GREEN')
    for i in range(int(threadCount.get())):
        threading.Thread(target=logicCaller).start()
    startButton.config(state=NORMAL,bg=orig_color)


root = Tk()
root.title("Auto_Survey")
root.iconbitmap('./appData/icon.ico')
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.minsize(800,600)
root.geometry("800x600")
root.maxsize(800,600)
root.resizable(width=False, height=False)  



username_label = Label(root, text="Username:")
username_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)

# Create entry field for username
username_entry = Entry(root)
username_entry.grid(row=0, column=1, padx=5, pady=5)

# Create label for password
password_label = Label(root, text="Password:")
password_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)

# Create entry field for password
password_entry = Entry(root, show="*")
password_entry.grid(row=1, column=1, padx=5, pady=5)

#load button
load_button = Button(root, text="LOAD EXCEL FILE", command=loadExcelFile)
load_button.grid(row=2,column=1,padx=0,pady=0)
orig_color = load_button.cget("background")

latestTime = StringVar()
latestTime.set("Please Load The File")
modifiedLabel = Label(root,textvariable=latestTime).grid(row=2,column=0,padx=0)

#output labels
output_label1 = Label(root, text="FAIL")
output_label1.grid(row=4, column=1, padx=5, pady=0, sticky=W)

output_label2 = Label(root, text="SUCCESS")
output_label2.grid(row=4, column=9, padx=5, pady=0, sticky=W)

output_label3 = Label(root, text="MEMORY")
output_label3.grid(row=4, column=16, padx=5, pady=0, sticky=W)


#outputs
outputArea1 = Text(root, width=20, height=28)
outputArea1.grid(row=5, column=0, columnspan=4, padx=50, pady=0)

outputArea2 = Text(root, width=20, height=28)
outputArea2.grid(row=5, column=7, columnspan=4, padx=50, pady=0)

outputArea3 = Text(root, width=20, height=28)
outputArea3.grid(row=5, column=14, columnspan=4, padx=50, pady=0)

def on_modified1(event):
    textOfOutputArea = outputArea1.get('1.0', 'end')
    with open ("./memo/failure.txt","w") as f:
        f.write(textOfOutputArea)

def on_modified2(event):
    textOfOutputArea = outputArea2.get("1.0", "end")
    with open ("./memo/success.txt","w") as f:
        f.write(textOfOutputArea)

def on_modified3(event):
    textOfOutputArea = outputArea3.get("1.0", "end")
    with open ("./memo/memory.txt","w") as f:
        f.write(textOfOutputArea)

def on_change1(event):
    textOfOutputArea = username_entry.get()
    with open ("./memo/username.txt","w") as f:
        f.write(textOfOutputArea)

def on_change2(event):
    textOfOutputArea = password_entry.get()
    with open ("./memo/password.txt","w") as f:
        f.write(textOfOutputArea)

outputArea1.bind('<KeyRelease>', on_modified1)
outputArea2.bind('<KeyRelease>', on_modified2)
outputArea3.bind('<KeyRelease>', on_modified3)
username_entry.bind('<KeyRelease>', on_change1)
password_entry.bind('<KeyRelease>', on_change2)


with open ("./memo/failure.txt","r") as f:
    for line in f:
        fillOutputArea(outputArea1,line,1)

with open ("./memo/success.txt","r") as f:
    for line in f:
        fillOutputArea(outputArea2,line,1)

with open ("./memo/memory.txt","r") as f:
    for line in f:
        fillOutputArea(outputArea3,line,1)

with open ("./memo/username.txt","r") as f:
    for line in f:
        username_entry.insert(0,line)

with open ("./memo/password.txt","r") as f:
    for line in f:
        password_entry.insert(0,line)

#logButton
###########################################################################################
def log_button_click():
    # Create a new window
    log_window = Toplevel(root)
    log_window.title("Log Window")

    # Create a Text widget in the new window
    log_text = Text(log_window)
    log_text.pack()

    # Redirect the standard output to the Text widget
    with open("./memo/log.txt", "r") as f:
        for line in f:
            log_text.insert(END, line)
            log_text.insert(END,"\n")
    
###########################################################################################
###########################################################################################

startButton = Button(root, text="START", command=startProgram)
startButton.grid(row=0,column=16,padx=3,pady=5)

labelAuthText = StringVar()
labelAuthText.set("")
labelAuth = Label(root,textvariable=labelAuthText)
labelAuth.grid(row=0, column=10, padx=0, pady=0, sticky=W)

labelDoneText = StringVar()
labelDoneText.set("")
labelDone = Label(root,textvariable=labelDoneText)
labelDone.grid(row=0, column=7, padx=0, pady=0, sticky=W)


threadCountLabel = Label(root, text="threads:")
threadCountLabel.grid(row=1, column=8, padx=0, pady=0, sticky=W)
threadCount = Entry(root)
threadCount.grid(row=1, column=9, padx=5, pady=5)


logButton = Button(root, text="LOG", command=log_button_click)
logButton.grid(row=0,column=15,padx=3,pady=5)
root.mainloop()


