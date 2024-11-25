"""
list of invoices [description,amount,wo]
extract the wos numbers in another list arranged
make dictioanary of the description and value
loop on the wo list 
add the code for the insertion of invoice details into fiix 


"""

# import needed libraries
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import PyQt5.QtCore
from PyQt5.QtWidgets import QWidget
import PyQt5.QtWidgets
from PyQt5.uic import * # type: ignore
from os import path
import sys

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os as os
import shutil
from pynput.keyboard import Key,Controller
# from fiixi import Ui_MainWindow
uiApp,_=loadUiType(path.join(path.dirname(__file__),"fiixi.ui")) # type: ignore
class Application(QMainWindow,uiApp):
   def __init__(self):
         super(Application,self).__init__()
         QMainWindow.__init__(self)
         self.setupUi(self)
         self.workOrderTabButtonsClicks()
         self.setWindowTitle("Auto FIIX")         
         self.lastItem=False
         
   def SelectFolder(self):
      self.folderPath=QFileDialog.getExistingDirectory(self,"Choose Folder")      

      if self.folderPath:
         InvoicesFiles=[os.path.normpath(os.path.join(self.folderPath,file)) for file in os.listdir(self.folderPath) if os.path.isfile(os.path.join(self.folderPath,file))]
         self.filesL.addItems(InvoicesFiles)
      firstItem=QListWidget.item(self.filesL,0) 
      QListWidget.setCurrentItem(self.filesL,firstItem)
      self.filesL.itemClicked.emit(firstItem)
   
   def ShowInvoiceImage(self,item):
      filePath=item.text()
      self.scene=QGraphicsScene()
      image=QPixmap(filePath)
      PixmapItem=self.scene.addPixmap(image)
      self.ImageView.setScene(self.scene)
      self.ImageView.fitInView(PixmapItem,Qt.KeepAspectRatioByExpanding)
   
   def clearInputs(self,inputsWidgets:list):
      for widget in inputsWidgets:
         if isinstance(widget,QLineEdit):
            widget.clear()
            
   def RenameFile(self,row,newName):
      
      rowItem=self.filesL.item(row)
      oldName=rowItem.text()         
      extension=oldName.split('.')[-1]
      newName=f"{newName}.{extension}"
      
      self.filesL.item(row).setText(newName)
      os.rename(oldName,os.path.normpath(os.path.join(self.folderPath,newName)))
      
   def AddInvoiceInfoToList(self):
      description=self.descriptionLE.text()
      qty=self.qtyLE.text()
      amount=self.amountLE.text()
      wo=self.woLE.text()
      if description and qty and amount and wo :
         if not self.lastItem:
            nextRow=QListWidget.currentRow(self.filesL)+1
            nextItem=self.filesL.item(nextRow) 
            if nextItem:

               QListWidget.setCurrentItem(self.filesL,nextItem)
               self.filesL.itemClicked.emit(nextItem)
               oneLineInfo=f"{description}_{qty}_{amount}_{wo}"
               QListWidget.addItem(self.invoicesL,QListWidgetItem(oneLineInfo))
               self.clearInputs([self.qtyLE,self.descriptionLE,self.amountLE,self.woLE])
               QLineEdit.setFocus(self.descriptionLE)
               self.RenameFile(nextRow-1,oneLineInfo)
            else:
               oneLineInfo=f"{description}_{qty}_{amount}_{wo}"
               QListWidget.addItem(self.invoicesL,QListWidgetItem(oneLineInfo))
               self.RenameFile(nextRow-1,oneLineInfo)
               self.clearInputs([self.qtyLE,self.descriptionLE,self.amountLE,self.woLE])
               QLineEdit.setFocus(self.descriptionLE) 
               self.lastItem=True
         else:
            QMessageBox.warning(self,"End of List","No Other files in the files list")   
      else:
         QMessageBox.warning(self,"Missing Invoice Details","Please Complete Missing Invoice Details")            
  
   def getIntrnetSpeed(self):
      if self.fastRB.isChecked():
         return 3
      if self.mediumRB.isChecked():
         return 6
      if self.slowRB.isChecked():
         return 10
   def StartIssuingInvoices(self):
      internetSpeed=self.getIntrnetSpeed()
      self.folderPath=os.path.normpath(self.folderPath) 
      
      # InvoicesFiles=[os.path.normpath(os.path.join(self.folderPath,file)) for file in os.listdir(self.folderPath) if os.path.isfile(os.path.join(self.folderPath,file))]
      # for file in InvoicesFiles:
      #    fullName=file
      #    listItem=QListWidget.findItems(self.filesL,fullName,Qt.MatchExactly)

      #    if listItem:
      #       item=listItem[0]
      #       row=QListWidget.row(self.filesL,item)
      #       QListWidget.takeItem(self.filesL,row)
      #       nextItem=QListWidget.currentItem(self.filesL) 
      #       if nextItem: 
      #          self.filesL.itemClicked.emit(nextItem)
      #          QApplication.processEvents()
      #       else:
      #          self.ImageView.setScene(None)       

      fiix=Fiix(self,self.folderPath,"https://arizona.macmms.com","ahmedasser1991@gmail.com","Asser@2019",internetSpeed)
      fiix.getInvoicesDetails()
      fiix.openFiix()
      fiix.login()
      fiix.navigateToWorkOrders()
      fiix.insertInvoicesToWos()

   def workOrderTabButtonsClicks(self):

      self.filesL.itemClicked.connect(self.ShowInvoiceImage) 
      self.selectFolderPB.clicked.connect(self.SelectFolder) 
      self.addInvoiceInfoPB.clicked.connect(self.AddInvoiceInfoToList)
      self.startPB.clicked.connect(self.StartIssuingInvoices)
          

class Fiix():
   def __init__(self,App,folderPath,link,userName,password,sleepTime=5) -> None:
      
      self.folderPath=folderPath
      self.App=App
      self.link=link
      self.userName=userName
      self.password=password
      self.invoicesDetails=[]
      self.foundWos=[]
      self.notFoundWos=[]
      # setting up the browser driver
      self.options = webdriver.ChromeOptions()
      self.options.add_experimental_option("detach", True)
      # Setup chrome driver
      self.driver = webdriver.Chrome(options=self.options)  
      self.keyboard=Controller() 
      self.sleepTime=sleepTime
   def openFiix(self):
      self.driver.get('https://arizona.macmms.com')
   def login(self):
      email="/html/body/app-root/div/div/app-login/div[1]/div/form/div[1]/fiix-label/label/fiix-input/input"
      WebDriverWait(self.driver, 25).until(EC.element_to_be_clickable((By.XPATH,email ))).send_keys(self.userName)

      # login password
      password="/html/body/app-root/div/div/app-login/div[1]/div/form/div[2]/fiix-label/label/fiix-input/input"
      WebDriverWait(self.driver,25).until(EC.element_to_be_clickable((By.XPATH,password ))).send_keys(self.password)

      # click login
      enter="//form//button[contains(text(),'Log In')]"
      WebDriverWait(self.driver,25).until(EC.element_to_be_clickable((By.XPATH,enter ))).click()             
   def getInvoicesDetails(self):
      names=os.listdir(self.folderPath)     
      for inv in names:
         a=inv.removesuffix('.jpeg')
         a=a.split("_") 
         wo=a[-1]
         invDetails=a[:3]
         invDetails.append(inv)
         invDict={wo:invDetails}
         self.invoicesDetails.append(invDict)
  
   def navigateToWorkOrders(self) :
            # maintanance tap
      time.sleep(self.sleepTime)      
      maintanance='//*[@id="maMainSidebarPane"]/div/div/div[2]/div/ul/li[2]/div/span[1]'
      WebDriverWait(self.driver,30).until(EC.element_to_be_clickable((By.XPATH,maintanance ))).click()
      # choose all workorders
      all_wo='//div[@class="autoSuggestDropdownButtonContainer35"]//div[@class="autoSuggestDropdownButton35"]'
      time.sleep(self.sleepTime)
      WebDriverWait(self.driver,20).until(EC.element_to_be_clickable((By.XPATH,all_wo ))).click()

      time.sleep(self.sleepTime)
      # click on all workorders
      choose='//div//p[contains(text(),"All work orders")]'
      WebDriverWait(self.driver,20).until(EC.element_to_be_clickable((By.XPATH,choose ))).click()
      time.sleep(15)
   
   def __searchForWo(self,wo):
      searchInputElement='//span[@class="listSearchLarge"]//input[@type="text"]'
      WebDriverWait(self.driver,20).until(EC.element_to_be_clickable((By.XPATH,searchInputElement ))).send_keys(wo)      

      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,searchInputElement ))).send_keys(Keys.ENTER)
      time.sleep(self.sleepTime)
      woPargraphElmnt=f'//p[contains(text(),"{wo}")]'
      woElement=self.driver.find_element(By.XPATH,woPargraphElmnt)
      if woElement:
         return woElement
      return None
   
   def __uploadInvoiceCopyToFiles(self,invName):
      self.driver.find_element(By.XPATH,"//p[contains(text(),'Files')]").click()
      time.sleep(self.sleepTime) 
      self.driver.find_element(By.XPATH,"//div[ @class='hasBootstrapAlttextTooltip dz-clickable']").click() 
      time.sleep(self.sleepTime)   
      self.keyboard.type(os.path.join(self.folderPath,invName))
      self.keyboard.press(Key.enter)
      self.keyboard.release(Key.enter)
      time.sleep(self.sleepTime)

   def __insertMiscCostValues(self,invDetails):
      self.driver.find_element(By.XPATH,"//p[contains(text(),'Misc Costs Page')]").click()
      time.sleep(self.sleepTime)
      addCost='''//span[contains(@onclick,"executeRequestWithPush('com.ma.cmms.ui.handler.workorder.MiscCostFormExUiHandler','New'")]'''
      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,addCost))).click()
      # send keyboard keys "parts"
      time.sleep(self.sleepTime)
      self.keyboard.type('parts')

      time.sleep(self.sleepTime)
      # costType=driver.find_elements(By.XPATH,"//p[contains(text(),'Parts')]" )
      Parts='//tr[.//td[@class="listColumnValueReadOnly"]//div[@style=";height:19px;width:300px;"]//p[contains(text(),"Parts")]]'
      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,Parts))).click()

      time.sleep(self.sleepTime)
      
      # add the description
      description='input[style*="width:300px;"]'
      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.CSS_SELECTOR,description))).send_keys(invDetails[0])
      
      time.sleep(self.sleepTime)
      # press tab 4 times to enter Qty
      for press in range(4):
         self.keyboard.press(Key.tab)
         self.keyboard.release(Key.tab)

      time.sleep(self.sleepTime)
      # enter qty
      self.keyboard.type(invDetails[1])
      # move to the total cost entry field
      time.sleep(self.sleepTime)
      self.keyboard.press(Key.tab)
      self.keyboard.release(Key.tab)

      self.keyboard.press(Key.tab)
      self.keyboard.release(Key.tab)

      self.keyboard.type(invDetails[2])
        
      time.sleep(self.sleepTime)
      self.keyboard.press(Key.enter)
      self.keyboard.release(Key.enter)
      
   def __saveChanges(self):
      time.sleep(self.sleepTime)
      save='//div[contains(text(),"Save")]'
      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,save))).click()
   def __back(self):
      time.sleep(self.sleepTime)
      back='//div[@class="actionButton action noselect"][.//span[@style="float:left;" and contains(text(),"Back")]]'
      WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.XPATH,back))).click()
  
   def __moveFileToIssued(self,fileName):
      fullName=os.path.join(self.folderPath,fileName)
      shutil.move(fullName,os.path.join(os.path.split(self.folderPath)[0],'issued'))   
      listItem=QListWidget.findItems(self.App.filesL,fullName,Qt.MatchExactly)

      if listItem:
         item=listItem[0]
         row=QListWidget.row(self.App.filesL,item)
         QListWidget.takeItem(self.App.filesL,row)
         nextItem=QListWidget.currentItem(self.App.filesL) 
         if nextItem: 
            self.App.filesL.itemClicked.emit(nextItem)
            QApplication.processEvents()
         else:
            self.App.ImageView.setScene(None)       
         
      

   def __clearSearchInput(self):
      time.sleep(self.sleepTime)
      searchInputElement='//span[@class="listSearchLarge"]//input[@type="text"]'
      WebDriverWait(self.driver,20).until(EC.element_to_be_clickable((By.XPATH,searchInputElement ))).clear()      

   def __insertInvoiceDetails(self,details):
      self.__uploadInvoiceCopyToFiles(details[-1])  
      self.__insertMiscCostValues(details) 
   def insertInvoicesToWos(self):
      time.sleep(12)
      for woDict in self.invoicesDetails:
         wo=list(woDict.keys())[0]
         details=list(woDict.values())[0]
         woElement=self.__searchForWo(wo)
         if woElement:
            woElement.click() 
            time.sleep(self.sleepTime)
            self.__insertInvoiceDetails(details)
            self.__saveChanges()
            self.__moveFileToIssued(details[-1])
            self.__back()
            self.__clearSearchInput()
            self.foundWos.append(wo)
            
            
            
            
         else:
            self.notFoundWos.append(wo)   
      print(f"unCompleted issues are{self.notFoundWos}")      
      self.driver.close()
                
def main():
    mainwind=QApplication(sys.argv)
    mainWindow=Application()
    mainWindow.show()
    mainwind.exec_()
    

class Updater():
   pass

if __name__=="__main__":
    main()
             
                
                
# if __name__=="__main__":

#    fiix=Fiix("https://arizona.macmms.com","ahmedasser1991@gmail.com","Asser@2019")
#    fiix.getInvoicesDetails()
#    fiix.openFiix()
#    fiix.login()
#    fiix.navigateToWorkOrders()
#    fiix.insertInvoicesToWos()
   
