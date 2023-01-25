#Note from the author:
#---------------------
#Hello this is the creator, if are reading this code im glad you made it this far.
#This application is nothing short of awesome so i hope you enjoy it. An expert im
#not but with passion i write, will the code be clean? sure as much as i can, will
#it be functional, thats my goal. In any case, any questions, dont ask :)
#Cheers 

#Technical:
#----------
#This app is built with python eel https://github.com/python-eel/Eel a very nifty library
#that allows non C programers like me do simple but pretty built desktop applications that
#rely on good old HTML5/CSS3/Javascript/Python311.


import os
import sys
import time
import eel
import logging
import traceback
import glob
import simplejson as json
from zebra import Zebra
from pandas import *
from zplgrf import GRF

#Initialize the Zebra object
z = Zebra()

#Get current script directory
current_dir = os.getcwd()

##################### 
# Utility Functions #
#####################

#Get a list of all printers in the current machine
@eel.expose
def get_printers(): 
    return z.getqueues()

#Set selected printed as default and end raw zpl, created at the front end to the selected printer
@eel.expose
def send_label(zpl, printer, quantity):
    z.setqueue(printer)
    try:
        for _ in range(int(quantity)):
            z.output(zpl)   
    except Exception as e:
        print(e)
        logging.error(traceback.format_exc())

#Get all data out of database excel and send to front end
@eel.expose
def get_database():
    database = glob.glob('**/*.xlsm', recursive=True)

    if database:
        if len(database) > 1:
            return [1, database]
        else:
            excel = ExcelFile(database[0])
            raw = excel.parse(excel.sheet_names[1]).to_dict()
            autocomplete = list(raw['Item'].values())
            return [json.dumps(raw, ignore_nan=True), autocomplete]
    else:
        return 0

#Open a given folder
@eel.expose
def open_folder(path=current_dir):
    path = os.path.realpath(path)
    print(path)
    os.startfile(path)

#Convert image to ZPL
@eel.expose
def image_to_zpl(filename):

    #Check for the image in the current script directories
    image = glob.glob(f'{os.getcwd()}/web/assets/custom/{filename}', recursive=True)

    if image:
        with open(image[0], 'rb') as img:
            grf = GRF.from_image(img.read(), 'LABEL')
            grf.optimise_barcodes()
            zpl = grf.to_zpl()
            return zpl
    else:
        return 0

#Close app by killing system processes
@eel.expose
def close_app():
    time.sleep(8)
    try:
        os.system("taskkill /F /IM chrome.exe /T")
        os.system("taskkill /F /IM cmd.exe /T")
        sys.exit()
    except:
        pass

#Application on close call back to clear any windows left open and allow a clean restart
@eel.expose
def close_callback(route, websockets):
    if not websockets:
        print(route)
        os.system("taskkill /F /IM chrome.exe /T")
        os.system("taskkill /F /IM cmd.exe /T")
        sys.exit()

###############
# Application #
###############

# Application configuration
LABEL = {
    'cache':  True,
    'close_callback': close_callback, #remove quotes for prod
    'cmd' :['--incognito', '--no-experiments'],
    'mode': 'chrome',
    'port': 10000,
    'position': (500, 200),
    'size' : (1010, 950),
    'start': 'label.html',
}

# Application initialization
if __name__ == "__main__":
    eel.init('web')
    eel.start(
        LABEL["start"],
        close_callback = LABEL['close_callback'],
        cmdline_args = LABEL['cmd'],
        disable_cache = LABEL['cache'],
        mode = LABEL["mode"],
        port = LABEL["port"],
        position = LABEL['position'],
        size = LABEL["size"],
    )