from seleniumTest import SeleniumRobot
import homologaFunc
import pandas as pd
import os

########################################
# Reads students list
########################################
data = pd.read_excel ("./listado.xlsx")
listado = data.values


########################################
# Loads teacher's sigaa credentials
########################################

sigaa = pd.read_json("./sigaa.json",typ = "series")
sigaaId = sigaa.values[0]
sigaaPassword = sigaa.values[1]

########################################
# Logins in the sigaa system
########################################
robot = SeleniumRobot()
robot.setup_method()
robot.initCapp(sigaaId, sigaaPassword)

########################################
# Reads the capp for each student and 
# draws his curriculum
########################################

for item in listado:
    name = item[0]
    id = str(item[1]).zfill(9)
    line = item[4]
    robot.getCapp(id)
    homologaFunc.homologa(name,id,line)
    if os.path.exists("./capp.html"): 
        os.remove("./capp.html")
    else:
        print("{} capp.html does not exist".format(id))
    if os.path.exists("./infoAdd.html"): 
        os.remove("./infoAdd.html")
    else:
        print("{} infoAdd.html does not exist".format(id))
