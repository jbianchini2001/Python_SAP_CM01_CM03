# Objective: this code shows how to connect to SAP and present a graph of the capacity and the workload for a work center (likewise transactions CM01 and CM03).
# Author: Juliano Bianchini
# Contact me: jbianchini2001@gmail.com
# GNU General Public License
# Requeriments: SAP logon pad installed and acces to a SAP environment.
# How to run:
# 1) Install packages from "includes" list below;
# 2) Set values for sap_server, work_center and plant variables (see them below);
# 3) Run the code.


#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
# pywin32 is required, please install with "pip install pywin32" command. More info at https://pypi.org/project/pywin32/
# numpy is required, please install with "pip install numpy" command  
# matplotlib is required, please install with "pip install matplotlib" command  

import sys, win32com.client, time
import numpy as np
import matplotlib.pyplot as plt
import locale
import ctypes

def Main():

  # Set some definitions  
  first_row = 7

  # Change the value below accordingly to SAP system you want to connect
  sap_server = "EP0 - ECC Produção" # it must be the same description as it appears on the SAP logon
  work_center = "04030365" # your work center code
  plant = "1304"
  transaction_code = "/nCM01"


  try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      MessageBox("Error on opening SAP logon pad.", 'Error!', 0)
      return

  except:
    #SapGuiAuto =  win32com.client.gencache.EnsureDispatch('SAPGUI.Application')
    #SapGuiAuto.Visible = True
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, "Open SAP Logon PAD first and then try again.", 'Error!', 0)

  try:
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      MessageBox("Error on running SAP", 'Error!', 0)
      return

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      MessageBox("Error on running SAP", 'Error!', 0)
      return

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      MessageBox("Error on running SAP", 'Error!', 0)
      return

  except:
    # Open a new connection
    application.OpenConnection(sap_server, True)
    connection = application.Children(0)
    session = connection.Children(0)

    # Wait until SAP opens on the main screen
    while session.findById("wnd[0]/titl").text != "SAP Easy Access":
      time.sleep(3)


  # Load data from SAP transaction
  try:
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txt[35,3]").text = work_center
    session.findById("wnd[0]/usr/txt[35,7]").text = plant
    session.findById("wnd[0]/usr/txt[35,7]").setFocus()
    session.findById("wnd[0]/usr/txt[35,7]").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[5]").press()

  except:
    print('Error by running SAP transaction. Details: ', sys.exc_info() )
    return

  # table with information of work load and capacity
  # start at line 7 and a counter is required to identify the last line

  contador = first_row
  while True:
    try:
      print("wnd[0]/usr/lbl[3,{}]".format(contador))
      session.findById("wnd[0]/usr/lbl[3,{}]".format(contador)).text
      contador = contador + 1
    except:
      # continue
      break

  # Save data on arrays
  work_load = []
  capacities = []
  days_array = []


  try:
    for i in range (first_row, contador):
      # Column 3 = date
      # Column 14 = capacity requeriment
      # Column 27 = available capacity 
      # Column 41 = % of overload
      # Column 48 = free capacity
      # Column 61 = unit of measure

      # This code is to identify the locale decimal separator.
      locale.setlocale(locale.LC_ALL, '')
      loc = locale.getlocale()

      work_load.append(locale.atof(session.findById("wnd[0]/usr/lbl[14,{}]".format(i)).text.strip()))
      capacities.append(locale.atof(session.findById("wnd[0]/usr/lbl[27,{}]".format(i)).text.strip()))
      days_array.append(session.findById("wnd[0]/usr/lbl[3,{}]".format(i)).text)

    # Create graph
    maximo_x = round (max(capacities))
    if max(work_load) > max(capacities):
      maximo_x = round(max(work_load))

    ind = np.arange(start=0, stop = maximo_x, step=round ((maximo_x/(contador - first_row))))
    largura_carga = 3
    largura_capacidade = 3.8
    carga_barra = plt.bar(ind, work_load, largura_carga)
    capacidade_barra = plt.bar(ind, capacities, largura_capacidade, edgecolor = 'r', linewidth= 1.5, color = 'none')

    # Add details to the graph
    # Add labels
    add_label(carga_barra, plt)
    add_label(capacidade_barra, plt)

    plt.xlabel(session.findById("wnd[0]/usr/lbl[3,5]").text)
    plt.ylabel(session.findById("wnd[0]/usr/lbl[61,{}]".format(i)).text)
    plt.title(session.findById("wnd[0]/usr/lbl[33,1]").text + " (" + session.findById("wnd[0]/usr/lbl[16,1]").text.strip() + ")" + " - " + session.findById("wnd[0]/titl").text)
    plt.xticks(ind, days_array, rotation='vertical')
    plt.yticks(ind)
    plt.legend((carga_barra[0],capacidade_barra[0]), (session.findById("wnd[0]/usr/lbl[14,5]").text, session.findById("wnd[0]/usr/lbl[27,5]").text))

    # Show graph
    plt.show()

  except:
    print('Error: ', sys.exc_info())

  finally:
    session = None
    connection = None
    application = None
    SapGuiAuto = None

def add_label (_barra, _plot):
    for valor in _barra:
      height = valor.get_height()
      _plot.text(valor.get_x() + valor.get_width()/2., 1.01*height,
              '%s' % height,
              ha='center', va='bottom')

#-Main------------------------------------------------------------------
Main()

#-End-------------------------------------------------------------------




