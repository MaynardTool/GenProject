__author__ = 'Administrator'

# import mtTkinter as Tkinter
from xml.dom import minidom
from win32com.client import Dispatch
from platform import machine
import os
import sys
import subprocess
import time
import datetime
import shutil
import urllib2
import webbrowser
# import fcntl
import threading
from threading import Thread, Event
import platform
# from Tkinter import *
from mtTkinter import *
from ttk import *
import tkMessageBox
import tkFileDialog
from ScrolledText import *
import math
import locale
import plistlib
import zipfile
import Queue
# from Foundation import NSBundle
from ConfigParser import SafeConfigParser
import trace
import _winreg
import errno
from cfgparse.compat import ConfigParser
import fileinput

class VerticalScrolledFrame(Frame):
    """
    * Use the 'interior' attribute to place widgets inside the scrollable frame
    * Construct and pack/place/grid normally
    * This frame only allows vertical scrolling
    """
    def __init__(self, parent, *args, **kw):
        Frame.__init__(self, parent, *args, **kw)

        # create a canvas object and a vertical scrollbar for scrolling it
        global canvas

        vscrollbar = Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=FALSE)
        canvas = Canvas(self, bd=0, highlightthickness=0,
                        yscrollcommand=vscrollbar.set)
        canvas.pack(side=LEFT, fill='both', expand=TRUE)
        canvas.configure(background=defaultColor)

        vscrollbar.config(command=canvas.yview)

        # reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)

        # create a frame inside the canvas which will be scrolled with it
        self.interior = interior = Frame(canvas)
        interior_id = canvas.create_window(0, 0, window=interior, anchor=NW)
        # interior_id = canvas.create_window(5, 5, window=interior, anchor=NW)

        # track changes to the canvas and frame width and sync them,
        # also updating the scrollbar
        def _configure_interior(event):
            # update the scrollbars to match the size of the inner frame
            size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            # canvas.config(scrollregion="0 0 50 50")
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the canvas's width to fit the inner frame
                # canvas.config(width=interior.winfo_reqwidth())
                canvas.config(width=900)
        interior.bind('<Configure>', _configure_interior)

        def _configure_canvas(event):
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the inner frame's width to fill the canvas
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())
                # canvas.itemconfigure(interior_id, width=100)
        canvas.bind('<Configure>', _configure_canvas)

        # Magic part to enable two finger / mouse wheel scroll on canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(-1*(event.delta), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)


def debugPrint(textToPrint):
    if verboseDebug:
      print textToPrint


def appendToLog(content):
    timeStamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    writeToFile(logFile, ('['+timeStamp+']'+' '+content+'\n'), 'a')
    return


def execBashCommand(bashcmd):
    try:
        p1 = subprocess.Popen( bashcmd, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE )
        output = p1.communicate()[0]
        return output
    except:
        basherror = "[Error] Could not execute bash command : %s " % bashcmd
        debugPrint(basherror)
        appendToLog(basherror)
        return basherror

'''
def changeApplicationName(new_name):
  # Override application name used as the title in app menubar
  ns_bundle = NSBundle.mainBundle()
  if ns_bundle:
    ns_info = ns_bundle.localizedInfoDictionary()
  if not ns_info:
    ns_info = ns_bundle.infoDictionary()
  if ns_info:
    if ns_info['CFBundleName'] == "Python":
      ns_info['CFBundleName'] = new_name
'''


def datestampConvert(datestamp): # Not needed
    try:
      dtFormat = "%Y%m%d%H%M%S"
      dtString = str(datestamp)
      dtConverted = datetime.datetime.strptime(dtString, dtFormat)
      # Localise date and time
      dtLocaleDate = locale.nl_langinfo(locale.D_FMT)
      dtLocaleTime = locale.nl_langinfo(locale.T_FMT)
      dtShow = dtConverted.strftime(dtLocaleDate)
    except:
      # Could not get datestamp from file
      dtShow='-'
    return dtShow


def checkInternetConnection(targetUrl):
    try:
        response=urllib2.urlopen(targetUrl, timeout=4)
        debugPrint("[Notice] Internet Connection Active")
        appendToLog("[Notice] Internet Connection Active")
        return True
    except:
        debugPrint("[Error] "+targetUrl+" Offline")
        appendToLog("[Error] "+targetUrl+" Offline")
        return False


def readRemoteFile(url): # Not needed
    # Read contents of remote file
    try:
      wp = urllib2.urlopen(url)
      remoteContent = wp.read()
      remoteContent = remoteContent.strip()
      wp.close()
    except:
      debugPrint("[Error] Cannot read Remote Version File "+url)
      appendToLog("[Error] Cannot read Remote Version File "+url)
      remoteContent = 0
    return remoteContent


# convert file size to human readable format
def humanSize(num):
    for x in ['bytes','KB','MB','GB']:
      if num < 1000.0: # was 1004 think this was wrong
        return "%3.1f %s" % (num, x)
      num /= 1000.0
    return "%3.1f%s" % (num, 'TB')


# extract file extension
def fileExtension(file):
    return os.path.splitext(file)[1][1:]



def writeToFile(fileName,content,flag):
    try:
      fp = open(fileName, flag)
      fp.write(content)
      fp.close()
    except IOError:
      writeError="File Not Saved"
    except:
      writeError="File Not Saved"
      raise
    return


def deleteFile(filepath, filename): # Not needed
    try:
        # Remove file
        if os.path.exists(filepath):
          os.remove(filepath)

        debugPrint("[Notice] %s Deleted" % filename )
        appendToLog("[Notice] %s Deleted" % filename )
    except:
        debugPrint("[Error] Could not delete %s" % filename )
        appendToLog("[Error] Could not delete %s" % filename )


def downloadTxtFile(url, localdir): # Not needed
    # Download to instalDir
    try:
      webFile = urllib2.urlopen(url)
      outFileName = url.split('/')[-1]
      outFile = os.path.join(localdir,outFileName)
      localFile = open(outFile, 'w')
      localFile.write(webFile.read())
      webFile.close()
      localFile.close()
    except:
      debugPrint("[Error] Cannot download file "+url)
      appendToLog("[Error] Cannot download file "+url)


# NEED TO REFACTOR THIS USING search_zip and extract_files FUNCTIONS
def downloadFile(url, localdir, item, proginc):
    # Get Current HeaderProgress
    currentHeaderProgress = headerProgPercent.get()
    # Create dir if does not exist
    if not os.path.exists(localdir):
      os.mkdir(localdir)

    # Download to localDir
    try:
      txData = None
      txHeaders = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.10; rv:37.0) Gecko/20100101 Firefox/37.0',
        'Accept-Language': 'en-US',
        'Accept-Encoding': 'gzip, deflate, compress;q=0.9',
        'Keep-Alive': '300',
        'Connection': 'keep-alive',
        'Cache-Control': 'max-age=0',
      }
      request = urllib2.Request(url, txData, txHeaders)
      webFile = urllib2.urlopen(request)
      outFileName = url.split('/')[-1]
      appendToLog('Outfilename: %s' %outFileName)
      # replace %20 in url with spaces in filesname
      outFileName = outFileName.replace('%20', ' ')
      appendToLog('Outfilename: %s' %outFileName)
      outFile = os.path.join(localdir,outFileName)
      appendToLog('Outfile: %s' %outFile)

      localFile = open(outFile, 'wb')

      try:
        totalSize = webFile.info().getheader('Content-Length').strip()
        debugPrint("[Notice] Download file size : "+str(totalSize))
        appendToLog("[Notice] Download file size : "+str(totalSize))
        header = True
      except AttributeError:
        header = False # a response doesn't always include the "Content-Length" header
        debugPrint("[Error] No Header for "+url)
        appendToLog("[Error] No Header for "+url)
      if header:
        totalSize = int(totalSize)

      totalDownloaded = 0

      while True:
        buffer = webFile.read(8192)
        if not buffer:
          sys.stdout.write('\n')
          break

        totalDownloaded += len(buffer)
        localFile.write(buffer)
        if not header:
          totalSize = totalDownloaded # unknown size

        percent = float(totalDownloaded) / totalSize
        percent = round(percent*100, 2)
        #sys.stdout.write("Downloaded %d of %d bytes (%0.2f%%)\r" % (totalDownloaded, totalSize, percent))
        progress = round(percent) / 2

        itemProgressPercent[item].set(progress)
        headerProgPercent.set(setHeaderProgress(item, proginc, currentHeaderProgress))
        headerProgLabelTxt.set("Downloading "+titleList[item]+" ... "+str(humanSize(totalDownloaded))+" / "+str(humanSize(totalSize)))

      webFile.close()
      localFile.close()
    except:
      debugPrint("[Error] Cannot download file "+url)
      appendToLog("[Error] Cannot download file "+url)


def readFile(filepath): # Not needed
    try:
      fp = open(filepath, 'r')
      output = fp.read()
      output = output.strip()
      fp.close()
      return output
    except:
      debugPrint("[Error] Cannot read Local Version File")
      appendToLog("[Error] Cannot read Local Version File")


def on_config_changed(): # Not needed
    # Update INI file
    writeConfig()


def checkListUpdate(): # Still checking
    global localVersion

    # Update Progressbar
    splashProgressPercent.set(40)
    refreshGui(splashWindow)

    # Get remote list version
    if checkInternetConnection('http://thefanclub.co.za') == True:
      # ONLINE
      remoteVersion = readRemoteFile(remoteVersionPath)
    else:
      # OFFLINE
      remoteVersion = 0

    # Update Progressbar
    splashProgressPercent.set(50)
    refreshGui(splashWindow)

    # Decide to update list or not
    if remoteVersion > localVersion :
      debugPrint("[Notice] Downloading new software list "+remoteVersion)
      appendToLog("[Notice] Downloading new software list "+remoteVersion)
      splashProgressPercent.set(60)
      refreshGui(splashWindow)
      # Download new version
      downloadTxtFile(remoteVersionPath, installDir)
      # Download new XML list
      downloadTxtFile(remoteXmlPath, installDir)
      localVersion = remoteVersion


def on_cellall_toggle():
    #global headerCheckAllState
    itemCount = len(titleList)
    for listItem in range(itemCount):
      # Check All OFF
      if headerCheckAllState.get() == 0 and installStateList[listItem] != 'removed':
        checkItem[listItem].set(0)

      # Check All ON
      if headerCheckAllState.get() == 1 and installStateList[listItem] != 'removed' :
        if installStateList[listItem] == 'not-installed' and installButtonTxt == 'Install' or updateList[listItem] :
          checkItem[listItem].set(1)

        if installStateList[listItem] == 'installed' and installButtonTxt != 'Install' :
          checkItem[listItem].set(1)

        if installButtonTxt == 'Remove' :
          checkItem[listItem].set(1)

    # Set icon
    on_cell_toggle()


def on_cell_toggle():
    # Toggle if not busy installing
    if installStatus != 'busy' and installStatus != 'complete' :
      # count selected items to decide if check all button should be unticked
      itemCount = len(titleList)
      itemSelectCount = 0
      for listItem in range(itemCount):
        itemSelected = checkItem[listItem].get()
        if itemSelected == 1 and installStateList[listItem] != 'removed':
          itemSelectCount = itemSelectCount + 1
      # Untick Select All if no items selected
      if itemSelectCount == 0:
        headerCheckAllState.set(0)

      # check Checkboxes one by one for change and set display
      for listItem in range(itemCount):
        itemInstallState = installStateList[listItem]
        iconPathMod = ''
        # Set icon depending on installed state
        if itemInstallState == 'installed':
          # Item Selected
          if checkItem[listItem].get() == 1:
            if updateList[listItem] or installButtonTxt == 'Install':
              iconPathMod = iconPathReinstall
            else:
              iconPathMod = iconPathOk

            itemTitle[listItem].configure(foreground='#000000')
            itemDescription[listItem].configure(foreground='#555555')

            if installButtonTxt == 'Uninstall':
              iconPathMod = iconPathError

            if installButtonTxt == 'Remove':
              iconPathMod = iconPathBlank

          # Item not selected
          if checkItem[listItem].get() == 0:
            if updateList[listItem] and installButtonTxt == 'Install':
              iconPathMod = iconPathReinstall
            else:
              iconPathMod = iconPathOk

            if installButtonTxt == 'Remove':
              iconPathMod = iconPathBlank

            if installButtonTxt == 'Install' and not updateList[listItem]:
              itemTitle[listItem].configure(foreground='#AAAAAA')
              itemDescription[listItem].configure(foreground='#AAAAAA')

        if itemInstallState == 'not-installed':
          # if not
          if installButtonTxt == 'Install' :
            itemTitle[listItem].configure(foreground='#000000')
            itemDescription[listItem].configure(foreground='#555555')

        # Set icon
        if iconPathMod:
          itemIconImage[listItem] = PhotoImage(file=iconPathRetinaMod(iconPathMod))
          itemIcon[listItem].configure(image=itemIconImage[listItem])
          itemIcon[listItem].image = itemIconImage[listItem]


def refreshGui(widget):
     widget.update()
     widget.update_idletasks()


#THIS IS WORKING,  CONFIRMED NO PROBLEM WITH THIS FUNCTION#
def extract_files(fname, outfile, item, proginc):
    currentHeaderProgress = headerProgPercent.get()
    # full path of zip file
    zip_fqname = os.path.join(localDownloads, fname)

    # Create dir if does not exist
    if not os.path.exists(outfile):
        os.makedirs(outfile)

    with zipfile.ZipFile(zip_fqname, 'r') as zip_file:
        uncompress_size = sum((file.file_size for file in zip_file.infolist()))
        extracted_size = 0
        for file in zip_file.infolist():
            outpath = outfile
            extracted_size += file.file_size
            percentage = float(extracted_size / uncompress_size)
            percentage = round(percentage * 100, 2)
            # percentage = extracted_size * 100 / uncompress_size
            progress = round(percentage) / 2
            itemProgressPercent[item].set(progress)
            headerProgPercent.set(setHeaderProgress(item, proginc, currentHeaderProgress))
            headerProgLabelTxt.set(
                "Extraction progress of " + titleList[item] + " ... " + str(humanSize(extracted_size)) + " / " + str(
                    humanSize(uncompress_size)))
            # zip_file.extractall(outfile)
            zip_file.extract(file, outpath)
    return


def setHeaderProgress(item, proginc, currentprogress):
    newHeaderProgress = currentprogress + (float(itemProgressPercent[item].get()) / 100 * proginc)
    return newHeaderProgress


def get_exitcode_stdout_stderr_zip_dvd(cmd):
    ### THIS IS WORKING ###
    try:
        p1 = subprocess.Popen(cmd, shell=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE )
        output = p1.communicate()[0]
        # exitcode = p1.returncode
        return output
    except:
        basherror = "[Error] Could not execute bash command : %s " % cmd
        debugPrint(basherror)
        appendToLog(basherror)
        return basherror


def on_install_button_active(button, model, selectcount):
    # Main install Section
    global pulseTimer
    global installStatus
    global itemIconImage
    global view
    global headerProgress

    # set busy flag
    installStatus = 'busy'

    # Count items
    itemCount = len(titleList)

    # Disable Checkboxes and header Checkbox
    for listItem in range(itemCount):
        itemCheckBox[listItem].configure(state=DISABLED)
    headerCheckAll.configure(state=DISABLED)

    # START installing apps one by one

    # using itemSelectCount to do progress increments
    progInc = float(100 / selectcount)
    itemIncCount = 0

    headerProgPercent.set(0)

    for listItem in range(itemCount):
        # Check which items are selected True in list column 0
        itemSelected = checkItem[listItem].get()
        if itemSelected == 1 and installStateList[listItem] != 'removed':
            # set currentHeaderProgress for each process at start
            currentHeaderProgress = headerProgPercent.get()

            # With selected items ...
            headerLabelTxt.set('Installing Software ' + str(itemIncCount + 1) + ' of ' + str(selectcount))
            appendToLog('[Notice] Installing Software ' + str(itemIncCount + 1) + ' of ' + str(selectcount))

            # Start Install software
            installError = ''

            if updateList[listItem]:
                updateText = 'Updating'
            else:
                updateText = 'Installing'

            headerProgLabelTxt.set(updateText + " " + titleList[listItem])
            debugPrint("%s %s" % (updateText, titleList[listItem]))
            appendToLog("%s %s" % (updateText, titleList[listItem]))

            # Set Focus on item
            itemCheckBox[listItem].focus_set()
            itemCheckBox[listItem].focus()

            # Get filename
            outFileName = urlList[listItem].split('\\')[-1]  # IP_ConfigServerMT64_8130027b1_ENU_windows.zip
            appendToLog('Outfilename: %s' % outFileName)
            outFileExt = fileExtension(outFileName)  # .zip
            appendToLog('Outfileext: %s' % outFileExt)
            outFile = os.path.join(appsFolder, titleList[listItem])
            appendToLog('Outfile: %s' % outFile)

            # Modify 'IPCommon', 'InstallPath' from INI file with titleList[listItem]
            ini = os.path.join(iniFile, titleList[listItem], 'genesys_silent.ini')
            appendToLog('INI file: %s' % ini)
            config_installpath = os.path.join(installationDir, titleList[listItem])
            appendToLog('Config installation path: %s' %config_installpath)
            appendToLog('[Notice] Modifying installation path of %s' %ini)
            writeConfig(ini, 'IPCommon', 'InstallPath', config_installpath)
            appendToLog('[Notice] Successfully modified installation path of %s' %ini)

            # ini = os.path.join(iniFile, titleList[listItem], 'genesys_silent.ini')
            # config_installpath = os.path.join(installationDir, titleList[listItem])

            # for zip files, extract first
            if outFileName:
                if outFileExt.lower() == 'zip':
                    try:
                        headerProgLabelTxt.set("Extracting " + titleList[listItem] + " ...")
                        debugPrint("[Notice] Extract started for %s" % titleList[listItem])
                        appendToLog("[Notice] Extract started for %s" % titleList[listItem])
                        extract_files(outFileName, outFile, listItem, progInc)
                        debugPrint("[Notice] Extracted %s" % outFile)
                        appendToLog("[Notice] Extracted %s" % outFile)
                    except:
                        debugPrint("[Error] Could not extract %s" % outFileName)
                        appendToLog("[Error] Could not extract %s" % outFileName)
                        installError = "[Error] Could not extract %s" % outFileName

                    # If installation folder exists
                    if os.path.exists(outFile):
                        headerProgLabelTxt.set(updateText + " " + titleList[listItem] + " ...")
                        # Applications Install folder path
                        appFilePath = os.path.join(outFile, 'ip')
                        # Change directory to execute setup.exe
                        # os.chdir(appFilePath), test
                        # appendToLog('[Notice] Current directory is %s' % os.getcwd())
                        try:
                            appendToLog('[Notice] Executing setup command ...')
                            # ini = os.path.join(iniFile, titleList[listItem], 'genesys_silent.ini')
                            appendToLog('[Notice] Starting to install %s using  - %s ...' % (titleList[listItem], ini))
                            # cmd = '.\setup.exe /s /z"-s %s -sl ./genesys_install_result.log"' % ini, test
                            cmd = '%s\setup.exe /s /z"-s %s -sl %s\genesys_install_result.log"' % (appFilePath, ini, appFilePath)
                            appendToLog('[Notice] Command is %s ' % cmd)
                            get_exitcode_stdout_stderr_zip_dvd(cmd)
                        except:
                            debugPrint('[Error] Installation of %s not successful.' % titleList[listItem])
                            appendToLog('[[Error] Installation of %s not successful.' % titleList[listItem])
                            installError = '[Error] Installation of %s not successful.' % titleList[listItem]
                # for DVD, install directly
                else:
                    headerProgLabelTxt.set("Installing CD/DVD version of " + titleList[listItem] + " ...")
                    appDVDPath = urlList[listItem]
                    try:
                        appendToLog('[Notice] Executing setup command ...')
                        appendToLog('[Notice] Starting to install %s using  - %s ...' % (titleList[listItem], ini))
                        # Change directory to execute setup.exe
                        # os.chdir(appDVDPath)
                        cmd = '%s\setup.exe /s /z"-s %s -sl %s\genesys_install_result.log"' % (appDVDPath, ini, appDVDPath)
                        appendToLog('[Notice] Command is %s ' % cmd)
                        get_exitcode_stdout_stderr_zip_dvd(cmd)
                    except:
                        debugPrint('[Error] Installation of %s not successful.' % titleList[listItem])
                        appendToLog('[[Error] Installation of %s not successful.' % titleList[listItem])
                        installError = '[Error] Installation of %s not successful.' % titleList[listItem]
            else:
                debugPrint("[Error] No ZIP/DVD file for %s" % titleList[listItem])
                appendToLog("[Error] No ZIP/DVD file for %s" % titleList[listItem])
                installError = "[Error] No ZIP/DVD file for " + titleList[listItem]

            itemProgressPercent[listItem].set(90)
            headerProgPercent.set(setHeaderProgress(listItem, progInc, currentHeaderProgress))

            # Check if component if indeed installed
            # need to refactor descriptionList[listItem] with appFileList[listItem]
            appInstallPath = os.path.join(installationDir, titleList[listItem], appFileList[listItem])
            if os.path.exists(appInstallPath) and not installError:
                debugPrint("[Notice] %s installed" % titleList[listItem])
                appendToLog("[Notice] %s installed" % titleList[listItem])
            else:
                debugPrint("[Error] Could not install %s" % titleList[listItem])
                appendToLog("[Error] Could not install %s" % titleList[listItem])
                installError = "[Error] Could not install %s" % titleList[listItem]

            # END of main item Install

            # De-select checkbox, I am now here
            checkItem[listItem].set(0)

            # Check if install ok and set icon and progress bar
            if installError == '':
                iconPathMod = iconPathOk
                installStateList[listItem] = 'installed'
                itemProgressPercent[listItem].set(100)
                headerProgPercent.set(setHeaderProgress(listItem, progInc, currentHeaderProgress))
            else:
                iconPathMod = iconPathError
                installStateList[listItem] = 'error'
                debugPrint("[Error] Installation failed : %s" % installError)
                appendToLog("[Error] Installation failed : %s" % installError)
                itemProgressPercent[listItem].set(100)
                headerProgPercent.set(setHeaderProgress(listItem, progInc, currentHeaderProgress))
                # Set to 0 after main progress update
                itemProgressPercent[listItem].set(0)

            # Set icon
            itemIconImage[listItem] = PhotoImage(file=iconPathRetinaMod(iconPathMod))
            itemIcon[listItem].configure(image=itemIconImage[listItem])
            itemIcon[listItem].image = itemIconImage[listItem]

            # If selected Inc for each item as we know not how many here
            # Move progress incrementally depending on number of install items
            itemIncCount = itemIncCount + 1
            displayInc = progInc * itemIncCount

            # Update main progress bar at the end of each item install
            headerProgPercent.set(displayInc)

    # All Software Items Installed ------------
    # Software Install Done - The End -
    headerProgPercent.set(100)  # not working?!
    headerProgLabelTxt.set('')  # not working?!
    # mainWindow.event_generate('<<InstallCompleteHeaderProgressBar>>', when='tail')
    # mainWindow.event_generate('<<InstallCompleteHeaderProgressText>>', when='mark')

    headerLabelTxt.set('Installation Complete')  # not working?!)
    debugPrint('[END] Installation Complete')
    appendToLog('[END] Installation Complete')

    # Reset install status
    installStatus = 'complete'
    # Remove Cancel Button
    cancelButton.destroy()  # not working?!
    # Activate Install/Done button and menus
    installButton.configure(state=NORMAL)  # not working?!
    # menuControl('normal')
    # mainWindow.event_generate('<<InstallComplete>>', when='tail')
    refreshGui(mainWindow)


def on_install_thread():
    global loop_thread
    global installButtonTxt
    global threadLock

    stop_threads = threading.Event()

    # If button set active and label set to done exit
    if installButtonTxt == 'Done':
        # if loop_thread:
        #     joinTreads()
        # mainWindow.destroy()
        sys.exit()

    if installButtonTxt == 'Remove':
       removeSoftwareItems()
       return

    if installButtonTxt == 'Uninstall':
       uninstallSoftwareItems()
       return

    appendToLog('Install sequence initiated - Install Now')

    # Count items before we start
    itemCount = len(titleList)
    # count selected items
    itemSelectCount = 0
    for listItem in range(itemCount):
      # Check which items are selected use get to get VarInt variable instance value
      itemSelected = checkItem[listItem].get()
      if itemSelected == 1 and installStateList[listItem] != 'removed' :
        itemSelectCount = itemSelectCount + 1
        appendToLog('Itenselectcount: %s' %itemSelectCount)

    debugPrint('Number of items selected for install : %s' % str(itemSelectCount))
    appendToLog('Number of items selected for install : %s' % str(itemSelectCount))

    # Do nothing if no items selected
    if itemSelectCount == 0:
      return

    # Set button and progress
    cancelButtonTxt = 'Cancel'
    cancelButton.configure(text=cancelButtonTxt)
    installButtonTxt = 'Done'
    installButton.configure(text=installButtonTxt)
    installButton.configure(state=DISABLED)
    # Disable menu items during install
    menuControl('disabled')

    headerProgPercent.set(5)

    headerLabelTxt.set('Installing new software ...')
    headerProgLabelTxt.set('Installation Started')

    appendToLog("Installing new software ...")

    # Main Install Thread Loop
    loop_thread = threading.Thread(target=on_install_button_active, args=['button', 'model', itemSelectCount])
    # Start Thread as Daemon - so that all threads terminate when window closes or user cancels and quits
    loop_thread.daemon = True
    # Start Main Install Threaded
    loop_thread.start()


    appendToLog('Install Thread started')


def joinTreads():
    global loop_thread
    # Join Threads every second for 0.5 sec and relax
    try:
      loop_thread.join(timeout=0.02)
    except:
      debugPrint('[Error] Cannot Join Threads')
      appendToLog('[Error] Cannot Join Threads')
    return True


def on_cancel_button(widget=''):
    global installButtonTxt
    global checkItemOldState

    if installStatus == 'busy':
      if tkMessageBox.askyesno(appName, "Are you sure you would like to Quit?", icon='question'):
        headerLabelTxt.set('Application Cancelled')
        headerProgPercent.set(100)
        headerProgLabelTxt.set('')
        debugPrint('[Warning] Application Cancelled.')
        appendToLog('[Warning] Application Cancelled.')
      else:
        return

    if str(widget) == '' and installButtonTxt == 'Remove' or installButtonTxt == 'Uninstall':
      # Reset after remove or uninstall for install
      installButtonTxt = 'Install'
      installButton.configure(text=installButtonTxt)

      cancelButtonTxt = 'Quit'
      cancelButton.configure(text=cancelButtonTxt)

      headerLabelTxt.set('Select the software you would like to install')
      headerProgLabelTxt.set('Click Install to Start')

      # Reset Icons and Checkboxes and progress bars
      for listItem in range(len(titleList)):
        # Reset checkbox
        checkItem[listItem].set(checkItemOldState[listItem])
        itemCheckBox[listItem].configure(state=NORMAL)
        itemProgressPercent[listItem].set(0)
        itemInfo[listItem].configure(text='i', state=NORMAL)

      # Flush temp var
      checkItemOldState = []
      # Update Display
      headerCheckAllState.set(1)
      headerProgPercent.set(0)

      # Reset Remove Button
      removeButton.configure(state=NORMAL)

      # Do display magic
      on_cell_toggle()
      refreshGui(mainWindow)
      return

    # Check if list has changed and prompt for save
    if listHasChanged == True :
      if tkMessageBox.askyesno(title=appName, message=xmlFilename + " list has changed", detail = "Save the changes to the list before quiting?", icon='warning'):
        # Export list before exit
        on_export_list()

    # If not busy just quit
    if loop_thread:
      joinTreads()
    mainWindow.destroy()
    sys.exit()


def renderPlatformDialog():
    if tkMessageBox.showinfo(appName, "Only 64bit versions of Windows supported", icon='warning'):
      splashWindow.destroy()
      sys.exit()


def renderAppClose(appname, pids):
    # Split pid's in array
    pids = pids.split()
    if tkMessageBox.askyesno(title=appName, message=appname + " is currently in use", detail = "Close and install new version?", icon='warning'):
      # Kill each PID of app
      for appProcess in pids:
        bashCommand = 'kill -9 %s' % appProcess
        execBashCommand(bashCommand)
      return


def renderOfflineDialog():
    if tkMessageBox.showinfo(title=appName + " - Offline", message="No Internet Connection", detail="Connect to the internet and try again.", icon='warning'):
      splashWindow.destroy()
      sys.exit()


def renderErrorDialog(errortxt, detailtxt):
    if tkMessageBox.showinfo(title=appName + " - Error", message=errortxt, detail=detailtxt, icon='error'):
      splashWindow.destroy()
      sys.exit()


def pixelRetinaMod(pixels):
    pixelmod = int(round(pixels * dpiScale))
    return pixelmod


def iconPathRetinaMod(iconpath):
    fileext = os.path.splitext(iconpath)[1]
    # Start big to small
    if dpiScale >= 4:
      return os.path.splitext(iconpath)[0] + 'X4' + fileext
    if dpiScale >= 3:
      return os.path.splitext(iconpath)[0] + 'X3' + fileext
    if dpiScale >= 2:
      return os.path.splitext(iconpath)[0] + 'X2' + fileext
    # else just return default
    return iconpath


def renderAboutDialog():
    global aboutWindow

    # Create Add Software window
    thisYear = datetime.datetime.now().strftime("%Y")

    aboutWindow = Toplevel()
    aboutWindow.title('About ')
    aboutWindow.configure(background=defaultColor)
    aboutWindow.resizable(FALSE, FALSE)

    # Center the main window
    x = (aboutWindow.winfo_screenwidth() - aboutWindow.winfo_reqwidth()) / 2
    y = (aboutWindow.winfo_screenheight() - aboutWindow.winfo_reqheight()) / 2
    aboutWindow.geometry("+%d+%d" % (x-pixelRetinaMod(40), y-pixelRetinaMod(140)))

    # Create an logo Image GIF only with PhotoImage
    aboutLogoImage = PhotoImage(file=iconPathRetinaMod(iconPath))
    aboutLogoLabel = Label(aboutWindow, image=aboutLogoImage)
    aboutLogoLabel.image = aboutLogoImage
    aboutLogoLabel.pack()

    # Create Text Label
    appNameLabel = Label(aboutWindow, text=appName, font=('default', 20, 'bold'))
    appNameLabel.pack()

    appVersionLabel = Label(aboutWindow, text='Version ' + appVersion)
    appVersionLabel.pack()

    appListLabel = Label(aboutWindow, text='List Date ' + datestampConvert(localVersion), font=('default', 11, 'normal'))
    appListLabel.pack()

    appCreditLabel = Label(aboutWindow, text=u"\u00a9 2013-"+thisYear+" The Fan Club", font=('default', 11, 'normal'))
    appCreditLabel.pack(pady=(pixelRetinaMod(10), pixelRetinaMod(5)))

    refreshGui(aboutWindow)


def renderStartupSplash():
    global splashWindow
    global splashProgressBar
    global splashProgressPercent
    global splashLabel
    global splashLabelText
    global dpiScale

    # Create Main window
    splashWindow = Tk()
    splashWindow.title(appName+' ' + appVersion)
    splashWindow.configure(background=defaultColor)
    splashWindow.resizable(FALSE, FALSE)

    # Center window
    x = (splashWindow.winfo_screenwidth() - splashWindow.winfo_reqwidth()) / 2
    y = (splashWindow.winfo_screenheight() - splashWindow.winfo_reqheight()) / 2

    # DPI for Retina
    dpi = splashWindow.winfo_fpixels('1i')
    dpiScale = float(dpi)/72
    #dpiScale = 1.4

    splashWindow.geometry("+%d+%d" % (x-pixelRetinaMod(50), y-pixelRetinaMod(80)))

    # Create an logo Image GIF only with PhotoImage
    logoImage = PhotoImage(file=iconPathRetinaMod(iconPath))
    logoLabel = Label(splashWindow, image=logoImage)
    logoLabel.image = logoImage

    # Create progress bar
    splashProgressPercent = IntVar()
    splashProgressBar = Progressbar(splashWindow, orient=HORIZONTAL, length=pixelRetinaMod(280), mode='determinate')
    splashProgressBar['variable'] = splashProgressPercent

    # Set initial progress bar value
    splashProgressBar.step(1)
    splashProgressPercent.set(10)

    # Create Text Label
    splashLabelText = StringVar()
    splashLabel = Label(splashWindow, textvariable=splashLabelText)
    splashLabelText.set('Initializing...')

    # a grid to attach the elements
    logoLabel.grid(column=0, row=0)
    splashLabel.grid(column=0, row=1, columnspan=1)
    splashProgressBar.grid(column=0, row=2, padx=pixelRetinaMod(10), pady=(pixelRetinaMod(3), pixelRetinaMod(10)))


def on_new_list(widget=''):
    global titleList
    global descriptionList
    global appFileList
    global urlList
    global versionList
    global selectBox
    global progressBox
    global installStateList
    global updateList
    global iconPathList
    global xmlPath
    global xmlFilename

    global checkItem
    global itemCheckBox
    global itemTitle
    global itemVersion
    global itemDescription
    global itemProgressPercent
    global itemProgress
    global itemIconImage
    global itemIcon
    global itemInfo

    # Remove all from display
    removeSoftwareItems('all')

    # Clear Arrays
    titleList = []
    descriptionList = []
    appFileList = []
    urlList = []
    versionList = []
    selectBox = []
    progressBox = []
    installStateList = []
    updateList = []
    iconPathList = []

    checkItem = []
    itemCheckBox = []
    itemTitle = []
    itemVersion = []
    itemDescription = []
    itemProgressPercent = []
    itemProgress = []
    itemInfo = []
    itemIconImage = []
    itemIcon = []

    xmlPath = xmlPathNew
    xmlFilename = xmlFilenameNew

    headerCheckAllState.set(0)
    # Set new window title
    mainWindow.title(appName + ' - ' + xmlFilename)

    # Update, display and call add software window
    onListHasChanged(False)

    # Open add new software dialog
    on_add_edit_software()


def on_import_list(widget=''):
    # Check if revert was clicked else open file dialog
    if str(widget) == 'revert':
      # check if file is set for reload
      newXmlFilePath = xmlPath
    else:
      # Render file open dialog
      myfiletypes = [('MAI Files', '*.mai'), ('XML Files', '*.xml'), ('All files', '*')]
      newXmlFilePath = tkFileDialog.askopenfilename(title='Choose a Software List File', initialdir=userHome, filetypes=myfiletypes)

    if os.path.exists(newXmlFilePath):
      # Remove lock
      fcntl.flock(lockFile, fcntl.LOCK_UN|fcntl.LOCK_NB)
      # Close MainWindow
      mainWindow.destroy()
      # Restart program
      python = sys.executable
      # Add new argv before restart
      if len(sys.argv) == 1:
        sys.argv.append(newXmlFilePath)
      else:
        sys.argv[1] = newXmlFilePath
      # Restart
      os.execl(python, python, * sys.argv)


def removeSoftwareItems(items=''):
    # Remove software from list
    for listItem in range(len(titleList)):
      # If selected remove from list
      if checkItem[listItem].get() == 1 or items == 'all':
        # Flag list has change
        onListHasChanged(True)
        # Remove Items from display
        itemCheckBox[listItem].grid_forget()
        itemTitle[listItem].grid_forget()
        itemVersion[listItem].grid_forget()
        itemDescription[listItem].grid_forget()
        itemProgress[listItem].grid_forget()
        itemIcon[listItem].grid_forget()
        itemInfo[listItem].grid_forget()
        # Update state
        installStateList[listItem] = 'removed'
    # Set Select All Checkbox
    headerCheckAllState.set(0)


def on_remove_software(widget=''):
    global installButtonTxt
    global cancelButtonTxt
    global checkItemOldState

    # Switch to remove mode
    headerLabelTxt.set('Select the software you would like to remove from the list')
    headerProgLabelTxt.set('Click Remove to Start')
    installButtonTxt = 'Remove'
    installButton.configure(text=installButtonTxt)
    cancelButtonTxt = 'Back'
    cancelButton.configure(text=cancelButtonTxt)
    # Disable Remove button
    removeButton.configure(state=DISABLED)

    for listItem in range(len(titleList)):
      if not len(checkItemOldState) < listItem:
        # Capture current state
        checkItemOldState.append(checkItem[listItem].get())

      # Deselect all
      checkItem[listItem].set(0)

      # Make sure all checkboxes are active
      itemCheckBox[listItem].configure(state=NORMAL)

      # Clear progress bars
      itemProgressPercent[listItem].set(0)

      # Remove Greyed out items if any
      itemTitle[listItem].configure(foreground='#000000')
      itemDescription[listItem].configure(foreground='#555555')

      # Change Info i button to Remove - button
      itemInfo[listItem].configure(text='-', state=NORMAL)

    # Set Select All Checkbox
    headerCheckAllState.set(0)
    # Set Header progress
    headerProgPercent.set(0)

    # Set icons and text according to state
    on_cell_toggle()


def readConfig(configfilepath, section, key):
    # Read Config
    config = SafeConfigParser()
    try:
        config.read(configfilepath)
        # Get product GUID for un-installation process
        appendToLog('Retrieving product GUID of %s ...' % configfilepath)
        value = config.get(section, key)
        appendToLog('[Notice] Successfully retrieved product GUID - %s' % value)
        return value
    except:
        debugPrint("[Error] Cannot read config : [%s] %s" % (section, key))
        appendToLog("[Error] Cannot read config : [%s] %s" % (section, key))


def writeConfig(configfilepath, section, key, value):
    config = ConfigParser()
    # Update INI file of a particular component
    try:
        config.read(configfilepath)
        config.set(section, key, value)
        with open(configfilepath, 'wb') as configFile:
            config.write(configFile)
    except:
        debugPrint("[Error] Cannot write config file")
        appendToLog("[Error] Cannot write config file")


def get_pguid(setupini):
    # Get product GUID for un-installation process
    # should receive path C:\Users\Administrator\Downloads\Genesys\ConfigServer\ip
    productGuid = ''
    try:
        appendToLog('Retrieving product GUID of %s ...' % setupini)
        # config = ConfigParser.ConfigParser()
        config = SafeConfigParser()
        config.read(setupini)
        productGuid = config.get('Startup', 'ProductGUID')
        time.sleep(3)
        appendToLog('[Notice] Successfully retrieved product GUID - %s' % productGuid)
        return productGuid
    except:
        appendToLog('[Error] Error occurred while retrieving product GUID %s' % productGuid)
        sys.exit()


def get_last_pguid(hdrfile):
    marker = '\x00\x00\x00\x00\x00\x00<Support>\x00\x00<Support>'
    lines = []
    f = open(hdrfile, "rb")
    searchlines = f.readlines()
    f.close()
    for i, line in enumerate(searchlines):
        if marker in line:
            text_with_guid = line.split('<Support>')
            # print(i) # line number
            # fifth element is the GUID
            lines.append(text_with_guid[4])

    # Convvert to string
    guid = ''.join(lines)
    # Strip off trailing \
    guid = guid.lstrip('\\')
    # Replace NULL character and convert to uppercase
    guid = guid.replace('\x00', '').upper()

    return guid


def replace_guid(iss_file, hdr_file, new_pguid):
    last_pguid = (get_last_pguid(hdr_file))
    pattern = re.compile(r'\b([A-Z0-9]{8})-([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{12})\b')
    count = 0
    for line in fileinput.input(iss_file, inplace=1, backup='.bak'):
        if count != 13:
            line = re.sub(pattern, new_pguid, line.rstrip())
            print(line)
        else:
            line = re.sub(pattern, last_pguid, line.rstrip())
            print(line)
        count += 1


def find_uninstallString(hkey, keypath, app, pguid):
    # Find uninstall string from the registry
    GUID = '{%s}' %pguid
    appendToLog('[Notice] Product GUID retrieved from AppsFolder - %s is %s' % (app, GUID))
    key = _winreg.OpenKey(hkey, keypath, 0,  _winreg.KEY_READ)
    for i in xrange(0, _winreg.QueryInfoKey(key)[0]):
        # product = '{009B1211-56D8-44FE-9933-5ABB66E17513}'
        skey_name = _winreg.EnumKey(key, i)
        # print(skey_name)
        skey = _winreg.OpenKey(key, skey_name)
        # print(skey)
        try:
            # appendToLog('[Notice] Checking if %s exists in the registry.' % GUID)
            guid = _winreg.QueryValueEx(skey, 'ProductGuid')[0]
            appendToLog('[Notice] Product GUID retrieved from the registry is %s' % guid)
            appendToLog('Checking if product GUID matches.')
            if GUID == guid:
                appendToLog('[Notice] Product GUID matches.')
                appendToLog('Retrieving uninstall string...')
                uninstallstring = _winreg.QueryValueEx(skey, 'UninstallString')[0]
                path = uninstallstring.split('"')
                path = str(path[1])

                stripExe = path.rsplit('\\', 1)
                stripExe = str(stripExe[0])
                # print(stripExe)
                appendToLog('[Notice] Successfully retrieved uninstall string: %s' % stripExe)
                return stripExe
            else:
                appendToLog('[Error] Not able to find uninstall string: ProductGuid mismatch')
        except OSError as e:
            if e.errno == errno.ENOENT:
                # DisplayName doesn't exist in this skey
                pass
        finally:
            skey.Close()


def uninstallSoftwareItems():
    # Declare variables
    keypath = r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"  # to be used to traverse registry
    # tested and working, using only one ISS file
    iss = os.path.join(issFIle, 'uninstall.iss')

    # Disable buttons
    installButton.configure(state=DISABLED)
    cancelButton.configure(state=DISABLED)
    headerCheckAll.configure(state=DISABLED)

    # Disable menu items during install
    menuControl('disabled')

    # Count items before we start
    itemCount = len(titleList)
    # count selected items
    itemSelectCount = 0
    for listItem in range(itemCount):
        # Check which items are selected use get to get VarInt variable instance value
        itemSelected = checkItem[listItem].get()
        if itemSelected == 1:
            itemSelectCount = itemSelectCount + 1

    debugPrint('Number of items selected for uninstall : %s' % str(itemSelectCount))
    appendToLog('Number of items selected for uninstall : %s' % str(itemSelectCount))

    # Do nothing if no items selected
    if itemSelectCount > 0:
        # using itemSelectCount to do progress increments
        progInc = float(100 / itemSelectCount)
        itemIncCount = 0
        headerProgPercent.set(0)

    for listItem in range(itemCount):
        # If selected remove from list
        if checkItem[listItem].get() == 1:
            # set currentHeaderProgress for each process at start
            currentHeaderProgress = headerProgPercent.get()

            # With selected items ...
            headerLabelTxt.set('Uninstalling Software ' + str(itemIncCount + 1) + ' of ' + str(itemSelectCount))
            appendToLog('[Notice] Uninstalling Software ' + str(itemIncCount + 1) + ' of ' + str(itemSelectCount))

            # Start Uninstall
            installError = ''
            headerProgLabelTxt.set("Uninstalling " + titleList[listItem])
            debugPrint("[Notice] Uninstalling %s" % titleList[listItem])
            appendToLog("[Notice] Uninstalling %s" % titleList[listItem])

            # Uninstall component from registry
            appInstallPath = os.path.join(installationDir, titleList[listItem], appFileList[listItem])
            if os.path.exists(appInstallPath):
                # path of setup.ini
                setupINI = os.path.join(appsFolder, titleList[listItem], 'ip', 'setup.ini')
                try:
                    if os.path.exists(setupINI):
                        # pGUId = get_pguid(setupINI)
                        # for ZIP files
                        # setupINI = os.path.join(appsFolder, titleList[listItem], 'ip', 'setup.ini')
                        hdr_file = os.path.join(appsFolder, titleList[listItem], 'ip', 'data1.hdr')
                        pGUID = readConfig(setupINI, 'Startup', 'ProductGUID')
                        replace_guid(iss, hdr_file, pGUID)
                    else:
                        # for DVD files
                        setupINI = os.path.join(urlList[listItem], 'setup.ini')
                        hdr_file = os.path.join(urlList[listItem], 'data1.hdr')
                        pGUID = readConfig(setupINI, 'Startup', 'ProductGUID')
                        replace_guid(iss, hdr_file, pGUID)
                    # proceed with the uninstallation
                    uninstall_path = find_uninstallString(_winreg.HKEY_LOCAL_MACHINE, keypath, titleList[listItem], pGUID)
                    # os.chdir(uninstall_path), test
                    # appendToLog('[Notice] Current directory is %s' % os.getcwd())
                    appendToLog('Starting to uninstall %s using  - %s...' % (titleList[listItem], iss))
                    # cmd = 'setup.exe /s /f1"%s"' % iss
                    cmd = '"%s\setup.exe" /s /f1"%s"' % (uninstall_path, iss)
                    appendToLog('[Notice] Command is %s ' % cmd)
                    # call get_exitcode_stdout_stderr_zip_dvd
                    get_exitcode_stdout_stderr_zip_dvd(cmd)
                    debugPrint("[Notice] Uninstalled : %s" % titleList[listItem])
                    appendToLog("[Notice] Uninstalled : %s" % titleList[listItem])
                    appendToLog('Cleaning up ...')
                    # remove leftover path
                    shutil.rmtree(uninstall_path, ignore_errors=True)
                    time.sleep(3)
                    # clean empty folder
                    # not working since dir is being used by another program
                    # The process cannot access the file because it is being used by another process
                    # removeEmptyFolders(uninstall_path)
                    appendToLog('Cleanup done.')
                except:
                    debugPrint("[Error] Could not uninstall : %s" % titleList[listItem])
                    appendToLog("[Error] Could not uninstall : %s" % titleList[listItem])
                    installError = "[Error] Could not uninstall %s" % titleList[listItem]

            # Set Checkbox off
            checkItem[listItem].set(0)
            itemCheckBox[listItem].configure(state=DISABLED)

            # Check if uninstall ok and set icon and progress bar
            if not installError:
                iconPathMod = iconPathBlank
                installStateList[listItem] = 'not-installed'
                itemProgressPercent[listItem].set(100)
                headerProgPercent.set(setHeaderProgress(listItem, progInc, currentHeaderProgress))
                # Grey out text
                itemTitle[listItem].configure(foreground='#AAAAAA')
                itemDescription[listItem].configure(foreground='#AAAAAA')
            else:
                iconPathMod = iconPathError
                debugPrint("[Error] Uninstall failed : %s" % installError)
                appendToLog("[Error] Uninstall failed : %s" % installError)
                itemProgressPercent[listItem].set(100)
                headerProgPercent.set(setHeaderProgress(listItem, progInc, currentHeaderProgress))
                # Set to 0 after main progress update
                itemProgressPercent[listItem].set(0)

            # Set Icon
            itemIconImage[listItem] = PhotoImage(file=iconPathRetinaMod(iconPathBlank))
            itemIcon[listItem].configure(image=itemIconImage[listItem])
            itemIcon[listItem].image = itemIconImage[listItem]

            # If selected Inc for each item as we know not how many here
            # Move progress incrementally depending on number of install items
            itemIncCount = itemIncCount + 1
            displayInc = progInc * itemIncCount

            # Update main progress bar at the end of each item install
            headerProgPercent.set(displayInc)
            refreshGui(mainWindow)

    # Set Progress, Buttons, Menu and Txt after
    if itemSelectCount > 0:
        headerLabelTxt.set('Uninstall Complete')
        headerProgPercent.set(100)
        headerProgLabelTxt.set('')
        debugPrint('[END] Uninstall Complete')
        appendToLog('[END] Uninstall Complete')

    # Enable buttons
    installButton.configure(state=NORMAL)
    cancelButton.configure(state=NORMAL)
    headerCheckAll.configure(state=NORMAL)
    # Enable menu items
    menuControl('normal')
    refreshGui(mainWindow)


def on_uninstall_software():
    global installButtonTxt
    global cancelButtonTxt
    global checkItemOldState

    # Switch to uninstall mode if not installing
    headerLabelTxt.set('Select the software you would like to uninstall')
    headerProgLabelTxt.set('Click Uninstall to Start')
    installButtonTxt = 'Uninstall'
    installButton.configure(text=installButtonTxt)
    cancelButtonTxt = 'Back'
    cancelButton.configure(text=cancelButtonTxt)

    for listItem in range(len(titleList)):
        if not len(checkItemOldState) < listItem:
            # Capture current state if not already done
            checkItemOldState.append(checkItem[listItem].get())

        # Deselect all
        checkItem[listItem].set(0)
        # Change Info Button to Uninstall button
        itemInfo[listItem].configure(text='x')

        # Disabled items that are not installed
        if installStateList[listItem] == 'not-installed':
            itemCheckBox[listItem].configure(state=DISABLED)
            itemInfo[listItem].configure(state=DISABLED)
            # Grey out text
            itemTitle[listItem].configure(foreground='#AAAAAA')
            itemDescription[listItem].configure(foreground='#AAAAAA')
        else:
            itemTitle[listItem].configure(foreground='#000000')
            itemDescription[listItem].configure(foreground='#555555')

    # Set De-Select All Checkbox
    headerCheckAllState.set(0)
    # Set icons and text according to state
    on_cell_toggle()


def processInstallVersion(title, appfile, version):
    # Checks versions for update
    (installedItem, versionLocalItem) = checkInstall(title, appfile)

    updateItem = False

    if installedItem:
      # Check if software can be updated
      updateItem = checkSoftwareUpdate(versionLocalItem, version)

      debugPrint("Installed : %s" % title)
      appendToLog("Installed : %s" % title)
      debugPrint("Version Local : [%s]" % str(versionLocalItem))
      appendToLog("Version Local : [%s]" % str(versionLocalItem))
      debugPrint("Version List  : [%s]" % str(version))
      appendToLog("Version List  : [%s]" % str(version))

      # Set Checkbox and Icon
      if updateItem:
        selectBox.append(1)
        iconpathmod = iconPathReinstall
        debugPrint(title + " update : " + str(updateItem))
        appendToLog(title + " update : " + str(updateItem))
      else:
        selectBox.append(0)
        iconpathmod = iconPathOk

      progressvalue = 0
      installStateList.append('installed')
    else:
      selectBox.append(1)
      progressvalue = 0
      iconpathmod = iconPathBlank
      installStateList.append('not-installed')

    return (progressvalue, iconpathmod, updateItem)


def AddSoftwareItem(item):
    # Parse add software input and add to display
    missingInput = ''

    if not versionInput.get():
      missingInput = versionInput
      inputError = "Enter Version Number"

    if not urlInput.get():
      missingInput = urlInput
      inputError = "Enter File Download URL"

    if not appFileInput.get():
      missingInput = appFileInput
      inputError = "Enter Application Filename"

    if not descriptionInput.get():
      missingInput = descriptionInput
      inputError = "Enter Description"

    if not titleInput.get():
      missingInput = titleInput
      inputError = "Enter Title"

    # Check if item is already on the list
    if titleInput.get() in titleList and item == '':
      missingInput = titleInput
      inputError = titleInput.get()+' already in the list'

    # Check if URL exists last
    if urlInput.get() and not missingInput:
      fileExt = fileExtension(os.path.basename(urlInput.get()))
      if not fileExt.lower() in ('dmg', 'bz', 'tgz', 'tar', 'gz', 'bz2', 'zip'):
        missingInput = urlInput
        inputError = 'Invalid file download URL'
      else:
        if not checkInternetConnection(urlInput.get()):
          missingInput = urlInput
          inputError = 'Download URL does not exist'

    # Create input loop until no more items are missing
    if missingInput:
      #print "Error : "+inputError
      missingInput.focus_set()
      # Show error
      errorLabel.configure(text=inputError)
    else:
      # If no more missing items then continue
      addSoftwareWindow.withdraw()

      # Append List arrays with new item
      # Get and set selectBox, iconPathMod, progressValue, installStateList
      (progressValue, iconPathMod, updateFlag) = processInstallVersion(titleInput.get(), appFileInput.get(), versionInput.get())

      if item == '':
        # Append to arrays
        titleList.append(titleInput.get())
        descriptionList.append(descriptionInput.get())
        appFileList.append(appFileInput.get())
        urlList.append(urlInput.get())
        versionList.append(versionInput.get())
        # Build Progress blank array
        progressBox.append(progressValue)
        # Build Icon Path mod array
        iconPathList.append(iconPathMod)
        updateList.append(updateFlag)
        newIndex = len(titleList)-1
        newState = 'add'
      else:
        # Edit current item if item has value
        titleList[item] = titleInput.get()
        descriptionList[item] = descriptionInput.get()
        appFileList[item] = appFileInput.get()
        urlList[item] = urlInput.get()
        versionList[item] = versionInput.get()
        # Build Progress blank array
        progressBox[item] = progressValue
        # Build Icon Path mod array
        iconPathList[item] = iconPathMod
        updateList[item] = updateFlag
        newIndex = item
        newState = 'update'

      debugPrint('[Notice] %s %s on list' % (titleInput.get(), newState))
      appendToLog('[Notice] %s %s on list' % (titleInput.get(), newState))
      # Close input window when vars processed
      addSoftwareWindow.destroy()

      # Append to display by using array data
      addItemToListDisplay(newIndex, newState)

      # Flaag list changed
      onListHasChanged(True)

      # Update display
      refreshGui(mainWindow)


def on_info_button(item):
	  # Decide if button is pressed as info or uninstall
    if installButtonTxt == 'Install' :
      # Info Button - View / Update
	    on_add_edit_software(item)

    if installButtonTxt == 'Uninstall' or installButtonTxt == 'Remove':
      # Uninstall Item
      # First De-select all if any have been selected via checkbox
      for listItem in range(len(titleList)):
        # Deselect all
        checkItem[listItem].set(0)
      # Set De-Select All Checkbox
      headerCheckAllState.set(0)
      # Then Select item
      checkItem[item].set(1)
      # Disable button
      itemInfo[item].configure(state=DISABLED)
      # Update display
      on_cell_toggle()
      # Uninstall item
      if installButtonTxt == 'Uninstall':
        if tkMessageBox.askyesno(appName, "Are you sure you would like to " + installButtonTxt+u" \n\n\u25cf  " + titleList[item], icon='question'):
          uninstallSoftwareItems()
        else:
          # De-Select item
          checkItem[item].set(0)
          # Enable button
          itemInfo[item].configure(state=NORMAL)
          # Update display
          on_cell_toggle()

      # Remove item
      if installButtonTxt == 'Remove':
        removeSoftwareItems()


def on_add_edit_software(item=''):
	  # Render window and use for info, update and add item
    global addSoftwareWindow
    global titleInput
    global descriptionInput
    global appFileInput
    global urlInput
    global versionInput
    global titleInputTxt
    global descriptionInputTxt
    global appFileInputTxt
    global urlInputTxt
    global versionInputTxt
    global errorLabel

    # Override binding call for cmd+s
    if 'instance' in str(item):
      item = ''

    # Add Button and heading text
    if item != '':
      addItemButtonTxt = 'Update'
    else:
      addItemButtonTxt = 'Add'

    # Check if add / update window open already
    if addSoftwareWindow:
      addSoftwareWindow.destroy()

    # Create Add Software window
    addSoftwareWindow = Toplevel()
    addSoftwareWindow.title(appName + ' - ' + addItemButtonTxt + ' Software')
    addSoftwareWindow.configure(background=defaultColor)
    addSoftwareWindow.resizable(FALSE, FALSE)
    addSoftwareWindow.geometry('%dx%d+0+0' % ((pixelRetinaMod(340), pixelRetinaMod(245))))

    # Center the main window
    x = (addSoftwareWindow.winfo_screenwidth() - addSoftwareWindow.winfo_reqwidth()) / 2
    y = (addSoftwareWindow.winfo_screenheight() - addSoftwareWindow.winfo_reqheight()) / 2
    addSoftwareWindow.geometry("+%d+%d" % (x-pixelRetinaMod(60), y-pixelRetinaMod(65)))

    addSoftwareWindow.lift()

    # Add Error Label
    errorLabel = Label(addSoftwareWindow, text='', foreground='#333333', anchor=CENTER, justify=CENTER, font=('default', 11, 'normal'))
    # Add Title to Top Frame
    titleLabel = Label(addSoftwareWindow, text='Title:')
    # Add Title input
    titleInput = Entry(addSoftwareWindow)
    # Add Description to Top Frame
    descriptionLabel = Label(addSoftwareWindow, text='Description:')
    # Add Description input
    descriptionInput = Entry(addSoftwareWindow)
    # Add Appfile to Top Frame
    appFileLabel = Label(addSoftwareWindow, text='App Filename:')
    # Add Appfile input
    appFileInput = Entry(addSoftwareWindow)

    # Add URL Label
    urlLabel = Label(addSoftwareWindow, text='Download URL:')
    # Add URL input
    urlInput = Entry(addSoftwareWindow)

    # Add Version label
    versionLabel = Label(addSoftwareWindow, text='Version:')
    # Add Version input
    versionInput = Entry(addSoftwareWindow)

    # If Info mode then display data we have
    if addItemButtonTxt == 'Update':
      titleInput.delete(0, END)
      titleInput.insert(0, titleList[item])
      descriptionInput.delete(0, END)
      descriptionInput.insert(0, descriptionList[item])
      appFileInput.delete(0, END)
      appFileInput.insert(0, appFileList[item])
      urlInput.delete(0, END)
      urlInput.insert(0, urlList[item])
      versionInput.delete(0, END)
      versionInput.insert(0, versionList[item])

    # Add Cancel Button
    cancelItemButton = Button(addSoftwareWindow, text='Cancel', command=addSoftwareWindow.destroy)

    addItemButton = Button(addSoftwareWindow, text=addItemButtonTxt, command=lambda i=item: AddSoftwareItem(i))

    # Pack grid
    errorLabel.grid(row=0, column=1, columnspan=2, pady=pixelRetinaMod(5), padx=pixelRetinaMod(10))

    titleLabel.grid(row=2, column=0, columnspan=1, pady=pixelRetinaMod(3), padx=(pixelRetinaMod(30), pixelRetinaMod(5)), sticky='E')
    titleInput.grid(row=2, column=1, columnspan=2, pady=pixelRetinaMod(3), padx=0, sticky='we')

    descriptionLabel.grid(row=3, column=0, columnspan=1, pady=pixelRetinaMod(3), padx=(pixelRetinaMod(30), pixelRetinaMod(5)), sticky='E')
    descriptionInput.grid(row=3, column=1, columnspan=2, pady=pixelRetinaMod(3), padx=0, sticky='we')

    appFileLabel.grid(row=4, column=0, columnspan=1, pady=pixelRetinaMod(3), padx=(pixelRetinaMod(30), pixelRetinaMod(5)), sticky='E')
    appFileInput.grid(row=4, column=1, columnspan=2, pady=pixelRetinaMod(3), padx=0, sticky='we')

    urlLabel.grid(row=5, column=0, columnspan=1, pady=pixelRetinaMod(3), padx=(pixelRetinaMod(30), pixelRetinaMod(5)), sticky='E')
    urlInput.grid(row=5, column=1, columnspan=2, pady=pixelRetinaMod(3), padx=0, sticky='we')

    versionLabel.grid(row=6, column=0, columnspan=1, pady=pixelRetinaMod(3), padx=(pixelRetinaMod(30), pixelRetinaMod(5)), sticky='E')
    versionInput.grid(row=6, column=1, columnspan=2, pady=pixelRetinaMod(3), padx=0, sticky='we')

    cancelItemButton.grid(row=8, column=1, columnspan=1, pady=(pixelRetinaMod(25), pixelRetinaMod(10)), padx=pixelRetinaMod(5), sticky='E')
    addItemButton.grid(row=8, column=2, columnspan=1, pady=(pixelRetinaMod(25), pixelRetinaMod(10)), padx=pixelRetinaMod(5), sticky='E')

    refreshGui(addSoftwareWindow)


def on_show_log():
    global showLogWindow
    # Create Add Software window
    showLogWindow = Toplevel()
    showLogWindow.title(appName + ' - Log')
    showLogWindow.configure(background=defaultColor)
    showLogWindow.resizable(TRUE, TRUE)
    showLogWindow.geometry('%dx%d+0+0' % (pixelRetinaMod(600), pixelRetinaMod(400)))

    # Center the main window
    x = (showLogWindow.winfo_screenwidth() - showLogWindow.winfo_reqwidth()) / 2
    y = (showLogWindow.winfo_screenheight() - showLogWindow.winfo_reqheight()) / 2
    showLogWindow.geometry("+%d+%d" % (x-pixelRetinaMod(60), y-pixelRetinaMod(65)))

    textPad = ScrolledText(showLogWindow, width=120, height=50)
    #textPad.grid(row=0, column=0, columnspan=3, pady=10, padx=5, sticky='news')
    textPad.pack(padx=0, pady=0, fill=BOTH, expand=True)

    contents = readFile(logFile)

    textPad.insert('1.0', contents)

    textPad.configure(state="disabled")
    # make sure the widget gets focus when clicked
    # on, to enable highlighting and copying to the
    # clipboard.
    textPad.bind("<1>", lambda event: textPad.focus_set())

    refreshGui(showLogWindow)


def onListHasChanged(status):
    global listHasChanged
    # set global flag
    listHasChanged = status
    # Set menu save button and window title
    if status == True:
      fileMenu.entryconfig("Save", state=NORMAL)
      mainWindow.title(appName+' - ' + xmlFilename + '*')
    if status == False:
      fileMenu.entryconfig("Save", state=DISABLED)
      mainWindow.title(appName + ' - ' + xmlFilename)


def writeXmlFile(xmlfilepath):
    # Check if the file exists and delete
    if os.path.exists(xmlfilepath):
      deleteFile(xmlfilepath, os.path.basename(xmlfilepath))

    # Build XML File
    contentHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<MacAfterInstall>\n\t<Software>\n'
    writeToFile(xmlfilepath, contentHeader, 'a')

    # Add each item
    for listItem in range(len(titleList)):
      if installStateList[listItem] != 'removed':
        contentLine = '\t\t<Item>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t\t<Title>' + titleList[listItem].replace(' & ', ' &amp; ') + '</Title>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t\t<Description>' + descriptionList[listItem].replace(' & ', ' &amp; ') + '</Description>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t\t<AppFile>' + appFileList[listItem].replace(' & ', ' &amp; ') + '</AppFile>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t\t<URL>' + urlList[listItem].replace(' & ', ' &amp; ') + '</URL>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t\t<Version>' + versionList[listItem].replace(' & ', ' &amp; ') + '</Version>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

        contentLine = '\t\t</Item>\n'
        writeToFile(xmlfilepath, contentLine, 'a')

    contentFooter = '\t</Software>\n</MacAfterInstall>'
    writeToFile(xmlfilepath, contentFooter, 'a')

    debugPrint('[Notice] Saved list to : %s' % xmlfilepath)
    appendToLog('[Notice] Saved list to : %s' % xmlfilepath)
    # Flag list state now as unchanged
    onListHasChanged(False)

# Dito nako, yehey!
def on_export_list(widget=''):
    global xmlPath
    global xmlFilename
    # On File Save as dialog
    myfiletypes = [('MAI Files', '*.mai'), ('XML Files', '*.xml'), ('All files', '*')]
    # Remove file extension
    currentXmlFilename = os.path.basename(xmlPath)
    currentXmlFilename = os.path.splitext(currentXmlFilename)[0]

    currentXmlFileDir = os.path.dirname(xmlPath)

    if currentXmlFileDir == installDir or not os.path.exists(xmlPath):
      currentXmlFileDir = userHome
      widget = 'save-as'

    if widget == 'save-as':
      # Render Save As Dialog
      newXmlFilePath = tkFileDialog.asksaveasfilename(title='Enter List Name', initialdir=currentXmlFileDir, initialfile=currentXmlFilename, defaultextension='mai', filetypes=myfiletypes)
      # return if no file selected
      if not newXmlFilePath:
        return
      # Add extension if missing
      if fileExtension(newXmlFilePath).lower() != 'xml' and fileExtension(newXmlFilePath).lower() != 'mai':
        newXmlFilePath = newXmlFilePath + '.mai'

      xmlPath = newXmlFilePath
      xmlFilename = os.path.basename(xmlPath)

    # Write List to File
    writeXmlFile(xmlPath)


def on_help_documentation():
    webbrowser.open(projectUrl, new=2, autoraise=True)


def on_help_github():
    webbrowser.open(projectGithubUrl, new=2, autoraise=True)


def menuControl(menuState):
    global addButton
    global removeButton
    # Either normal or disabled
    softwareMenu.entryconfig("Add to List", state=menuState)
    softwareMenu.entryconfig("Remove from List", state=menuState)
    softwareMenu.entryconfig("Uninstall Software", state=menuState)
    fileMenu.entryconfig("New List", state=menuState)
    fileMenu.entryconfig("Open", state=menuState)
    fileMenu.entryconfig("Revert to Saved List", state=menuState)

    # Do menu buttons
    if menuState == 'disabled':
      addButton.configure(state=DISABLED)
      removeButton.configure(state=DISABLED)
    if menuState == 'normal':
      addButton.configure(state=NORMAL)
      removeButton.configure(state=NORMAL)


def addItemToListDisplay(item, state):
    global checkItem
    global itemCheckBox
    global itemTitle
    global itemVersion
    global itemDescription
    global itemProgressPercent
    global itemProgress
    global itemIconImage
    global itemIcon
    global itemInfo

    if state == 'add':
      checkItem.append(IntVar())
      itemCheckBox.append(Checkbutton(frame.interior, variable=checkItem[item], command=on_cell_toggle))
      # I can increase the width
      itemTitle.append(Label(frame.interior, text=titleList[item], width=20, justify=CENTER))
      itemVersion.append(Label(frame.interior, text=versionList[item], foreground='#AAAAAA', width=20, justify=CENTER, anchor=CENTER))
      itemDescription.append(Label(frame.interior, text=descriptionList[item], foreground='#555555', width=30, justify=CENTER))
      itemProgressPercent.append(IntVar())
      itemProgress.append(Progressbar(frame.interior, orient=HORIZONTAL, length=pixelRetinaMod(150), mode='determinate'))
      # Need to understand this
      itemProgress[item]['variable'] = itemProgressPercent[item]

      itemIconImage.append(PhotoImage(file=iconPathRetinaMod(iconPathList[item])))
      itemIcon.append(Label(frame.interior, image=itemIconImage[item]))
      itemIcon[item].image = itemIconImage[item]

      itemInfo.append(Button(frame.interior, text='i', command=lambda i=item: on_info_button(i), width=1))

      checkItem[item].set(selectBox[item])

      # Add to layout
      itemCheckBox[item].grid(row=item, column=0, padx=20)
      itemTitle[item].grid(row=item, column=1, padx=10)
      itemVersion[item].grid(row=item, column=2)
      itemDescription[item].grid(row=item, column=3, padx=15)
      itemProgress[item].grid(row=item, column=4, padx=10)
      itemIcon[item].grid(row=item, column=5, padx=(pixelRetinaMod(5), pixelRetinaMod(5)))
      itemInfo[item].grid(row=item, column=6, padx=(pixelRetinaMod(5), pixelRetinaMod(10)))

    if state == 'update':
      itemTitle[item].configure(text=titleList[item])
      itemVersion[item].configure(text=versionList[item])
      itemDescription[item].configure(text=descriptionList[item])

      itemIconImage[item] = PhotoImage(file=iconPathRetinaMod(iconPathList[item]))
      itemIcon[item].configure(image=itemIconImage[item])
      itemIcon[item].image = itemIconImage[item]

    # Set Checkbox depending on install state
    # must set value because IntVar object instance is created
    # must use checkItem[i].get to get current state
    # Grey out installed items
    if installStateList[item] == 'installed' and iconPathList[item] != iconPathReinstall and installButtonTxt != 'Remove':
      # Grey out text
      itemTitle[item].configure(foreground='#AAAAAA')
      itemDescription[item].configure(foreground='#AAAAAA')


def on_install_complete(event):
    cancelButton.destroy()
    installButton.configure(state=NORMAL)


def on_install_complete_header_progress_bar(event):
    headerProgPercent.set(100)


def on_install_complete_header_progress_text(event):
    headerProgLabelTxt.set('')
    headerLabelTxt.set('Installation Complete')

def exit_program(event):
    sys.exit()


def removeEmptyFolders(path, removeRoot=True):
    'Function to remove empty folders'
    if not os.path.isdir(path):
        return

    # remove empty subfolders
    files = os.listdir(path)
    if len(files):
        for f in files:
            fullpath = os.path.join(path, f)
            if os.path.isdir(fullpath):
                removeEmptyFolders(fullpath)

    # if folder empty, delete it
    files = os.listdir(path)
    if len(files) == 0 and removeRoot:
        print "Removing empty folder:", path
        os.rmdir(path)


def renderMainWindow():
    global headerLabelTxt
    global headerProgress
    global headerProgPercent
    global headerProgLabelTxt
    global headerCheckAll
    global headerCheckAllState
    global itemProgressPercent
    global installButtonTxt
    global cancelButton
    global installButton
    global mainWindow
    global checkItem
    global itemTitle
    global itemDescription
    global itemInfo
    global itemCheckBox
    global itemIcon
    global itemIconImage
    global itemVersion
    global itemProgress
    global frame
    global softwareMenu
    global fileMenu
    global addButton
    global removeButton
    global checkAutoUpdate
    global checkBackup

    # Create Main window
    mainWindow = Tk()
    mainWindow.title(appName + ' - ' + xmlFilename)
    mainWindow.wm_iconbitmap(iconWindow) # this is mine
    mainWindow.configure(background=defaultColor)
    mainWindow.resizable(TRUE, TRUE)
    scaledWidth = pixelRetinaMod(650)
    scaledHeight = pixelRetinaMod(540)
    mainWindow.geometry('%dx%d+0+0' % (scaledWidth, scaledHeight))
    mainWindow.maxsize(width=scaledWidth, height=10000)
    mainWindow.minsize(width=scaledWidth, height=scaledHeight)


    # Center the main window
    x = (mainWindow.winfo_screenwidth() - mainWindow.winfo_reqwidth()) / 2
    y = (mainWindow.winfo_screenheight() - mainWindow.winfo_reqheight()) / 2
    mainWindow.geometry("+%d+%d" % (x-pixelRetinaMod(190), y-pixelRetinaMod(175)))

    # Create Menu Bar
    menuBar = Menu(mainWindow)

    # create App Menu
    # need to update this control key for Windows
    mainMenu = Menu(menuBar, tearoff=0, name='windows') # name creates system app menu
    mainMenu.add_command(label="About " + appName, command=renderAboutDialog)
    mainMenu.add_separator()
    mainMenu.add_command(label="Quit " + appName, command=lambda: on_cancel_button('quit'), accelerator="Ctrl+Q")
    mainMenu.bind_all("<Control-q>", on_cancel_button)
    menuBar.add_cascade(label=appName, menu=mainMenu)

    # create File Menu
    # create a pulldown menu, and add it to the menu bar
    fileMenu = Menu(menuBar, tearoff=0)
    fileMenu.add_command(label="New List", command=on_new_list, accelerator="Ctrl+N")
    fileMenu.bind_all("<Control-n>", on_new_list)
    fileMenu.add_command(label="Open", command=on_import_list, accelerator="Ctrl+O")
    fileMenu.bind_all("<Control-o>", on_import_list)
    fileMenu.add_command(label="Save", command=on_export_list, accelerator="Ctrl+S")
    fileMenu.bind_all("<Control-s>", on_export_list)
    fileMenu.add_command(label="Save As", command=lambda: on_export_list('save-as'), accelerator="Shift+Ctrl+S")
    fileMenu.bind_all("<Shift-Control-s>", on_export_list)
    fileMenu.add_separator()
    fileMenu.add_command(label="Revert to Saved List", command=lambda: on_import_list('revert'))
    fileMenu.add_separator()
    fileMenu.add_command(label="Quit "+appName, command=lambda : on_cancel_button('quit'), accelerator="Ctrl+Q")
    menuBar.add_cascade(label="File", menu=fileMenu)

    # create the Software Menu
    softwareMenu = Menu(menuBar, tearoff=0)
    # Autoupdate
    checkAutoUpdate = BooleanVar()
    # Set value from config.
    # checkAutoUpdate.set(configAutoUpdate)
    softwareMenu.add_checkbutton(label="Auto Update List", onvalue=True, offvalue=False, variable=checkAutoUpdate, command=on_config_changed) # think I don't need this
    softwareMenu.add_separator()
    softwareMenu.add_command(label="Add to List", command=on_add_edit_software, accelerator="Ctrl+A")
    softwareMenu.bind_all("<Control-a>", on_add_edit_software)
    softwareMenu.add_command(label="Remove from List", command=on_remove_software, accelerator="Ctrl+R")
    softwareMenu.bind_all("<Control-r>", on_remove_software)
    softwareMenu.add_separator()
    softwareMenu.add_command(label="Uninstall Software", command=on_uninstall_software)
    softwareMenu.add_separator()
    # Backup DMG, for MAC OSX only
    # checkBackup = BooleanVar()
    # Set value from Config
    # checkBackup.set(configBackup)
    # softwareMenu.add_checkbutton(label="Create Backup DMG", onvalue=True, offvalue=False, variable=checkBackup, command=on_config_changed)
    menuBar.add_cascade(label="Software", menu=softwareMenu)

    # create Help Menu
    # mainWindow.createcommand('::tk::mac::ShowHelp', on_help_documentation) # override mac sys Help menu item, MAC only
    helpMenu = Menu(menuBar, tearoff=0, name='_help') # name creates system app menu note underscore
    helpMenu.add_separator()
    helpMenu.add_command(label="Project Home Page", command=on_help_documentation)
    #helpMenu.add_separator()
    helpMenu.add_command(label="Project on GitHub", command=on_help_github)
    helpMenu.add_separator()
    helpMenu.add_command(label="View Log File", command=on_show_log)
    menuBar.add_cascade(label="Help", menu=helpMenu)

    # display the menu
    mainWindow.configure(menu=menuBar)

    # Create Top Frame
    topFrame = Frame(mainWindow)
    topFrame.pack(side=TOP, fill=X)

    # Add Header to Top Frame
    headerLabelTxt = StringVar()
    headerLabel = Label(topFrame, textvariable=headerLabelTxt, justify=CENTER, anchor=CENTER)
    headerLabelTxt.set('Select the software you would like to install')
    headerLabel.pack(pady=pixelRetinaMod(10))

    headerProgPercent = IntVar()
    headerProgress = Progressbar(topFrame, length=pixelRetinaMod(560), orient=HORIZONTAL, mode='determinate')
    headerProgress['variable'] = headerProgPercent
    headerProgress.pack()

    headerProgLabelTxt = StringVar()
    headerProgressLabel = Label(topFrame, textvariable=headerProgLabelTxt, foreground='#777777', justify=CENTER)
    headerProgLabelTxt.set('Click Install to Start')
    headerProgressLabel.pack(pady=(0, pixelRetinaMod(10)))

    headerCheckAllState = IntVar()
    headerCheckAll = Checkbutton(topFrame, variable=headerCheckAllState, command=on_cellall_toggle)
    headerCheckAllState.set(1)
    headerCheckAll.pack(side=LEFT, padx=pixelRetinaMod(14))

    # Bottom Frame
    bottomFrame = Frame(mainWindow)
    bottomFrame.pack(side=BOTTOM, fill=X)

    addButtonTxt = '+'
    addButton = Button(bottomFrame, text=addButtonTxt, command=on_add_edit_software, width=1)
    addButton.pack(side=LEFT, pady=pixelRetinaMod(10), padx=(pixelRetinaMod(15), pixelRetinaMod(1)))

    removeButtonTxt = '-'
    removeButton = Button(bottomFrame, text=removeButtonTxt, command=on_remove_software, width=1)
    removeButton.pack(side=LEFT, pady=pixelRetinaMod(10), padx=pixelRetinaMod(1))

    installButtonTxt = 'Install'
    installButton = Button(bottomFrame, text=installButtonTxt, command=on_install_thread, width=9)
    installButton.pack(side=RIGHT, pady=pixelRetinaMod(10), padx=(pixelRetinaMod(2), pixelRetinaMod(35)))

    cancelButtonTxt = 'Quit'
    cancelButton = Button(bottomFrame, text=cancelButtonTxt, command=on_cancel_button, width=9)
    cancelButton.pack(side=RIGHT, pady=pixelRetinaMod(10), padx=(pixelRetinaMod(5), pixelRetinaMod(2)))

    # Middle Frame RENDER LAST
    # Call VerticalScrolledFrame to do the frame + grid + scrollbar magic
    # Need to fix this
    frame = VerticalScrolledFrame(mainWindow)
    frame.pack(side="right", fill="y", padx=(pixelRetinaMod(0), 0))
    # frame.pack(side="right", fill="y")

    # Build list display line for line as array append
    checkItem = []
    itemCheckBox = []
    itemTitle = []
    itemVersion = []
    itemDescription = []
    itemProgressPercent = []
    itemProgress = []
    itemInfo = []
    itemIconImage = []
    itemIcon = []

    # Add each item to 'frame.interior' NOT mainWindow for scrollbars
    for i in range(len(titleList)):
      addItemToListDisplay(i, 'add')

    # Set save menu item and window title
    onListHasChanged(listHasChanged)

    # mainWindow.bind('<<InstallCompleteHeaderProgressBar>>', on_install_complete_header_progress_bar)
    # mainWindow.bind('<<InstallCompleteHeaderProgressText>>', on_install_complete_header_progress_text)
    # mainWindow.bind('<<ExitProgram>>', exit_program)

    # Main display rendered
    appendToLog('[Notice] Main display rendered')
    # End of Main Menu Render


def getText(nodelist):
    rc = []
    for node in nodelist:
        if node.nodeType == node.TEXT_NODE:
            rc.append(node.data)
    return ''.join(rc)


def checkSoftwareUpdate(localver, newver):
    # Determine which version is newer
    # Remove word 'Build' if present
    tempVerLocalPre = localver.replace('Build', '').strip()
    tempVerListPre = newver.replace('Build', '').strip()

    # Remove bit after space - assume version is first
    tempVerLocalPre = tempVerLocalPre.split(' ')[0]
    tempVerListPre = tempVerListPre.split(' ')[0]

    # Split . into array
    tempVerLocal = tempVerLocalPre.split('.')
    tempVerList = tempVerListPre.split('.')

    # Count array items
    tempVerLocalLen = len(tempVerLocal)
    tempVerListLen = len(tempVerList)

    # Get version with least items to compare
    if tempVerLocalLen > tempVerListLen:
      verItems = tempVerListLen
    else:
      verItems = tempVerLocalLen

    setUpdateItem = False

    try:
        while True:
          for verNumItem in range(verItems):
            # Break if ordinal value higher than local
            if int(tempVerList[verNumItem]) > int(tempVerLocal[verNumItem]):
              setUpdateItem = True
              break
            # Break if ordinal value less than local
            if int(tempVerList[verNumItem]) < int(tempVerLocal[verNumItem]):
              break
          # break out of while loop when done checking each number
          break
    except:
        debugPrint('[Warning] Cannot compare version numbers')
        appendToLog('[Warning] Cannot compare version numbers')

    return setUpdateItem

# For Mac OS X only
def parsePlistXml(url, searchkey):
    # Get value for plist xml using plistlib
    dom = plistlib.readPlist(url)
    try:
        searchkeyvalue = dom[searchkey]
    except:
        searchkeyvalue = ''
    return searchkeyvalue

def get_version_number(filename):
    parser = Dispatch('Scripting.FileSystemObject')
    version = parser.GetFileVersion(filename)
    return version


# Check if software is installed, TODO
def checkInstall(subfolder, ext):
    # subfolder = ConfigServer
    # software = confserv.exe
    appPath = os.path.join(installationDir, subfolder, ext)
    # Check each item and if one is missing assume not installed
    if os.path.exists(appPath):
        isInstalled = True  # true if it is installed
        # Check installed version of app
        try:
            appInstVersion = get_version_number(appPath)
        except:
            print('[Warning] Cannot get installed version of %s' % ext)
            appInstVersion = '0'
    else:
        isInstalled = False
        appInstVersion = '0'

    # if installed return true/false and app installed version
    return (isInstalled, str(appInstVersion))



############################# Main Loop
if __name__ == "__main__":

    # Vars
    verboseDebug = True

    # Main Env Vars
    appName = 'Auto Install'
    procName = 'auto-install'
    appVersion = '1.0 beta'
    userHome = os.getenv('USERPROFILE')
    timeStamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dateStamp = datetime.datetime.now().strftime("%Y-%m-%d")
    installDir = os.getcwd()
    installationDir = os.path.join(os.environ['PROGRAMFILES'], 'GCTI')
    localDownloads = os.path.join(userHome, 'Downloads')
    appsFolder = os.path.join(localDownloads, 'Genesys')
    # appsFolder = os.path.join(os.environ['PROGRAMFILES(X86)']) # for testing purpose
    logFile = os.path.join(os.environ['APPDATA'], 'autoinstall', 'auto-install.log')
    lockFilePath = os.path.join(os.environ['APPDATA'], 'autoinstall', 'lock', 'auto-install.lock')
    iconPath = os.path.join(installDir, 'lib', 'logos', 'autoinstall.gif')
    iconPathError = os.path.join(installDir, 'lib', 'icons', 'red.gif')
    iconPathReinstall = os.path.join(installDir, 'lib', 'icons', 'orange.gif')
    iconPathOk = os.path.join(installDir, 'lib', 'icons', 'green.gif')
    iconPathBlank = os.path.join(installDir, 'lib', 'icons', 'dark-grey.gif')
    iconWindow = os.path.join(installDir, 'lib', 'icons', 'computer.ico')
    xmlFilename = procName + '.xml'
    xmlFilenameNew = 'Untitled.mai'
    xmlPath = os.path.join(installDir, xmlFilename)
    xmlPathNew = os.path.join(userHome, xmlFilenameNew)
    configFilename = procName + '.ini'
    configOsPath = os.path.join(userHome, '.config')
    configFilePath = os.path.join(configOsPath, appName, configFilename)
    VersionFilename = 'version.txt'
    localVersionPath = os.path.join(installDir, VersionFilename)
    # PRE-CONFIGURED FILE
    iniFile = os.path.join(localDownloads, 'INI')
    issFIle = os.path.join(localDownloads, 'ISS')
    testdict = {
        'Google Chrome': 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    }
    '''
    Think I don't need this
    remoteVersionDir = 'https://www.thefanclub.co.za/sites/default/files/mac-after-install/'
    remoteVersionPath = os.path.join(remoteVersionDir, VersionFilename)
    remoteXmlPath = os.path.join(remoteVersionDir, xmlFilename)
    '''
    # localDownloads = os.path.join(userHome, 'Downloads', 'autoinstall_downloads')
    '''
    Think I don't need this
    localBackups = '/tmp/mac-after-install_backups'
    localBackupsDmgTmp = os.path.join(localDownloads, appName.replace(' ', '')+'_'+dateStamp+'.temp.dmg')
    # localBackupsDmg = os.path.join(userHome, 'Downloads', appName.replace(' ', '')+'_'+dateStamp+'.dmg')
    localBackupDmgVolName = (appName+" - Backup "+dateStamp).replace(' ', '\ ')
    projectUrl = 'https://www.thefanclub.co.za/how-to/mac-after-install/'
    '''
    projectGithubUrl = 'https://github.com/thefanclub/mac-after-install/'  # TODO

    # App Vars
    connectionStatus = False
    installStatus = ''
    loop_thread = ''
    pulseTimer = 0
    timerAsync = 0
    progressTimer = 0
    defaultColor = '#E9E9E9'
    listHasChanged = False
    checkItemOldState = []
    addSoftwareWindow = ''

    # Set Application Title
    # changeApplicationName(appName)

    # Lock File, TODO for Windows
    '''
    try:
        lockFile = open(lockFilePath,'w')
        # Try to aquire lock
        fcntl.flock(lockFile, fcntl.LOCK_EX|fcntl.LOCK_NB)
        # File has not been locked before
        fileIsLocked = False
    except:
        # File is already locked
        fileIsLocked = True
    '''
    '''
    if fileIsLocked:
      sys.exit('[Notice] '+procName+' instance already running or you do not have admin rights to run the program.')

    lockFile.write('%d\n'%os.getpid())
    lockFile.flush()
    '''
    # Start Log file after use appendToLog
    writeToFile(logFile, ('[' + timeStamp + '] ' + appName + ' ' + appVersion + ' - Started' + '\n'), 'w')
    debugPrint("[Notice] Log file created")
    appendToLog("[Notice] Log file created")

    # Get Dist Info
    debugPrint("Platform : %s" % platform.platform())
    appendToLog("Platform : %s" % platform.platform())
    debugPrint("Machine : %s" % platform.machine())
    appendToLog("Machine : %s" % platform.machine())

    '''
    Think I don't need this
    # Get and Set Locale
    # locale.setlocale(locale.LC_ALL, 'en_GB.UTF-8')
    currentLocale = locale.getlocale()
    debugPrint("Current Locale : %s" % str(currentLocale))
    appendToLog("Current Locale : %s" % str(currentLocale))
    '''

    # START Splash Screen
    renderStartupSplash()
    refreshGui(splashWindow)
    # need this to render splash screen at this point

    # Check if system is 64bit
    if not machine() == 'AMD64':
        renderPlatformDialog()

    # Check internet connection
    splashLabelText.set('Checking internet connection...')
    splashProgressPercent.set(20)
    refreshGui(splashWindow)

    '''
    Comment out since you know Bulacan net!
    if checkInternetConnection('http://www.google.com.ph') == False:
      # OFFLINE
      renderOfflineDialog()
    '''

    # Read Config File if present TODO
    '''
    if os.path.exists(configFilePath):
      splashLabelText.set('Reading configuration...')
      splashProgressPercent.set(30)
      refreshGui(splashWindow)


      # Set defaults if available in config file
      # Use eval to get bool value from string
      configAutoUpdate = eval(readConfig(configFilePath, 'main', 'autoupdate'))
      if configAutoUpdate == None or configAutoUpdate == '' :
        # Default if no value in config
        configAutoUpdate = True

      configBackup = eval(readConfigget(configFilePath, 'main', 'backup'))
      if configBackup == None or configBackup == '' :
        # Default if no value in config
        configBackup = False

    else:
      # Set Defaults if no Config File exists
      configAutoUpdate = True
      configBackup = False
    '''

    # Get local list version
    localVersion = readFile(localVersionPath)
    if not localVersion:
        localVersion = 0

    # Check for list update
    '''
    if configAutoUpdate == True :
      splashLabelText.set('Checking for updates...')
      splashProgressPercent.set(40)
      refreshGui(splashWindow)


      # Check for online updates of list if selected
      checkListUpdate()
    '''

    # Parse the XML install file
    splashLabelText.set('Checking installed software...')
    splashProgressPercent.set(70)
    refreshGui(splashWindow)

    appendToLog('Parsing XML...')

    # Swap out XML file is list is imported or dragged in .xml or .mai format
    if len(sys.argv) > 1:
        if os.path.exists(sys.argv[1]) and fileExtension(sys.argv[1]) == 'xml' or fileExtension(sys.argv[1]) == 'mai':
            xmlPath = sys.argv[1]
            xmlFilename = os.path.basename(xmlPath)
            # convert time in sec since epoch from file to time object
            tempTime = time.gmtime(os.path.getmtime(xmlPath))  # gmtime = get the UTC time, gettmtime = last modifcation
            localVersion = time.strftime("%Y%m%d%H%M%S", tempTime)

    debugPrint("List Path : %s" % xmlPath)
    appendToLog("List Path : %s" % xmlPath)
    debugPrint("List Filename : %s" % xmlFilename)
    appendToLog("List Filename : %s" % xmlFilename)

    # Load XML file for parsing and create opject
    try:
        xmldoc = minidom.parse(xmlPath)
    except:
        debugPrint("[Error] Cannot Parse XML List File : %s" % xmlPath)
        appendToLog("Error] Cannot Parse XML List File : %s" % xmlPath)
        renderErrorDialog('Invalid List File', 'Cannot process ' + os.path.basename(xmlPath))

    # Declare arrays and get xml obj
    titleList = []
    xmlTitleListObj = xmldoc.getElementsByTagName('Title')

    descriptionList = []
    xmlDescriptionListObj = xmldoc.getElementsByTagName('Description')

    appFileList = []
    xmlAppFileListObj = xmldoc.getElementsByTagName('AppFile')

    urlList = []
    xmlUrlListObj = xmldoc.getElementsByTagName('URL')

    versionList = []
    xmlVersionListObj = xmldoc.getElementsByTagName('Version')

    # Declare select box bool value array and progress
    selectBox = []
    progressBox = []
    installStateList = []
    updateList = []
    iconPathList = []

    itemCount = len(xmlTitleListObj)
    item = 0
    realCount = 0

    appendToLog('Building ListStore')

    # Copy xml obj values to build arrays for easy reference of each item
    # and build liststore for Gtk
    for item in range(itemCount):
        # First item is the select box value
        # Check if software is installed here for each item before setting True
        # Also set progress to 100 if installed
        titleItem = getText(xmlTitleListObj[item].childNodes).strip()
        appFileItem = getText(xmlAppFileListObj[item].childNodes).strip()
        urlItem = getText(xmlUrlListObj[item].childNodes).strip()
        versionItem = getText(xmlVersionListObj[item].childNodes).strip()

        # Check if item is installed and installed version
        # Get and set selectBox, iconPathMod, progressValue, installStateList
        (progressValue, iconPathMod, updateFlag) = processInstallVersion(titleItem, appFileItem, versionItem)

        # Build own arrays from xml
        titleList.append(titleItem)
        descriptionList.append(getText(xmlDescriptionListObj[item].childNodes))
        appFileList.append(appFileItem)
        urlList.append(urlItem)
        versionList.append(versionItem)
        # Build Progress blank array
        progressBox.append(progressValue)
        # Build Icon Path mod array
        iconPathList.append(iconPathMod)
        # Set True of update available
        updateList.append(updateFlag)

        # Build listStore select list array at the same time
        # using realCount to keep track of counters in the arrays
        # because not all items are added the count goes out
        realCount = realCount + 1

    appendToLog('ListStore Created')

    splashLabelText.set('All Done')
    splashProgressPercent.set(100)
    refreshGui(splashWindow)

    # Splash mainloop never called so mainWindow.mainloop is the main GUI thread
    # Kill Splash Screen
    splashWindow.destroy()

    # Render Main Window
    renderMainWindow()
    mainWindow.lift()
    # Start Mainwindow loop
    mainWindow.mainloop()

    # End
    sys.exit()