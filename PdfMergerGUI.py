from PyPDF2 import PdfReader, PdfFileWriter
import os, datetime, ast, threading
import pyautogui as pag

# Tkinter Components

from tkinter import *
from tkinter import messagebox
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import tkinter.filedialog as fd

class AISroller:
    def __init__(self):

        self.pdfWriter = PdfFileWriter()
        self.addPageCnt = 0

        self.tmpFilesArr = []
        self.addFilesArr = []
        self.finalFilesList = []

        self.tmpFilesIndex = 1
        self.addFilesIndex = 1

        self.outputFolder = "./OutputPDF"
        isExist = os.path.exists(self.outputFolder)

        if not isExist:
            os.mkdir(self.outputFolder)
        
        self.PDFMWind= tk.Tk()
        #self.scrWind.geometry("412x208+400+100")
        self.PDFMWind.geometry("502x280+400+100")
        self.PDFMWind.resizable(0,0)
        self.PDFMWind.overrideredirect(True)
        self.PDFMWind.title("PDF Merger")

        # bgColor = '#2176c7'
        winBgColor = '#000'
        self.PDFMWind.configure(bg=winBgColor)
        bgColor = 'yellow'#'#4B0082'
        labFgColor = '#fff'
        fgColor = '#000'
        repFgColor = '#fff'

        self.pdflabel = Label(self.PDFMWind, text = "PDF Merger", font=('Times New Roman',13, "bold"), cursor='fleur', height=2, bg=bgColor, fg=fgColor, padx=180)

        self.filesList = Listbox(self.PDFMWind, height = 4, width = 37, bg = "#000", activestyle = 'dotbox', font=('Courier New', 11), fg = "#fff")

        self.slctBtn = Button(self.PDFMWind, text= "Select", padx=15, command=self._selectFiles, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)    
        self.moveUp = Button(self.PDFMWind, text= "UP", padx=13, command=self._moveUpFile, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)
        self.moveDown = Button(self.PDFMWind, text= "Down", padx=10, command=self._moveDownFile, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)

        self.startBtn = Button(self.PDFMWind, text= "Start", padx=18, command=self._mergeActionThread, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)
        self.minBtn = Button(self.PDFMWind, text= "Minimize", padx=16, command=self._amsMinimizeWind, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)
        self.ctime = Label(self.PDFMWind, text="", font=("Times New Roman",10, "bold"), padx=15, bg=winBgColor, fg=labFgColor)
        self.abtLab = Label(self.PDFMWind, text = "Created By :: PS Thamizhan", font=('Times New Roman', 9, "bold"), height=1, bg=winBgColor, fg=labFgColor)
        self.exitBtn = Button(self.PDFMWind, text= "Exit", padx=14, command=self._closeTab, bd=0, bg=bgColor, activebackground='#fff', activeforeground='#000', fg=fgColor, pady=0)

        ######### Create Dir For Screenshots & Delete Previous Dirs #########

        
    def _closeTab(self):
        extBtn = messagebox.askquestion("PDF Merger Alert", "Are you sure to exit ?                                                  ")
        if(extBtn == 'yes'):
           self.PDFMWind.destroy()


    def mouse_down(self, event):
        global x, y
        x, y = event.x, event.y

    def mouse_up(self, event):
        global x, y
        x, y = None, None
    
    def _moveWindow(self, event):
        global x, y
        deltax = event.x - x
        deltay = event.y - y
        x0 = self.PDFMWind.winfo_x() + deltax
        y0 = self.PDFMWind.winfo_y() + deltay
        self.PDFMWind.geometry("+%s+%s" % (x0, y0))

    def _amsMinimizeWind(self):
        self.PDFMWind.state('withdrawn')
        self.PDFMWind.overrideredirect(False)
        self.PDFMWind.state('iconic')

    def _exitMinimizeWind(self, event=None):
        self.PDFMWind.overrideredirect(True)
        self.PDFMWind.state("normal")
       # self.PDFMWind.iconbitmap("./amsLogo.bmp")

    def _createFileds(self):

        self.PDFMWind.bind("<Map>", self._exitMinimizeWind)
        
        self.pdflabel.bind('<ButtonPress-1>', self.mouse_down)
        self.pdflabel.bind('<B1-Motion>', self._moveWindow)
        self.pdflabel.bind('<ButtonRelease-1>', self.mouse_up)

        self.pdflabel.place(x = 0,y = 5)        
        self.filesList.place(x= 8, y = 78)
        
        self.slctBtn.place(x = 8, y = 198)
        self.moveUp.place(x = 95, y = 198)
        self.moveDown.place(x = 154, y = 198)

        self.startBtn.place(x = 232, y = 198)
        self.minBtn.place(x = 317, y = 198)
        self._clock()

        self.abtLab.place(x = 8,y = 244)
        self.exitBtn.place(x = 431, y = 198)

        return True

    def _clock(self):
        now = datetime.datetime.now().strftime("%I:%M:%S %p")
        self.ctime.place(x = 362,y = 244)
        self.ctime.config(text=now)
        self.ctime.after(1000, self._clock)

    def _selectFiles(self):
        slctPdfFiles = fd.askopenfilenames(parent=self.PDFMWind, title='Choose a file')#tkFileDialog.askopenfilenames(mode ='r', )#, filetypes =[('PDF', '*.pdf')])
        if slctPdfFiles is not None:
            pdfFiles = self.PDFMWind.tk.splitlist(slctPdfFiles)

            self.addFilesArr.clear()
            self.tmpFilesArr.clear()

            self.tmpFilesIndex = 1
            self.addFilesIndex = 1
            
            for PdfPath in pdfFiles:
                
                tmpVal = {'id':str(self.tmpFilesIndex), 'path':str(PdfPath)}
                self.tmpFilesArr.append(tmpVal)
                self.tmpFilesIndex += 1

                PdfPath = PdfPath.split('/')
                countNames = len(PdfPath)
                getFileNameIndex = countNames-1
                #print(PdfPath[getFileNameIndex])
                
                tmpActArr = {'id':str(self.addFilesIndex), 'filename': str(PdfPath[getFileNameIndex])}
                self.addFilesArr.append(tmpActArr)
                self.addFilesIndex += 1


            if len(self.addFilesArr) > 0:
                
                self.filesList.delete(0,END)
                i = 1
   
                for names in self.addFilesArr:
                    tmpReader = ast.literal_eval(str(names))
                    #print(tmpReader)
                    self.filesList.insert(str(i), str(tmpReader['filename']))
                    i += 1
            
        else:
            print(0)

        #print(self.tmpFilesArr)
            
        return True

    def _moveUpFile(self):
        try:
            self.idxs = self.filesList.curselection()
            if not self.idxs:
                return
            for pos in self.idxs:
                if pos==0:
                    continue
                text=self.filesList.get(pos)
                self.filesList.delete(pos)
                self.filesList.insert(pos-1, text)
                self.filesList.pop(pos)
                self.filesList.insert(pos-1, text)
                self.filesList.select_set(pos-1)
        except:
            pass

        return True

    def _moveDownFile(self, *args):
        try:
            self.idxs = self.filesList.curselection()
            #print("Len :: "+str(self.filesList.size()))
            if not self.idxs:
                return
            for pos in self.idxs:
                if pos == self.filesList.size()-1:
                    continue
                text=self.filesList.get(pos)
                self.filesList.delete(pos)
                self.filesList.insert(pos+1, text)
                self.filesList.pop(pos)
                self.filesList.insert(pos+1, text)
            self.filesList.selection_set(pos + 1)
        except:
            pass

        return True

    def _mergeActionThread(self):
        self.offPDFList = self.filesList.get(0, tk.END)

        self.finalFilesList.clear()

        for filename in self.offPDFList:
            for tmpfiles in self.addFilesArr:
                astVal = ast.literal_eval(str(tmpfiles))

                if str(astVal['filename']) == str(filename):
                    getActId = astVal['id']

                    for filePath in self.tmpFilesArr:
                        pathAstVal = ast.literal_eval(str(filePath))
                        if str(pathAstVal['id']) == str(getActId):
                            fileActPath = pathAstVal['path']
                            self.finalFilesList.append(str(fileActPath))


        if len(self.finalFilesList) > 0:
            mergeActionThread = threading.Thread(target=self._mergeFiles, name='Start Merging Files')
            mergeActionThread.start()
        else:
            messagebox.showinfo("PDF Merger Alert", "Please select files                            ")
        
        return True

    def _mergeFiles(self):

        for filePath in self.finalFilesList:
            #filePath = os.getcwd()+'/Template Files/'+str(i)
            if filePath.endswith(".pdf"):
                readPDF = PdfReader(filePath)
                for page in readPDF.pages:
                    self.pdfWriter.addPage(page)
                    self.addPageCnt += 1

        if self.addPageCnt > 0:
            now = datetime.datetime.now().strftime("%I_%M_%S_%p")
            finalFileName = str(self.outputFolder)+"/mergedfile_"+str(now)+".pdf"
            with open(finalFileName, "wb") as outputStream:
                self.pdfWriter.write(outputStream)

            self.finalFilesList.clear()
            self.filesList.delete(0,END)
            messagebox.showinfo("PDF Merger Alert", "PDF Merged Successfully.\nFile Name :: "+str(finalFileName))
        else:
            messagebox.showinfo("PDF Merger Alert", "No Files")

        return True
        
    def run(self):
        self._createFileds()
        self.PDFMWind.mainloop()
        
runClass = AISroller()
runClass.run()
