
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from serial_interface import arroyo
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from math import *
from pylab import *
from time import sleep
import pyqtgraph as pg
from tkinter import filedialog
from datetime import *
import threading
import random
from Powermeter import PW_1936r
import numpy as np
import os
import time

# DLL pour le powermeter
Powermeter = PW_1936r(LIBNAME=r'C:\Users\PC_labo_2\OneDrive - InSpek\Documents\Software\Labo optique\Interface_python_ordi_Dorian\interface_dorian\usbdll.dll', product_id=0xCEC7)


try :
    TECpak586 = arroyo()
    TECpak586.set_output_las(0)
    TECpak586.set_output(0)
except:
    TECpak586 = None
try:
    Powermeter.write('PM:AUTO 1')
except :
    pass

class Worker_signal(QObject):
    finished = pyqtSignal(list)
    progress = pyqtSignal(int)

class Worker_signal_laser(QObject):
    progress = pyqtSignal(list)

class Worker_laser(QThread):
    signal = Worker_signal_laser()

    def __init__(self, arroyo):
        super(Worker_laser, self).__init__()
        self.trigger = True
        self.arroyo = arroyo

    def run(self):
        self.trigger = True
        while self.trigger:
            time.sleep(0.5)
            self.signal.progress.emit([self.arroyo.read_output_Las(),
                                      self.arroyo.read_output_voltage(),self.arroyo.read_output_Photodiod_Current(), self.arroyo.read_temp(),
                                      self.arroyo.read_current()])

class Worker(QThread):
    
    
    def __init__(self,min,max,step,choice,Time):
        super(Worker,self).__init__()
        self.min = min
        self.max = max
        self.step = step
        self.choice = choice
        self.Counter = 0
        self.Time = Time
        self.signal = Worker_signal()
        

    def run(self):
        L= []
        mini = self.min
        maxi = self.max + self.step
        stepp = self.step
        M = np.arange(mini,maxi,stepp)
        time.sleep(2.0)
        for k in range(len(M)):
            if self.Counter == 1:
                return
            elif self.choice == "TEC":
                TECpak586.set_temp(self.min+k*self.step)
                sleep(self.Time)
                L.append(float(Powermeter.ask("PM:Power?").decode('UTF-8')))
            else:
                output = self.min + k * self.step
                print(output)
                TECpak586.set_las(output)
                sleep(self.Time)
                L.append(float(Powermeter.ask("PM:Power?").decode('UTF-8')))
            self.signal.progress.emit(k)
        self.signal.finished.emit(L)
        print('PASS')



class MyWindow(QMainWindow, QWidget):

    def __init__(self):
        super().__init__()
        self._createMenuBar()
        self._createCentralWidget()

    def _createCentralWidget(self):
        self.setWindowTitle('CONTROL UNIT')
        self.setCentralWidget(MyScene())

    def _createMenuBar(self):
        self.scene = MyScene()
        open = QAction("Open", self)
        open.setStatusTip("Open a file")
        open.triggered.connect(self.scene.load_spectrum)
        self.statusBar()
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')  # Creating menus using a QMenu object
        fileMenu.addAction(open)
        fileMenu.addAction("Saving Parameters")
        fileMenu.addAction("Exit")
        menubar.addMenu(fileMenu)
        editMenu = menubar.addMenu("&Edit")  # Creating menus using a title
        helpMenu = menubar.addMenu("&Help")
        self.setMenuBar(menubar)


    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Window Close', 'Are you sur you want to close the window ? ',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            TECpak586.set_output(0)
            TECpak586.set_output_las(0)
            TECpak586.close()
            Powermeter.close_device()
            event.accept()
        else:
            event.ignore()


class MyScene(QWidget, QGraphicsScene):

    def __init__(self):
        super(MyScene, self).__init__()
        self.Worker = None
        self.laser_TEC_status = "OFF"
        self.laser_current_status = "OFF"
        self.launched = "OFF"
        if not os.path.exists("Name.txt"):
            open("Name.txt", "w+")
        self.f = open(r"Name.txt")
        self.file = self.f.readline()
        self.M = []
        self.Auto = "ON"



        # Fonts
        serifFont = QFont("Overlock", 12, QFont.Bold)
        sansFont = QFont("Overlock", 10)
        sansFont.setBold(False)

        # Laser Box
        GroupBox_combo_source = QGroupBox('Laser Setup')
        GroupBox_combo_source.setFont(serifFont)
        GroupBox_combo_source.setMaximumWidth(600)
        GroupBox_combo_source.setMinimumWidth(600)

        # GroupBox_combo_source.addPixmap(QIcon('laser_icon.png'))

        GroupBox_laser_current = QGroupBox('Laser Current')
        GroupBox_laser_current.setFont(sansFont)
        GroupBox_laser_TEC = QGroupBox('Laser TEC')
        GroupBox_laser_TEC.setFont(sansFont)
        grid = QGridLayout()
        layout1 = QVBoxLayout()
        layout2 = QVBoxLayout()
        self.mywidgetlaser_current = MyWidgetLaser_current()
        self.mywidgetlaser_TEC = MyWidgetLaser_TEC()
        layout1.addWidget(self.mywidgetlaser_current)
        layout2.addWidget(self.mywidgetlaser_TEC)
        GroupBox_laser_current.setLayout(layout1)
        GroupBox_laser_TEC.setLayout(layout2)
        grid.addWidget(GroupBox_laser_current, 1, 1)
        grid.addWidget(GroupBox_laser_TEC, 1, 2)
        GroupBox_combo_source.setLayout(grid)
        GroupBox_combo_source.setMaximumHeight(400)
        GroupBox_combo_source.setMinimumWidth(600)



        # Plot widget
        GroupBox_plot = QGroupBox('Plot')
        GroupBox_plot.setFont(serifFont)
        layout = QVBoxLayout()
        self.mywidgetplot = MyPlottedSpectrum()
        layout.addWidget(self.mywidgetplot)
        GroupBox_plot.setLayout(layout)
        GroupBox_plot.setMinimumHeight(800)
        GroupBox_plot.setMaximumWidth(1000)

        # Plot widget

        GroupBox_savespectro = QGroupBox('Settings')
        GroupBox_savespectro.setFont(serifFont)
        GroupBox_savespectro.setMaximumHeight(1000)
        GroupBox_savespectro.setMaximumWidth(600)
        GroupBox_savespectro_button = QGroupBox()
        GroupBox_savespectro_button.setFont(sansFont)
        grid = QGridLayout()
        layout = QVBoxLayout()
        self.mywidgetspectro2 = MyPlottedSpectrum2()
        layout.addWidget(self.mywidgetspectro2)
        GroupBox_savespectro_button.setLayout(layout)
        grid.addWidget(GroupBox_savespectro_button, 1, 1)
        GroupBox_savespectro.setLayout(grid)
        # Lshow()

        # #Background grid
        # Back_grid = QGridLayout()
        # Back_grid.addWidget(GroupBox_combo_source,1, 1)
        # Back_grid.addWidget(GroupBox_spectro,2, 1)
        # Back_grid.addWidget(GroupBox_plot, 2, 2)

        mainLayout = QGridLayout()  # Create the main layout
        mainLayout.addWidget(GroupBox_combo_source, 0, 0, 1, 1)
        mainLayout.addWidget(GroupBox_plot, 0, 1, 0, 1)
        mainLayout.addWidget(GroupBox_savespectro, 1, 0, 1, 1)
        self.value = 20
        self.setLayout(mainLayout)
        self.mywidgetlaser_TEC.button_SET.clicked.connect(self.change_laser_TEC)
        self.mywidgetlaser_current.button_SET.clicked.connect(self.change_laser_current)
        self.mywidgetlaser_current.button_ON_OFF.clicked.connect(self.change_laser_current_status)
        self.mywidgetlaser_TEC.button_ON_OFF.clicked.connect(self.change_laser_TEC_status)
        self.mywidgetspectro2.path.clicked.connect(self.choose_path)
        self.mywidgetspectro2.launch_plotting.clicked.connect(self.Run)
        self.mywidgetspectro2.launch_plotting.setIcon(QIcon('led-red-on.png'))
        print(self.file)
        self.mywidgetspectro2.Text_path.setText(str(self.file))
        self.mywidgetspectro2.Text_File.setText("Name")
        self.mywidgetspectro2.Range.currentIndexChanged.connect(self.Change_Range)

    def Change_Range(self,value):
        if value != 0:
            if self.Auto == "ON":
                Powermeter.write('PM:AUTO 0')
                self.Auto = "OFF"
            Powermeter.write("PM:RAN " + str(value-1))
        else:
            Powermeter.write('PM:AUTO 1')
            self.Auto = "ON"

    def connect_arroyo(self,TECpak586):
        TECpak586 = arroyo()


    def Progress(self,value):
        self.mywidgetspectro2.progressBar.setValue(value)

    def Quit(self):
        self.Worker.signal.finished.disconnect(self.Plot_Axes)
        self.Worker.signal.progress.disconnect(self.Progress)
        self.Worker.quit()
        if not self.mywidgetspectro2.radiobutton.isChecked():
            self.change_laser_current_status()
            TECpak586.set_las(self.mywidgetlaser_current.MinValue.value())


    def Plot_Axes(self, reading):
        temp_results= np.array(reading)
        mask= temp_results <= 1 # avoid saturation peaks
        temp=self.M[mask]
        self.M = np.array(temp)
        filtered_results = temp_results[mask]
        self.mywidgetplot.p6.clear()
        color2 = (255,0,0)
        scatterplot = pg.PlotCurveItem(self.M, filtered_results, pen=color2)
        self.mywidgetplot.p6.addItem(scatterplot)
        self.mywidgetplot.p6.enableAutoRange()
        self.save_spectrum(filtered_results,self.M)
        self.mywidgetspectro2.launch_plotting.setIcon(QIcon('led-red-on.png'))
        self.launched = "OFF"
        self.Quit()


    def Run(self):
        if self.launched == "OFF":
            if self.mywidgetspectro2.radiobutton.isChecked():
                styles = {'font-size': '15px'}
                self.mywidgetplot.p6.setLabel('bottom', 'Temperature (C°)', **styles)
                Choice = "TEC"
                if self.laser_TEC_status == "OFF":
                    Tk().withdraw()
                    messagebox.showerror("message", "Laser Tec has to be ON")
                    return
                self.M = np.arange(self.mywidgetlaser_TEC.MinValue.value(),
                              self.mywidgetlaser_TEC.MaxValue.value() + self.mywidgetlaser_TEC.Step.value(),
                              self.mywidgetlaser_TEC.Step.value())
                self.Worker = Worker(self.mywidgetlaser_TEC.MinValue.value(), self.mywidgetlaser_TEC.MaxValue.value(),
                                     self.mywidgetlaser_TEC.Step.value(), Choice, self.mywidgetspectro2.Time.value())

            else:
                Choice = "Current"
                styles = {'font-size': '15px'}
                self.mywidgetplot.p6.setLabel('bottom', 'Current (mA)', **styles)
                if self.laser_current_status == "OFF":
                    Tk().withdraw()
                    messagebox.showerror("message", "Laser Current has to be ON")
                    return
                self.M = np.arange(self.mywidgetlaser_current.MinValue.value(),
                              self.mywidgetlaser_current.MaxValue.value() + self.mywidgetlaser_current.Step.value(),
                              self.mywidgetlaser_current.Step.value())
                self.Worker = Worker(self.mywidgetlaser_current.MinValue.value(), self.mywidgetlaser_current.MaxValue.value(),
                                     self.mywidgetlaser_current.Step.value(), Choice, self.mywidgetspectro2.Time.value())
            self.mywidgetspectro2.progressBar.setMaximum(len(self.M) - 1)
            self.mywidgetspectro2.progressBar.setValue(0)
            self.Worker.signal.progress.connect(self.Progress)

            self.Worker.signal.finished.connect(self.Plot_Axes)
            self.Worker.start()
            self.launched = "ON"
            self.mywidgetspectro2.launch_plotting.setIcon(QIcon('green-led-on.png'))
        else:
            self.Worker.Counter = 1
            self.Quit()
            self.mywidgetspectro2.launch_plotting.setIcon(QIcon('led-red-on.png'))
            self.launched = "OFF"

    def Delete_line(self):
        with open(r"Name.txt", 'r') as f:
            lines = f.readlines()
            f.close()

        # Write file
        with open(r"Name.txt", 'w') as f:
            for line in lines:
                f.write(line)
            f.close()

    def choose_path(self):
        self.Delete_line()
        path = QFileDialog.getExistingDirectory(self, "Select Directory")
        self.file = str(path)
        self.mywidgetspectro2.Text_path.setText(self.file)
        QApplication.processEvents()
        f = open("Name.txt", "w")
        f.write(self.file)
        f.close()



    def change_laser_current(self):
        current_value = self.mywidgetlaser_current.current_laser_box.value()
        TECpak586.set_las(current_value)

    def change_laser_TEC(self):

        current_value = self.mywidgetlaser_TEC.TEC_laser_box.value()
        TECpak586.set_temp(current_value)
        self.mywidgetlaser_TEC.UpdateTemp()

    def change_laser_TEC_status(self):
        if self.laser_TEC_status == "ON":
            TECpak586.set_output(0)
            self.laser_TEC_status = "OFF"
            self.mywidgetlaser_TEC.button_ON_OFF.setIcon(QIcon('led-red-on.png'))
            self.laser_worker.progress.disconnect(self.change_settings)
            self.laser_worker.quit()
        else:
            self.laser_TEC_status = "ON"
            TECpak586.set_output(1)
            self.laser_worker = Worker_laser(TECpak586)
            self.laser_worker.signal.progress.connect(self.change_settings)
            self.laser_worker.start()
            self.mywidgetlaser_TEC.button_ON_OFF.setIcon(QIcon('green-led-on.png'))

        print("Laser TEC status %s" % TECpak586.read_temp())
    def change_settings(self,list):
        self.mywidgetlaser_current.current.setText("%s mA" % list[0])
        self.mywidgetlaser_current.voltage.setText("%s V" % list[1])
        self.mywidgetlaser_current.current_PD.setText("%s μA" % list[2])
        self.mywidgetlaser_TEC.temp.setText("%s C°" % list[3])
        self.mywidgetlaser_TEC.current.setText("%s mA" % list[4])

    def change_laser_current_status(self):
        if self.laser_current_status == "ON":
            self.laser_current_status = "OFF"
            self.mywidgetlaser_current.button_ON_OFF.setIcon(QIcon('led-red-on.png'))
            TECpak586.set_output_las(0)
        else:
            if TECpak586.read_output() == 0:
                Tk().withdraw()
                messagebox.showerror("message", "Laser Tec has to be ON")
            else:
                self.laser_current_status = "ON"
                self.mywidgetlaser_current.button_ON_OFF.setIcon(QIcon('green-led-on.png'))
                TECpak586.set_output_las(1)

        print("Laser current status " + str(TECpak586.read_Las_limit))

    def save_spectrum(self, reading,M):
        now = datetime.now()
        file = open(self.mywidgetspectro2.Text_path.text() + "/ " + self.mywidgetspectro2.Text_File.text() + "-%d-0%d-0%d_%d-%d-%d" % (
        now.year, now.month, now.day, now.hour, now.minute,
        now.second) + ".txt", 'w')
        np_spectrum = np.c_[M, reading]
        np.savetxt(file, np_spectrum, fmt=['%.4f','%.10e'])
        file.close()


    def load_spectrum(self):
        Tk().withdraw()
        file = filedialog.askopenfile()
        if file is None:
            print('None')
            return
        else:
            print(file)
            a = np.loadtxt(file, dtype=np.float64)
            print(list(a[:,0]))
            self.mywidgetplot.p6.clear()
            self.mywidgetplot.p6.enableAutoRange()
            self.mywidgetplot.p6.plot(a[:, 0], a[:, 1])


class MyWidgetLaser_current(QWidget):
    def __init__(self):
        super().__init__()
        serifFont = QFont("Overlock", 12, QFont.Bold)

        fbox = QFormLayout()
        self.current_laser_box = QDoubleSpinBox(self)
        self.current_laser_box.setRange(0,510)
        self.current_laser_box.setMaximum(500)
        self.current_laser_box.setSingleStep(0.01)
        fbox.addRow(QLabel("Set Point (mA) :"), self.current_laser_box)
        self.current = QLabel("- mA")
        fbox.addRow(QLabel("Io = "), self.current)
        self.voltage = QLabel("- mV")
        fbox.addRow(QLabel("V0 = "), self.voltage)
        self.current_PD = QLabel("- μA")
        fbox.addRow(QLabel("IPD = "), self.current_PD)
        self.laser_status = "OFF"
        self.button_ON_OFF = QPushButton("ON/OFF")
        self.button_SET = QPushButton("Set Settings")
        self.button_ON_OFF.setIcon(QIcon('led-red-on.png'))
        fbox.addRow(self.button_ON_OFF)
        fbox.addRow(self.button_SET)
        self.MinValue = QDoubleSpinBox(self)
        self.MinValue.setRange(-20, 510)
        self.MinValue.setValue(90)
        self.MinValue.setSingleStep(0.01)
        fbox.addRow(QLabel("Set MinPoint (mA):"), self.MinValue)
        self.MaxValue = QDoubleSpinBox(self)
        self.MaxValue.setRange(-20, 510)
        self.MaxValue.setValue(100)
        self.MaxValue.setSingleStep(0.01)
        fbox.addRow(QLabel("Set MaxPoint (mA):"), self.MaxValue)
        self.Step = QDoubleSpinBox(self)
        self.Step.setRange(-20, 510)
        self.Step.setValue(1)
        self.Step.setSingleStep(0.001)
        fbox.addRow(QLabel("Set Step (mA):"), self.Step)
        self.setLayout(fbox)




class MyWidgetLaser_TEC(QWidget):
    def __init__(self, parent=None):
        super(MyWidgetLaser_TEC, self).__init__(parent=parent)
        fbox = QFormLayout()
        self.TEC_laser_box = QDoubleSpinBox(self)
        self.TEC_laser_box.setRange(-20, 40)
        self.TEC_laser_box.setValue(25)
        self.TEC_laser_box.setSingleStep(0.01)
        self.button_SET = QPushButton("Set Settings")
        fbox.addRow(QLabel("Set Point (°C):"), self.TEC_laser_box)
        self.temp = QLabel("- C°", self)
        fbox.addRow(QLabel("T = "), self.temp)
        self.current = QLabel("- mA", self)
        fbox.addRow(QLabel("I = "), self.current)
        self.laser_status = "OFF"
        self.button_ON_OFF = QPushButton("ON/OFF")
        self.button_ON_OFF.setIcon(QIcon('led-red-on.png'))
        fbox.addRow(self.button_ON_OFF)
        fbox.addRow(self.button_SET)
        self.MinValue = QDoubleSpinBox(self)
        self.MinValue.setRange(-20, 40)
        self.MinValue.setValue(20)
        self.MinValue.setSingleStep(0.01)
        fbox.addRow(QLabel("Set MinPoint (°C):"), self.MinValue)
        self.MaxValue = QDoubleSpinBox(self)
        self.MaxValue.setRange(-20, 40)
        self.MaxValue.setValue(25)
        self.MaxValue.setSingleStep(0.01)
        fbox.addRow(QLabel("Set MaxPoint (°C):"), self.MaxValue)
        self.Step = QDoubleSpinBox(self)
        self.Step.setDecimals(4)
        self.Step.setRange(-20, 510)
        self.Step.setValue(1.00)
        self.Step.setSingleStep(0.001)
        fbox.addRow(QLabel("Set Step (°C):"), self.Step)
        self.setLayout(fbox)




class MyPlottedSpectrum(pg.GraphicsLayoutWidget):

    def __init__(self):
        super(MyPlottedSpectrum, self).__init__()
        l = QVBoxLayout()
        self.pw = pg.PlotWidget(name='Plot')
        self.setBackground("w")
        l.addWidget(self.pw)

        pg.setConfigOptions(antialias=True)
        self.p6 = self.addPlot(title="")
        font = QFont()
        font.setPixelSize(1000)
        styles = {'font-size': '15px'}
        self.p6.setLabel('bottom', '', **styles)
        self.p6.setLabel('left', 'Power (W)', **styles)
        self.p6.enableAutoRange()
        pg.setConfigOptions(antialias=True)
        fn = QFont()
        fn.setPointSize(13)
        self.p6.getAxis("bottom").setTickFont(fn)
        self.p6.getAxis("left").setTickFont(fn)


class MyPlottedSpectrum2(QWidget):

    def __init__(self):
        super().__init__()
        fbox = QFormLayout()
        self.radiobutton = QRadioButton("Change TEC")
        fbox.addRow(self.radiobutton)
        self.radiobutton.setChecked(True)
        self.radiobutton_2 = QRadioButton("Change Current")
        fbox.addRow(self.radiobutton_2)
        self.launch_plotting = QPushButton('Launch Plotting')
        fbox.addRow(self.launch_plotting)
        self.Text_File = QLineEdit(self)
        fbox.addRow(QLabel("File Name :"),self.Text_File)
        self.Text_path = QLineEdit(self)
        fbox.addRow(QLabel("Path :"), self.Text_path)
        self.path = QPushButton('Choose Path')
        fbox.addRow(self.path)
        self.Time = QDoubleSpinBox(self)
        self.Time.setRange(-20, 100)
        self.Time.setValue(1)
        self.Time.setSingleStep(0.01)
        fbox.addRow(QLabel("Set Time sleep (sec):"), self.Time)
        self.progressBar = QProgressBar()
        fbox.addRow(self.progressBar)
        self.Range = QComboBox(self)
        self.Range_List = [self.tr("Auto Range"),self.tr('5.116 nW'), self.tr('51.16 nW'),self.tr('511.6 nW'),self.tr('5.116 μW'), self.tr('51.16 μW'),self.tr('511.6 μW'),self.tr('5.116 mW'), self.tr('51.16 mW')]
        self.Range.addItems(self.Range_List)
        fbox.addRow(QLabel("Range :"),self.Range)
        self.setLayout(fbox)



def main():
    app = QApplication(sys.argv)
    fenetre = MyWindow()
    fenetre.show()

    app.exec()


if __name__ == '__main__':
    main()