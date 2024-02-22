from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt5 import uic
from PyQt5.QtGui import QPixmap
import warnings
import sys
from pedestrian import peatonal
from vehicle import vehicular
from gradient import gradient_analysis
from tiles import tile_report
from eyb import times_eyb
from order import order_files
from duplicate_excels import excels_duplicated
from duplicate_sheets import sheets_duplicated
from duplicate_vehicles import duplicated_carsandmc, duplicated_pedestrians

warnings.filterwarnings("ignore",category=DeprecationWarning)

class UI(QMainWindow):
    def __init__(self):
        super().__init__() #herencia
        uic.loadUi("./images/supervisor.ui",self) #traer la interfaz creada
        logo = QPixmap("./images/logo.jpg")
        self.label.setPixmap(logo)
        self.pushButton.clicked.connect(self.open_file)
        self.b_veh.clicked.connect(self.start_veh)
        self.b_pea.clicked.connect(self.start_ped)
        self.pushButton_Gradient.clicked.connect(self.start_gradient)
        self.pushButton_EyB.clicked.connect(self.start_EyB)
        self.pushButton_LC.clicked.connect(self.start_LC)
        self.pushButton_Order.clicked.connect(self.start_Order)
        self.pushButton_sheets.clicked.connect(self.start_sheets)
        self.pushButton_excels.clicked.connect(self.start_excels)
        self.pushButton_cars.clicked.connect(self.start_cars)
        self.pushButton_mc.clicked.connect(self.start_pedestrian)
        self.lineEdit.setText("")

    def open_file(self):
        self.entregable_path = QFileDialog.getExistingDirectory(self,"Seleccionar Entregable","c:\\")
        if self.entregable_path:
            self.lineEdit.setText(self.entregable_path)

    def start_veh(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        vehicular(self.entregable_path)
        self.label_4.setText("Finalizado")

    def start_ped(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        peatonal(self.entregable_path)
        self.label_5.setText("Finalizado")

    def start_gradient(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        gradient_analysis(self.entregable_path)
        self.label_6.setText("Finalizado")
    
    def start_LC(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        tile_report(self.entregable_path)
        self.label_7.setText("Finalizado")

    def start_EyB(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        times_eyb(self.entregable_path)
        self.label_8.setText("Finalizado")

    def start_Order(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        order_files(self.entregable_path)
        self.label_9.setText("Finalizado")

    def start_sheets(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        sheets_duplicated(self.entregable_path)
        self.label_10.setText("Finalizado")

    def start_excels(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        excels_duplicated(self.entregable_path)
        self.label_11.setText("Finalizado")

    def start_cars(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        length = self.spinBox.value()
        duplicated_carsandmc(self.entregable_path, length)
        self.label_12.setText("Finalizado")

    def start_pedestrian(self):
        if self.lineEdit.text() == "" or self.lineEdit.text() == "Selecciona un directorio":
            return self.lineEdit.setText("Selecciona un directorio")
        length = self.spinBox_2.value()
        duplicated_pedestrians(self.entregable_path, length)
        self.label_13.setText("Finalizado")

def main():
    app = QApplication(sys.argv)
    ventana = UI()
    ventana.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()