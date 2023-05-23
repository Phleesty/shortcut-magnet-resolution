import os
import sys
import shutil
import win32com.client
import winreg # нужен для обнаружения пути до рабочего стола, если вдруг вы его перенесли
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QDesktopWidget, QFileDialog, QLineEdit


# Открываем ключ реестра Shell Folders и получаем путь до рабочего стола
with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
    myDesktop = winreg.QueryValueEx(key, "Desktop")[0]


#Создание папки shortcuts-magnet-resolution в AppData\Roaming и копирование в нее необходимой command-line utility NirCmd.exe
def get_nircmd_path():
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, "nircmd.exe")
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "nircmd.exe")

def nircmd_folder():
    src_path = get_nircmd_path()
    dest_folder = os.path.join(os.environ['APPDATA'], 'shortcuts-magnet-resolution')
    dest_path = os.path.join(dest_folder, 'nircmd.exe')

    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)

    if not os.path.exists(dest_path):
        shutil.copy(src_path, dest_path)
    return dest_path

app2 = QApplication(sys.argv)
screen = app2.primaryScreen()


desktop = QDesktopWidget()
game_path = ""
game_resolution_width = desktop.screenGeometry().width()
game_resolution_height = desktop.screenGeometry().height()
game_resolution_bpp = screen.depth()
game_resolution_freq = int(screen.refreshRate())
my_resolution_width = desktop.screenGeometry().width()
my_resolution_height = desktop.screenGeometry().height()
my_resolution_bpp = screen.depth()
my_resolution_freq = int(screen.refreshRate())
home_dir = os.path.expanduser("~")

class DragAndDropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        url = event.mimeData().urls()[0]
        file_path = url.toLocalFile()
        self.setText(file_path)
        global game_path
        game_path = file_path


class Ui_ShortcutCustomRes(object):
   

    def setupUi(self, ShortcutCustomRes):
        ShortcutCustomRes.setObjectName("ShortcutCustomRes")
        ShortcutCustomRes.resize(700, 160)
        ShortcutCustomRes.setMaximumSize(QtCore.QSize(16777215, 160))
        ShortcutCustomRes.setMinimumSize(QtCore.QSize(700, 160))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'magnet.ico')), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        ShortcutCustomRes.setWindowIcon(icon)
        ShortcutCustomRes.setStatusTip("")
        ShortcutCustomRes.setWhatsThis("")
        ShortcutCustomRes.setAccessibleName("")
        ShortcutCustomRes.setAccessibleDescription("")
        self.centralwidget = QtWidgets.QWidget(ShortcutCustomRes)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.object_game_path = DragAndDropLineEdit(self.centralwidget)
        self.object_game_path.setInputMask("")
        self.object_game_path.setDragEnabled(True)
        self.object_game_path.setPlaceholderText("")
        self.object_game_path.setObjectName("object_game_path")
        self.gridLayout.addWidget(self.object_game_path, 2, 1, 1, 7)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 1, 1, 1, 1)
        self.object_game_resolution_width = QtWidgets.QLineEdit(self.centralwidget)
        self.object_game_resolution_width.setInputMask("")
        self.object_game_resolution_width.setObjectName("object_game_resolution_width")
        self.gridLayout.addWidget(self.object_game_resolution_width, 3, 2, 1, 1)
        self.object_my_resolution_freq = QtWidgets.QLineEdit(self.centralwidget)
        self.object_my_resolution_freq.setObjectName("object_my_resolution_freq")
        self.gridLayout.addWidget(self.object_my_resolution_freq, 4, 6, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setObjectName("label_9")
        self.gridLayout.addWidget(self.label_9, 4, 7, 1, 2)
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setObjectName("label_8")
        self.gridLayout.addWidget(self.label_8, 4, 5, 1, 1)
        self.object_my_resolution_width = QtWidgets.QLineEdit(self.centralwidget)
        self.object_my_resolution_width.setObjectName("object_my_resolution_width")
        self.gridLayout.addWidget(self.object_my_resolution_width, 4, 2, 1, 1)
        self.object_game_resolution_bpp = QtWidgets.QLineEdit(self.centralwidget)
        self.object_game_resolution_bpp.setObjectName("object_game_resolution_bpp")
        self.gridLayout.addWidget(self.object_game_resolution_bpp, 3, 9, 1, 1)
        

        #space bars

        spacerItem_top = QtWidgets.QSpacerItem(20, 5, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.gridLayout.addItem(spacerItem_top, 0, 1, 1, 1)
        spacerItem_left = QtWidgets.QSpacerItem(10, 0, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem_left, 3, 0, 2, 1)    
        spacerItem_right = QtWidgets.QSpacerItem(10, 1, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem_right, 3, 10, 1, 1)
        

        
        spacerItem6 = QtWidgets.QSpacerItem(20, 26, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.gridLayout.addItem(spacerItem6, 3, 2, 1, 1)

        spacerItem_center = QtWidgets.QSpacerItem(20, 5, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.gridLayout.addItem(spacerItem_center, 5, 2, 1, 2)
        spacerItem_down = QtWidgets.QSpacerItem(20, 2, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        self.gridLayout.addItem(spacerItem_down, 7, 2, 1, 1)


        self.object_save_settings = QtWidgets.QPushButton(self.centralwidget)
        self.object_save_settings.setObjectName("object_save_settings")
        self.object_save_settings.clicked.connect(self.save_settings)
        self.gridLayout.addWidget(self.object_save_settings, 6, 1, 1, 3)


        
        #custom
        self.github_link = QtWidgets.QLabel(self.centralwidget)
        self.github_link.setAlignment(QtCore.Qt.AlignCenter)
        self.github_link.setOpenExternalLinks(True)
        self.github_link.setText('<a style="text-decoration: none; color: #5655b1;" href="https://github.com/phleesty/shortcuts-magnet-resolution">by Phleesty | GitHub</a>')
        self.github_link.setObjectName("github_link")
        self.gridLayout.addWidget(self.github_link, 6, 9, 1, 1)


        
        self.bat_radioButton = QtWidgets.QRadioButton(self.centralwidget)
        self.bat_radioButton.setObjectName("bat_radioButton")
        self.gridLayout.addWidget(self.bat_radioButton, 6, 4, 1, 4)
        self.bat_radioButton.toggled.connect(self.on_bat_radioButton_toggled)    


        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 3, 7, 1, 2)

        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setObjectName("label_7")
        self.gridLayout.addWidget(self.label_7, 4, 1, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 3, 3, 1, 1)
        self.object_game_resolution_freq = QtWidgets.QLineEdit(self.centralwidget)
        self.object_game_resolution_freq.setObjectName("object_game_resolution_freq")
        self.gridLayout.addWidget(self.object_game_resolution_freq, 3, 6, 1, 1)
        self.object_filedialog = QtWidgets.QPushButton(self.centralwidget)
        self.object_filedialog.setObjectName("object_filedialog")
        self.object_filedialog.clicked.connect(self.open_file_dialog)
        self.gridLayout.addWidget(self.object_filedialog, 2, 8, 1, 2)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 3, 5, 1, 1)
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 4, 3, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 3, 1, 1, 1)
        self.object_game_resolution_height = QtWidgets.QLineEdit(self.centralwidget)
        self.object_game_resolution_height.setClearButtonEnabled(False)
        self.object_game_resolution_height.setObjectName("object_game_resolution_height")
        self.gridLayout.addWidget(self.object_game_resolution_height, 3, 4, 1, 1)
        self.object_my_resolution_bpp = QtWidgets.QLineEdit(self.centralwidget)
        self.object_my_resolution_bpp.setObjectName("object_my_resolution_bpp")
        self.gridLayout.addWidget(self.object_my_resolution_bpp, 4, 9, 1, 1)
        self.object_my_resolution_height = QtWidgets.QLineEdit(self.centralwidget)
        self.object_my_resolution_height.setObjectName("object_my_resolution_height")
        self.gridLayout.addWidget(self.object_my_resolution_height, 4, 4, 1, 1)
        ShortcutCustomRes.setCentralWidget(self.centralwidget)





        self.retranslateUi(ShortcutCustomRes)
        QtCore.QMetaObject.connectSlotsByName(ShortcutCustomRes)


    def __init__(self, parent=None):
        self.create_shortcut = True

    def on_bat_radioButton_toggled(self, checked):
        self.create_shortcut = not checked

    def retranslateUi(self, ShortcutCustomRes):

        _translate = QtCore.QCoreApplication.translate
        ShortcutCustomRes.setWindowTitle(_translate("ShortcutCustomRes", "Shortcuts Magnet Resolution"))
        self.object_game_path.setPlaceholderText(_translate("ShortcutCustomRes", home_dir))

        #custom


        self.bat_radioButton.setText(_translate("ShortcutCustomRes", "Создать только .bat файл рядом с приложением"))

        #custom end


        self.label.setText(_translate("ShortcutCustomRes", "Путь до файла:"))
        
        self.label_9.setText(_translate("ShortcutCustomRes", "Глубина цвета: "))
        self.label_9.setToolTip(_translate("ShortcutCustomRes", "Оставьте значение по умолчанию, если вы не уверены. Для большинства современных мониторов значение равно 32 бита"))
        self.label_8.setText(_translate("ShortcutCustomRes", "Частота обновления:"))

        self.object_save_settings.setText(_translate("ShortcutCustomRes", "Сохранить настройки и создать ярлык"))
        self.label_5.setText(_translate("ShortcutCustomRes", "Глубина цвета: "))
        self.label_5.setToolTip(_translate("ShortcutCustomRes", "Оставьте значение по умолчанию, если вы не уверены. Для большинства современных мониторов значение равно 32 бита"))
        self.label_7.setText(_translate("ShortcutCustomRes", "Настройки рабочего стола:"))
        self.label_3.setText(_translate("ShortcutCustomRes", "x"))
        
        self.object_filedialog.setText(_translate("ShortcutCustomRes", "Обзор"))
        self.label_4.setText(_translate("ShortcutCustomRes", "Частота обновления:"))
        self.label_6.setText(_translate("ShortcutCustomRes", "x"))
        self.label_2.setText(_translate("ShortcutCustomRes", "Разрешение в приложении:"))
        
        self.object_game_path.setText(str(game_path))
        self.object_game_resolution_width.setText(str(game_resolution_width))
        self.object_my_resolution_freq.setText(str(my_resolution_freq))
        self.object_my_resolution_width.setText(str(my_resolution_width))
        self.object_game_resolution_bpp.setText(str(game_resolution_bpp))
        self.object_game_resolution_freq.setText(str(game_resolution_freq))
        self.object_game_resolution_height.setText(str(game_resolution_height))
        self.object_my_resolution_bpp.setText(str(my_resolution_bpp))
        self.object_my_resolution_height.setText(str(my_resolution_height))



        self.object_game_resolution_width.textChanged.connect(self.on_game_resolution_width_text_changed)
        self.object_game_resolution_height.textChanged.connect(self.on_game_resolution_height_text_changed)
        self.object_game_resolution_bpp.textChanged.connect(self.on_game_resolution_bpp_text_changed)
        self.object_game_resolution_freq.textChanged.connect(self.on_game_resolution_freq_text_changed)
        self.object_my_resolution_width.textChanged.connect(self.on_my_resolution_width_text_changed)
        self.object_my_resolution_height.textChanged.connect(self.on_my_resolution_height_text_changed)
        self.object_my_resolution_bpp.textChanged.connect(self.on_my_resolution_bpp_text_changed)
        self.object_my_resolution_freq.textChanged.connect(self.on_my_resolution_freq_text_changed)
        
    def on_game_resolution_width_text_changed(self, text):
        global game_resolution_width
        game_resolution_width = int(text)

    def on_game_resolution_height_text_changed(self, text):
        global game_resolution_height
        game_resolution_height = int(text)

    def on_game_resolution_bpp_text_changed(self, text):
        global game_resolution_bpp
        game_resolution_bpp = text

    def on_game_resolution_freq_text_changed(self, text):
        global game_resolution_freq
        game_resolution_freq = text

    def on_my_resolution_width_text_changed(self, text):
        global my_resolution_width
        my_resolution_width = int(text)

    def on_my_resolution_height_text_changed(self, text):
        global my_resolution_height
        my_resolution_height = int(text)

    def on_my_resolution_bpp_text_changed(self, text):
        global my_resolution_bpp
        my_resolution_bpp = text

    def on_my_resolution_freq_text_changed(self, text):
        global my_resolution_freq
        my_resolution_freq = text

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self.centralwidget, 'Выберите файл', '.', 'All files (*.*)')
        if file_path:
            self.object_game_path.setText(file_path)
            global game_path
            game_path = file_path


    # функция для сохранения параметров в файл настроек
    def save_settings(self):

            
            nircmd_path = nircmd_folder()
            game_exe = os.path.splitext(self.object_game_path.text())[0]
            lnk_name = os.path.splitext(os.path.basename(self.object_game_path.text()))[0]
            settings = f'@chcp 65001\n' \
                    f'@echo off\n' \
                    f'echo {lnk_name} with a resolution of {self.object_game_resolution_width.text()} x {self.object_game_resolution_height.text()} and a refresh rate of {self.object_game_resolution_freq.text()} Hz has been successfully launched. After closing the application or this window, your settings return to the values: {self.object_my_resolution_width.text()} x {self.object_my_resolution_height.text()}, {self.object_my_resolution_freq.text()} Hz \n' \
                    f'set "game={self.object_game_path.text()}"\n' \
                    f'set /a inGameWidth={self.object_game_resolution_width.text()}\n' \
                    f'set /a inGameHeight={self.object_game_resolution_height.text()}\n' \
                    f'set /a inGamebpp={self.object_game_resolution_bpp.text()}\n' \
                    f'set /a inGamefreq={self.object_game_resolution_freq.text()}\n' \
                    f'set /a myWidth={self.object_my_resolution_width.text()}\n' \
                    f'set /a myHeight={self.object_my_resolution_height.text()}\n' \
                    f'set /a myBpp={self.object_my_resolution_bpp.text()}\n' \
                    f'set /a myFreq={self.object_my_resolution_freq.text()}\n' \
                    f'"{nircmd_path}" setdisplay %inGameWidth% %inGameHeight% %inGamebpp% %inGamefreq%\n' \
                    f'start /wait "" "%game%"\n' \
                    f'"{nircmd_path}" setdisplay %myWidth% %myHeight% %myBpp% %myFreq%\n'
            settings_filename = f'{game_exe}.bat'
            
            with open(settings_filename, "w", encoding="utf-8") as f:
                f.write(settings)

            if self.create_shortcut:
                base_name = lnk_name
                shortcut_path = os.path.expanduser(f"{myDesktop}\{lnk_name}.lnk")
                target = os.path.realpath(settings_filename)
                wDir = os.path.dirname(target)
                icon = f'{self.object_game_path.text()}'
                shell = win32com.client.Dispatch("WScript.Shell")
                count = None

                while os.path.exists(shortcut_path):
                    if count is None:
                        count = 1
                    else:
                        count += 1
                    lnk_name = f"{base_name} {game_resolution_height}p {game_resolution_freq}Hz"
                    if count > 1:
                        lnk_name += f" {count}"
                    shortcut_path = os.path.expanduser(f"{myDesktop}\{lnk_name}.lnk")

                
            if self.bat_radioButton.isChecked():
                message = f"{settings_filename} с разрешением {self.object_game_resolution_width.text()} x {self.object_game_resolution_height.text()} и частотой обновления {self.object_game_resolution_freq.text()} Гц был успешно создан."
            else:
                message = f"Ярлык под названием {lnk_name} с разрешением {self.object_game_resolution_width.text()} x {self.object_game_resolution_height.text()} и частотой обновления {self.object_game_resolution_freq.text()} Гц был успешно создан на рабочем столе."
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target
                shortcut.WorkingDirectory = wDir

                # Проверяем существование файла иконки
                if os.path.exists(icon):
                    shortcut.IconLocation = icon
                else:
                    message += " Однако путь к файлу не содержит иконки, вы сможете заменить её сами в свойствах ярлыка."

                shortcut.save()



            # Создаем экземпляр QMessageBox
            msg_box = QtWidgets.QMessageBox()

            # Устанавливаем иконку для QMessageBox
            icon = QtGui.QIcon()
            icon.addPixmap(QtGui.QPixmap(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'magnet.ico')), QtGui.QIcon.Normal, QtGui.QIcon.Off)
            msg_box.setWindowIcon(icon)

            # Настраиваем свойства QMessageBox
            msg_box.setIcon(QtWidgets.QMessageBox.Information)
            msg_box.setWindowTitle("Ярлык создан")
            msg_box.setText(message)
            
            # Отображаем QMessageBox
            msg_box.exec_()

        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ShortcutCustomRes = QtWidgets.QMainWindow()
    
    ui = Ui_ShortcutCustomRes()
    ui.setupUi(ShortcutCustomRes)
    ShortcutCustomRes.show()
    sys.exit(app.exec_())
