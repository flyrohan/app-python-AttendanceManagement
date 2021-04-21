# install modules
# $> pip install PyQt5
# $> pip install pyqt5-tools
# $> pip install pandas
# $> pip install pyqt-led
# $> pyinstaller --onedir --clean --hidden-import=PyQt5.sip <file>.py

import os
import sys
from typing import IO, Tuple
import pandas
import numpy as np
import json
import subprocess
import requests
import datetime
import urllib.parse as urlparse
import time
import re
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QThread, pyqtSignal, QMutex, pyqtSlot, QWaitCondition
from enum import Enum

_DEBUG_ = 1
_FILE_WINDOW_UI_ = "window.ui"
_FILE_TARGET_ = '구글 DOCS'
_FILE_SCHEDULE_ = "schedule.json"
_FILE_HOLIDAY_ = "holiday.json"
_FILE_REPORT_ = "report.txt"
_FILE_PATTERN_OVER = '^\d\d\d\d\d\d\d\d_근태기록_over.xlsx'
_FILE_PATTERN = [ '^\d\d\d\d\d\d\d\d_근태기록.xlsx', _FILE_PATTERN_OVER ]

window_ui = uic.loadUiType(os.path.join(os.path.dirname(os.path.abspath( __file__ )), _FILE_WINDOW_UI_))[0]

class DAY_TYPE(Enum):
    HOLIDAY = 0
    DAYOFF = 0
    HALFDAYOFF =4
    FAMILYDAY  = 4
    WORKDAY = 9
    INVALID = 10

class SCHEDULE_QUALTER(Enum):
    Q1 = 1
    Q2 = 4
    Q3 = 7
    Q4 = 10

def errorBox(message):
    if not message:
        return
    msg = QMessageBox()
    msg.setStandardButtons(QMessageBox.Ok)
    msg.setWindowTitle("Error!")
    msg.setText(message)
    msg.exec_()

# <Y><M><D> : ex. 20200101
def REPLACE_DATETIME(date):
    return date.replace(" ", "").replace("-", "").replace(":", "").replace(",", "")

def FILE_DATE(file):
    return os.path.basename(file).split("_")[0]

class ViewCheckedItem():
    def __init__(self, viewlist):
        self.listwidget = viewlist
        self.listwidget.itemClicked.connect(self.check_item)
        self.items = list()

    def check_item(self, item):
        if item.checkState() == Qt.Checked:
            self.items.append(item.text())
        else:
            try: self.items.remove(item.text())
            except ValueError: pass
        if _DEBUG_: print("SELECT", self.items)

    def get_checked_item(self):
        return self.items

    def add(self, item):
        w_item = QListWidgetItem()
        w_item.setText(item)
        w_item.setFlags(w_item.flags() | Qt.ItemIsUserCheckable)
        w_item.setCheckState(Qt.Unchecked)
        self.listwidget.addItem(w_item)

    def check_all(self, check):
        for i in range(self.listwidget.count()):
            item = self.listwidget.item(i)
            if check == True:
                item.setCheckState(Qt.Checked)
                self.items.append(item.text())
            else:
                item.setCheckState(Qt.Unchecked)
                self.items.remove(item.text())

    def clear(self):
        self.listwidget.clear()


class AttendSchedule:
    _SCHEDULE_NAME = 'Name'
    _SCHEDULE_HOLIDAY = 'Holiday'
    _SCHEDULE_FAMAILYDAY = 'Familyay'
    _SCHEDULE_DAYOFF = 'Dayoff'
    _SCHEDULE_HALFOFF = 'Halfoff'
    def __init__(self):
        self.__url = 'http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService'
        self.__operation = 'getHoliDeInfo' # 국경일 + 공휴일 정보 조회 오퍼레이션
        self.__url_key = '1CcWl1yYPdl0%2BTWYFzESRBK4dlyXZ%2Fc0qaQWjgtk3MRQgQ4RObrMQi5dOAacilxjc7VmpLjgj7OxKQbqgz65kQ%3D%3D'
        self.__holidayfile = ""
        self.__json = None
        self.__ko_holiday = dict()
        self.__holiday = list()
        self.__familyday = list()
        self.__names = list()

    def load_schedule(self, file):
        try:
            file = os.path.join(os.path.dirname(os.path.abspath( __file__ )), file)
            with open(file) as f:
                self.__json = json.load(f)
        except:
            errorBox("Check FILE {f} format !!!<br><br>$> cat {f} | python -m json.tool".format(f=file))
            return

        def __get(x):
            return (lambda x: sorted(list(set(self.__json.get(x)))) if self.__json.get(x) else [])(x)

        self.__holiday = __get(self._SCHEDULE_HOLIDAY)
        self.__familyday = __get(self._SCHEDULE_FAMAILYDAY)
        self.__names = __get(self._SCHEDULE_NAME)

    def load_holiday(self, file):
        self.__holidayfile = file
        try:
            file = os.path.join(os.path.dirname(os.path.abspath( __file__ )), file)
            with open(file) as f:
                self.__ko_holiday = json.load(f)
        except ValueError:
            errorBox("FILE {f} format !!!<br><br>$> cat {f} | python -m json.tool".format(f=file))
        except IOError:
            try:
                open(file, 'w+').close()
            except IOError as err:
                errorBox("FILE {f}, ERR: {e}".format(f=file, e=str(err)))
                pass

    def get_url_key(self):
            return self.__url_key

    def set_url_key(self, key):
            self.__url_key = key

    def get_names(self):
        if not self.__json or not self.__names:
            return None
        return self.__names

    def update_holiday(self, date):
        year = date[:4]; month = date[4:6]; day = date[6:8]; key = year + month
        if self.__ko_holiday.get(key):
            return
        self.__ko_holiday[key] = list()
        param = urlparse.urlencode( {'solYear':str(year), 'solMonth':str(month) })
        URL = self.__url + '/' + self.__operation + '?' + param + '&' + 'serviceKey' + '=' + self.__url_key + "&_type=json"
        resp = requests.get(url=URL)
        if True == resp.ok:
            items = resp.json().get('response').get('body').get('items')
            if items:
                item = items.get('item')
                if list == type(item):
                    ld = list(map(lambda x : x.get("locdate"), item))
                else:
                    ld = [item.get("locdate")]

                self.__ko_holiday[key] = list(map(str, ld))
                if not self.__holidayfile:
                    return
                try:
                    file = self.__holidayfile
                    json.dump(self.__ko_holiday, open(file, "w+"), indent=4, ensure_ascii=False)
                    if _DEBUG_: print("update holiday: ", date)
                except ValueError:
                    errorBox("FILE {f} format !!!<br><br>$> cat {f} | python -m json.tool".format(f=file))

    def get_day_type(self, name, date):
        date = REPLACE_DATETIME(date)
        if self.__json.get(name) and date in self.__json.get(name)[self._SCHEDULE_HALFOFF]:
            return DAY_TYPE.HALFDAYOFF
        if self.__json.get(name) and date in self.__json.get(name)[self._SCHEDULE_DAYOFF]:
            return DAY_TYPE.DAYOFF
        if date in self.__familyday:
            return DAY_TYPE.FAMILYDAY
        if datetime.date(int(date[:4]), int(date[4:6]), int(date[6:8])).weekday() > 4:
            return DAY_TYPE.HOLIDAY
        if date in self.__holiday:
            return DAY_TYPE.HOLIDAY
        if date[:6] in self.__ko_holiday and date in self.__ko_holiday[date[:6]]:
            return DAY_TYPE.HOLIDAY
        return DAY_TYPE.WORKDAY

class AttendDateTime:
    _WORKING_DATE = '날짜'
    _WORKING_ON_HOUR = '출근시간'
    _WORKING_ON_MIN = '출근분'
    _WORKING_OUT_HOUR = '퇴근시간'
    _WORKING_OUT_MIN = '퇴근분'
    _WORKING_TIME = '근무시간'
    _WORKING_TYPE = '구분'
    _WORKING_STATUS = 'STATUS'
    def __init__(self, name):
        self.__name = name
        self.__worktimes = list()
        self.__err_msg = ''

    def __check_work_time(self, worktime, hour, type):
        date = worktime[self._WORKING_DATE]
        on = worktime[self._WORKING_ON_HOUR]
        week = ["월", "화", "수", "목", "금", "토", "일"]
        wd = datetime.date(int(date[:4]), int(date[4:6]), int(date[6:8])).weekday()
        if on >= 10 or \
         ((hour < 9 and type != DAY_TYPE.HOLIDAY and type != DAY_TYPE.DAYOFF) and \
          (hour < 4 and type != DAY_TYPE.FAMILYDAY and type != DAY_TYPE.HALFDAYOFF)):
            self.__err_msg = "{n}: {d} ({w}), Work {h} hour {s} on {o} ".format(\
                n=self.__name, d=date, w=week[wd], h=int(hour), s=type.name,  o=on)
            worktime[self._WORKING_STATUS] = False

    def __exist(self, file):
        date = REPLACE_DATETIME(FILE_DATE(file))
        data = next((i for i in self.__worktimes if i[self._WORKING_DATE] == date), None)
        if data:
            return True
        return False

    def get_name(self):
        return self.__name

    def error_message(self):
        return self.__err_msg

    def append_date(self, file, date, ontime, outtime, schedule):
        if self.__exist(file):
            return
        date = REPLACE_DATETIME((lambda f, d: d if d else FILE_DATE(f))(file, date))
        year = int(date[:4]); month = int(date[4:6]); day = int(date[6:8])
        onhour = (lambda x: int(x.split(':')[0]) if x else 0)(ontime)
        onmin = (lambda x: int(x.split(':')[1]) if x else 0)(ontime)
        outhour = (lambda x: int(x.split(':')[0]) if x else 0)(outtime)
        outmin = (lambda x: int(x.split(':')[1]) if x else 0)(outtime)
        delta = datetime.datetime(year, month, day, outhour, outmin) - datetime.datetime(year, month, day, onhour, onmin)
        type = schedule.get_day_type(self.__name, date)
        worktime = {
            self._WORKING_DATE : date,
            self._WORKING_ON_HOUR : onhour,
            self._WORKING_ON_MIN : onmin,
            self._WORKING_OUT_HOUR : outhour,
            self._WORKING_OUT_MIN : outmin,
            self._WORKING_TIME : str(delta),
            self._WORKING_TYPE : type.name,
            self._WORKING_STATUS : True
        }
        self.__check_work_time(worktime, delta.seconds/3600, type)
        self.__worktimes.append(worktime)
        self.__worktimes = sorted(self.__worktimes, key=(lambda x: x[self._WORKING_DATE]))
        if _DEBUG_:
            print(self.__name, date, ontime, outtime, "[", type.name, ":", delta, "]", worktime[self._WORKING_STATUS])

    def remove_date(self, date):
        data = next((i for i in self.__worktimes if i[self._WORKING_DATE] == date), None)
        if data:
            self.__worktimes.remove(data)

    def get_work_times(self):
        return self.__worktimes


class AttendFileData():
    _EXCEL_NAME = '이름'
    _EXCEL_NAME_EXTEND = '(카드)'
    _EXCEL_COLUM_TIME = '발생시각'
    def __init__(self):
        self.__frame = list() # [ {"file" : file, "dataframe" : dataframe, "names" : names }, ...]
        self.__dataframe = None
        self.__file = ''
        self.__err_msg = ''

    def error_message(self):
        return self.__err_msg

    def append(self, file):
        try:
            dataframe = pandas.read_excel(file)
        except Exception as err:
            self.__err_msg = "FILE {f}, ERR: {e}".format(f=file, e=str(err))
            return False
        self.__file = file            
        self.__err_msg = ''
        # delete row with missing values
        dataframe.dropna(subset=[self._EXCEL_NAME], how='all', inplace = True)
        self.__frame.append({
                            "file" : file,
                            "dataframe" : dataframe,
                            "names" : list(set(dataframe[self._EXCEL_NAME].tolist()))
                            })
        # merge over dataframe
        self.__dataframe = pandas.concat([self.__dataframe, dataframe])
        self.__dataframe.drop_duplicates()
        self.__dataframe = self.__dataframe.sort_values(by = self._EXCEL_COLUM_TIME, ascending = True)

        # remove over-date dataframe
        date = FILE_DATE(file)
        date = datetime.datetime.strptime(date, '%Y%m%d') + datetime.timedelta(days=1)
        dataframe = pandas.to_datetime(self.__dataframe[self._EXCEL_COLUM_TIME], format="%Y-%m-%d %H:%M:%S")
        index = self.__dataframe[dataframe > date].index.tolist()
        self.__dataframe = self.__dataframe.drop(index=index)
        return True

    def remove(self, file):
        dataframe = next((i for i in self.__frame if i["file"] == file), None)
        if not dataframe:
            return

        self.__frame.remove(dataframe)
        if 0 == len(self.__frame):
            return

        # restructure dataframe
        if 1 == len(self.__frame):
            self.__dataframe = self.__frame[0]["dataframe"]
        else:
            for i in range(1, len(self.__frame)):
                self.__dataframe = pandas.concat([self.__frame[i - 1]["dataframe"], self.__frame[i]["dataframe"]])
                self.__dataframe.drop_duplicates()

        # remove over date
        self.__dataframe = self.__dataframe.sort_values(by = self._EXCEL_COLUM_TIME, ascending = True)
        date = FILE_DATE(file)
        date = datetime.datetime.strptime(date, '%Y%m%d') + datetime.timedelta(days=1)
        dataframe = pandas.to_datetime(self.__dataframe[self._EXCEL_COLUM_TIME], format="%Y-%m-%d %H:%M:%S")
        index = self.__dataframe[dataframe > date].index.tolist()
        self.__dataframe = self.__dataframe.drop(index=index)

    def get_files(self):
        files = list()
        for i in self.__frame:
            files.append(i['file'])
        return sorted(files)

    def get_names(self):
        names = list()
        for i in self.__frame:
           names.extend(i["names"])
        return sorted(names)

    def set_attendance(self, attendance: AttendDateTime, schedule: AttendSchedule):
        # filter with name
        name = attendance.get_name()
        dataframe = self.__dataframe
        df = (dataframe[self._EXCEL_NAME] == name) | \
             (dataframe[self._EXCEL_NAME] == name + self._EXCEL_NAME_EXTEND)
        df = dataframe[df]
        li = df[self._EXCEL_COLUM_TIME].tolist()
        li.sort()
        if not li:
            attendance.append_date(self.__file, None, None, None, schedule)
            return
        li[0], li[-1] = map(lambda x: str(x) if type(x) is pandas._libs.tslibs.timestamps.Timestamp else x, [li[0], li[-1]])
        date = li[0].split(' ')[0]
        on = li[0].split(' ')[1]
        out = li[-1].split(' ')[1]
        attendance.append_date(self.__file, date, on, out, schedule)

class AttendParser():
    def __init__(self):
        self.__filedata = {} # { "<file name>": 'AttendFileData', ... }
        self.__err_msg = ''
        self.__quarter = 0

    def error_message(self):
        return self.__err_msg

    def check_file_type(self, file: str):
        pattern = r'|'.join(_FILE_PATTERN)
        file = os.path.basename(file)
        rep = re.compile(pattern)
        if not rep.search(file):
            self.__err_msg = "FILE: {f} <br> - Invalid format {p}".format(f=file, p=pattern)
            return False
        if self.__quarter:
            mon = int(FILE_DATE(file)[4:6])
            print(self.__quarter, mon, file)
            if mon < (self.__quarter) or mon > (self.__quarter + 2):
                self.__err_msg = "FILE: {f} <br> - Not in {q} Quarter !".format(f=file, q=(int((self.__quarter+2)/3)))
                return False
        self.__err_msg = ''
        return True

    def set_quarter(self, quarter: int):
        self.__quarter = quarter

    def append_file(self, file: str):
        if self.check_file_type(file) == False:
            return False
        if file in self.get_files():
            return True
        if _DEBUG_: print("+: ", file)

        # create AttendFileData instance
        if not file in self.__filedata:
            # merge "over.xlsx"
            rep = re.compile(_FILE_PATTERN_OVER)
            if rep.search( os.path.basename(file)):
                merged = False
                for i in self.__filedata.keys():
                    # comapre <date>_근태기록_over with <date>_근태기록
                    if FILE_DATE(file) in i:
                        if self.__filedata[i].append(file) == False:
                            self.__err_msg = self.__filedata[i].error_message()
                            return False
                        merged = True
                if not merged:
                    self.__err_msg = "Not found base file for over: {f}".format(f=file)
                    return False
            else:
                filedata = AttendFileData()
                if filedata.append(file) == False:
                    self.__err_msg = filedata.error_message()
                    return False
                self.__filedata.update({ file : filedata })
        return True

    def remove_file(self, file: str):
        if _DEBUG_: print("-: ", file)
        if file in self.__filedata:
            del self.__filedata[file]
        else:
            for dataframe in self.__filedata.values():
                dataframe.remove(file)

    def get_files(self):
        files = list()
        for dataframe in self.__filedata.values():
            files.extend(dataframe.get_files())
        return sorted(files)

    def get_names(self):
        names = list()
        for dataframe in self.__filedata.values():
            names.extend(dataframe.get_names())
            names = list(set(names))
        return sorted(names)

    def set_attendance(self, attendance: AttendDateTime, schedule: AttendSchedule):
        for filedata in self.__filedata.values():
            filedata.set_attendance(attendance, schedule)


class ParserThread(QThread):
    parseProgress = pyqtSignal(int)
    parseUpdate = pyqtSignal(bool)
    errorMessage = pyqtSignal(str)
    def __init__(self, attend_file, inFiles, atList, parent=None):
        super().__init__(parent)
        self.mutex = QMutex()
        self.atFile = attend_file
        self.inFiles = inFiles
        self.atList = atList

    def run(self):
        self.mutex.lock()
        percent = 100 / len(self.inFiles); count = -1
        update = False
        for file in self.inFiles:
            if file not in self.atFile.get_files():
                if self.atFile.append_file(file) == False:
                    self.errorMessage.emit(self.atFile.error_message())
                    self.mutex.unlock()
                    return
                update = True
            count += 1
            self.parseProgress.emit(int(percent * count))

        for file in self.atFile.get_files():
            if file not in self.inFiles:
                self.atFile.remove_file(file)
                for at in self.atList:
                    at.remove_date(FILE_DATE(file))
                update = True

        self.parseProgress.emit(100)
        self.parseUpdate.emit(update)
        self.mutex.unlock()


class MainWindow(QMainWindow, window_ui):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("근태관리")
        self.setFixedSize(840, 660)
        self.setSizePolicy(QSizePolicy.Fixed , QSizePolicy.Fixed)

        self.btn_toFile.clicked.connect(self.on_button_tofile)
        self.btn_inFile.clicked.connect(self.on_button_infile)
        self.btn_run.clicked.connect(self.on_button_run)
        self.edit_toFile.textChanged.connect(self.do_edit_tofile)
        self.edit_inFiles.textChanged.connect(self.do_edit_infile)
        self.chb_selectall.stateChanged.connect(self.do_select_all)
        self.edit_schedulekey.textChanged.connect(self.do_edit_schedulekey)
        self.rbt_1q.clicked.connect(self.do_button_quarter)
        self.rbt_2q.clicked.connect(self.do_button_quarter)
        self.rbt_3q.clicked.connect(self.do_button_quarter)
        self.rbt_4q.clicked.connect(self.do_button_quarter)
        self.statusbar.addPermanentWidget(self.label_status, 1)
        self.statusbar.addPermanentWidget(self.progressbar, 2)
        self.result_txt = QTextEdit()
        self.names = ViewCheckedItem(self.listwidget_name)
        self.atFile = AttendParser()
        self.schedule = AttendSchedule()
        self.mutex = QMutex()

        self.schedule.load_schedule(_FILE_SCHEDULE_)
        self.schedule.load_holiday(_FILE_HOLIDAY_)

        self.toFile = ''
        self.inFiles = list()
        self.atList = list()
        self.status = False
        self.err_msg = False
        self.quarter = None

        self.btn_inFile.setDisabled(True)
        self.edit_inFiles.setDisabled(True)
        self.result_txt.setWindowTitle("Attendance Time")
        self.edit_schedulekey.setText(self.schedule.get_url_key())
        self.statusbar.setStyleSheet("color: blue;"
                             "background-color: lightgreen;"
                             "font-weight:bold"
                             )
        self.update_name()

    def update_statusbar(self, message, success):
        if success == True:
            self.statusbar.setStyleSheet("color: blue;" "background-color: lightgreen;" "font-weight:bold")
        else:
            self.statusbar.setStyleSheet("color: blue;" "background-color: red;" "font-weight:bold")
        self.label_status.setText(message)

    def on_button_tofile(self):
        self.toFile = QFileDialog.getOpenFileName(self)
        self.edit_toFile.setText(self.toFile[0])

    def on_button_infile(self):
        inFiles = QFileDialog.getOpenFileNames(self, 'Open file','./', ("*.*"))
        self.edit_inFiles.appendPlainText('\n'.join(inFiles[0]))

    def do_button_quarter(self):
        if self.rbt_1q.isChecked():
            quarter = SCHEDULE_QUALTER.Q1.value
        elif self.rbt_2q.isChecked():
            quarter = SCHEDULE_QUALTER.Q2.value
        elif self.rbt_3q.isChecked():
            quarter = SCHEDULE_QUALTER.Q3.value
        else:
            quarter = SCHEDULE_QUALTER.Q4.value

        if self.quarter != quarter:
            for file in self.atFile.get_files():
                self.atFile.remove_file(file)
                for at in self.atList:
                    at.remove_date(FILE_DATE(file))
            self.update_name()
            self.edit_inFiles.clear()

        self.quarter = quarter
        self.atFile.set_quarter(self.quarter)
        self.edit_inFiles.setDisabled(False)
        self.btn_inFile.setDisabled(False)

    def do_edit_tofile(self):
        self.toFile = self.edit_toFile.text()

    def do_edit_infile(self):
        self.inFiles = self.edit_inFiles.toPlainText().split()
        # check duplicated
        if len(self.inFiles) != len(set(self.inFiles)):
            files = dict()
            message = ""
            for i in self.inFiles:
                if i not in files.keys():
                    files[i] = 1
                else:
                    files[i] += 1
            for i in files.keys():
                if files[i] != 1:
                    message += "<br> - " + i
            self.status = False
            self.err_msg = "Duplicated: {x}".format(x=message)
            self.update_statusbar("FAILED: Parsing !", False)
            errorBox(self.err_msg)
            return
        self.edit_inFiles.setDisabled(True)
        self.parse_file()

    def do_edit_schedulekey(self):
        self.schedule.set_url_key(self.edit_schedulekey.text())

    def do_select_all(self):
        self.names.check_all(self.chb_selectall.isChecked())

    def on_button_run(self):
        select = self.names.get_checked_item()
        if not self.inFiles: errorBox("No input files !"); return
        if not select: errorBox("Not selected name !"); return
        if not (self.status == True): errorBox(self.err_msg); return

        self.update_statusbar("RUN", True)
        for name in select:
            at = None
            for i in self.atList:
                if name == i.get_name():
                    at = i
                    break
            if not at:
                at = AttendDateTime(name)
                self.atList.append(at)
            self.atFile.set_attendance(at, self.schedule)

        self.report_error()
        self.update_statusbar("DONE: RUN", True)

    def update_name(self):
        names = (lambda x, y: x if x else y)(self.schedule.get_names(), self.atFile.get_names())
        self.names.clear()
        for i in names:
            self.names.add(i)

    def report_error(self):
        datetimes = dict()
        popup = False
        for i in self.atList:
            name = i.get_name()
            datetimes[name] = list()
            for n in i.get_work_times():
                if not (n[AttendDateTime._WORKING_STATUS] == True):
                    datetimes[name].append(n.copy())
                    popup = True

        try:
            file = os.path.join(os.path.dirname(os.path.abspath( __file__ )), _FILE_REPORT_)
            with open(file, "wt+", encoding='utf-8') as f:
                for k in datetimes.keys():
                    s = str(k) + ' : ' + '\n'
                    for i in datetimes[k]:
                        s = s + "\t" + json.dumps(i, ensure_ascii=False) + '\n'
                    f.write(s)
        except Exception as err:
            errorBox("ERROR: {f} <br> {e}".format(f=file, e=str(err)))
            return

        if popup == True:
            with open(file, "rt", encoding='utf-8') as f:
                self.result_txt.setWindowTitle(file)
                self.result_txt.setText(f.read())
                self.result_txt.setReadOnly(True)
                self.result_txt.resize(1000, 500)
                self.result_txt.show()

    def signal_parse_done(self, result):
        if result == True:
            self.edit_inFiles.setPlainText('\n'.join(self.atFile.get_files()))
        if not self.schedule.get_names():
            self.update_name()

        for date in map(lambda x : FILE_DATE(x), self.atFile.get_files()):
            self.schedule.update_holiday(date)

        self.status = True
        self.update_statusbar("DONE: Parsing", True)
        self.edit_inFiles.setDisabled(False)

    def signal_parse_error(self, message):
        if message:
            errorBox(message)
        self.status = False
        self.err_msg = message
        self.update_statusbar("FAILED: Parsing !", False)
        self.edit_inFiles.setDisabled(False)

    def parse_file(self):
        if not self.inFiles:
            return
        self.update_statusbar("Parsing", True)
        parser = ParserThread(self.atFile, self.inFiles, self.atList, self)
        parser.parseProgress.connect(self.progressbar.setValue)
        parser.parseUpdate.connect(self.signal_parse_done)
        parser.errorMessage.connect(self.signal_parse_error)
        parser.start()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec_()
