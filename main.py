import sys, re, pandas, statistics, scipy, math, statsmodels.api, copy, seaborn
from statistics import variance
from statistics import mean
from statistics import pstdev
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCore import *
from form.form import Ui_MainWindow as MainWindowSample
from form.form_reg import Ui_Dialog as RegWindowSample
from form.form_correl import Ui_Dialog as CorWindowSample
from form.form_stat import Ui_Dialog as StatWindowSample
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg

def seaborndisplot(data, key1, key2):
    return seaborn.relplot(data,x = key1, y = key2)


def seabornregplot(data, key1, key2):
    return seaborn.lmplot(x = key1, y = key2, data = data)

class MplCanvas(FigureCanvasQTAgg):
    def __init__(self, plot):
        self.fig = plot.figure
        self.axes = plot.axes
        super().__init__(self.fig)

def read_excel(file_name :str) -> list[list]:
    df = pandas.read_excel(file_name)
    data = [[row[col] for col in df.columns] for row in df.to_dict('records')]
    df = df.T
    data1 = [[row[col] for col in df.columns] for row in df.to_dict('records')]
    headers = pandas.read_excel(file_name).columns
    data_dict = {}
    for i in range(0,len(headers)):
        data_dict[headers[i]] = data1[i]
    return data, list(headers), data_dict

class StatWindow(QDialog):
    def __init__(self) -> None:
        super(StatWindow, self).__init__()
        self.ui = StatWindowSample()
        self.ui.setupUi(self)

class RegWindow(QDialog):
    def __init__(self) -> None:
        super(RegWindow, self).__init__()
        self.ui = RegWindowSample()
        self.ui.setupUi(self)


class CorWindow(QDialog):
    def __init__(self) -> None:
        super(CorWindow, self).__init__()
        self.ui = CorWindowSample()
        self.ui.setupUi(self)




class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super(MainWindow, self).__init__()
        self.ui = MainWindowSample()
        self.ui.setupUi(self)
        self.regWindow = RegWindow()
        self.corWindow = CorWindow()
        self.statWindow = StatWindow()
        self.data, self.headers,self.data_dict = read_excel('data/crimePrav.xlsx')
        self.set_dataset()
        self.checkboxes = [self.ui.checkBox,
            self.ui.checkBox_2,
            self.ui.checkBox_3,
            self.ui.checkBox_4,
            self.ui.checkBox_5,
            self.ui.checkBox_6,
            self.ui.checkBox_7,
            self.ui.checkBox_8,
            self.ui.checkBox_9,
            self.ui.checkBox_10,
            self.ui.checkBox_11,
            self.ui.checkBox_12,
            self.ui.checkBox_13,
            self.ui.checkBox_14]
        self.set_name_checkBoxes()
        self.set_name_comboBoxes()
        self.statistic_combobox()
        self.ui.correl_button.clicked.connect(self.exec_correlation)
        self.ui.regression_button.clicked.connect(self.exec_regression)
        self.ui.static_button.clicked.connect(self.exec_statistic)
        pass
    def exec_statistic(self):
        self.statistic_values()
        self.statWindow.resize(QSize(400, 300))
        self.statWindow.exec()

    def exec_regression(self):
        self.regression()
        self.regWindow.resize(QSize(900, 600))
        self.regWindow.exec()


    def exec_correlation(self):
        self.set_table_correl()
        self.corWindow.resize(QSize(900, 600))
        self.corWindow.ui.coef_correl_label.setText(str(self.correlation()))
        self.corWindow.exec()
    
    def set_name_checkBoxes(self):
        headers = self.headers[2:]
        for i in range(0,len(headers)):
            self.checkboxes[i].setText(headers[i])


    def set_name_comboBoxes(self):
        headers = self.headers[1:]
        for i in range(0,len(headers)):
            self.ui.col1_combobox.addItem(headers[i])
            self.ui.col2_combobox.addItem(headers[i])


    def correlation(self):  # Расчёт корреляции
        key1 = self.ui.col1_combobox.currentText()
        key2 = self.ui.col2_combobox.currentText()
        first_column = self.data_dict[key1]
        second_column = self.data_dict[key2]
        ResultList = [first_column, second_column]
        alpha = 0.05
        kor = statistics.correlation(ResultList[0], ResultList[1])
        sample_value = (kor * math.sqrt(len(ResultList[0]) - 2)) / math.sqrt(1 - (kor ** 2))
        theoretical_value = scipy.stats.t.isf(alpha / 2, len(ResultList[0]) - 2)
        if abs(sample_value) > theoretical_value:
            self.corWindow.ui.znam_correl_label.setText("Значим")
        else:  
            self.corWindow.ui.znam_correl_label.setText("Незначим")
        resultDict = {
            key1: ResultList[0],
            key2: ResultList[1]
        }
        removeWidget = self.corWindow.ui.horizontalLayout.takeAt(1)
        removeWidget.widget().setParent(None)
        self.corWindow.ui.horizontalLayout.removeWidget(removeWidget.widget())
        graphic = MplCanvas(plot=seaborndisplot(pandas.DataFrame(resultDict), key2, key1))
        self.corWindow.ui.horizontalLayout.addWidget(graphic)
        self.corWindow.ui.horizontalLayout.setStretch(1,1)
        return kor
    
    def set_table_correl(self):
        key1 = self.ui.col1_combobox.currentText()
        key2 = self.ui.col2_combobox.currentText()
        self.corWindow.ui.table_correl.setRowCount(len(self.data))
        self.corWindow.ui.table_correl.setColumnCount(2)
        self.corWindow.ui.table_correl.setHorizontalHeaderLabels([key1, key2])
        list_data = [self.data_dict[key1], self.data_dict[key2]]
        for i, row in enumerate(list_data):
            for j, item in enumerate(row):
                new_item = QTableWidgetItem(str(item))
                self.corWindow.ui.table_correl.setItem(j, i, new_item)

        header = self.corWindow.ui.table_correl.horizontalHeader()
        for i in range(0, header.count()):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        self.corWindow.ui.table_correl.setSortingEnabled(True)
    def set_regression_table(self, data_dict):
        self.list_checkBox_true = ["Total_crimes"]

        for i in self.checkboxes:
            if (i.isChecked() == 1):
                self.list_checkBox_true.append(i.text())
        self.regWindow.ui.table_reg_2.clear()
        self.regWindow.ui.table_reg_2.setRowCount(len(self.data))
        self.regWindow.ui.table_reg_2.setColumnCount(len(self.list_checkBox_true))
        self.regWindow.ui.table_reg_2.setHorizontalHeaderLabels(self.list_checkBox_true)
        list_data_reg = []

        for i in self.list_checkBox_true:
            list_data_reg.append(data_dict[i])

        for i, row in enumerate(list_data_reg):
            for j, item in enumerate(row):
                new_item = QTableWidgetItem(str(item))
                self.regWindow.ui.table_reg_2.setItem(j, i, new_item)

        header = self.regWindow.ui.table_reg_2.horizontalHeader()
        for i in range(0, header.count()):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        self.regWindow.ui.table_reg_2.setSortingEnabled(True)
    
    def regression(self):
        self.list_checkBox_true = ["Total_crimes"]
        for i in self.checkboxes:
            if (i.isChecked() == 1):
                self.list_checkBox_true.append(i.text())
        ResultList = []
        data_copy_dict = copy.deepcopy(self.data_dict)
        for i in self.list_checkBox_true:
            ResultList.append(data_copy_dict[i])
        past_length = 0
        while len(ResultList[0]) != past_length:
            past_length = len(ResultList[0])
            index = 0
            while index < len(ResultList):
                j = 0
                mid = mean(ResultList[index])
                standard = 3 * pstdev(ResultList[index])
                while j < len(ResultList[index]):
                    if ResultList[index][j] > mid + standard or ResultList[index][j] < mid - standard:
                        ResultList[index].pop(j)
                        for ind in range(len(ResultList)):
                            if ind != index:
                                ResultList[ind].pop(j)
                    j += 1
                index += 1
        list_data_regression = ResultList[1:]
        list_Y = ResultList[0]
        X = pandas.DataFrame(list_data_regression).T
        Y = pandas.DataFrame(list_Y)
        yx = pandas.DataFrame(self.data_dict)
        x1 = statsmodels.api.add_constant(X)
        reg = statsmodels.api.OLS(Y, x1).fit()
        function = "Y = " + str(round(reg.params["const"], 4))
        for i in range(len(reg.params) - 1):
            if reg.params[i] >= 0:
                function += " + "
            else:
                function += " - "
            function = function + str(abs(round(reg.params[i], 4))) + " * X" + str(i + 1)
        self.regWindow.ui.coef_reg_label_2.setText(function)


        if reg.f_pvalue > 0.05:
            f_test = str(reg.f_pvalue) + " (модель незначима)"
        else:
            f_test = str(reg.f_pvalue) + " (модель значима)"
        self.regWindow.ui.znam_mod_label_2.setText(f_test)


        keys = ["const"]
        if reg.pvalues["const"] > 0.05:
            values = [str(reg.pvalues["const"]) + " (коэффициент незначим)"]
        else:
            values = [str(reg.pvalues["const"]) + " (коэффициент значим)"]
        for i in range(len(reg.pvalues) - 1):
            keys.append("a" + str(i+1))
            if reg.pvalues[i] > 0.05:
                values.append(str((reg.pvalues[i])) + " (коэффициент незначим)")
            else:
                values.append(str((reg.pvalues[i])) + " (коэффициент значим)")

        t_test = {}
        for i in range(len(keys)):
            key = keys[i]
            value = values[i]
            t_test[key] = value

        self.regWindow.ui.znam_coef_label_2.setText('')
        self.regWindow.ui.label_2.setText(str(reg.rsquared))
        for key, value in t_test.items():
            key_items = (key, ':', value)
            self.regWindow.ui.znam_coef_label_2.setText(self.regWindow.ui.znam_coef_label_2.text() + str(key) + "  :  " + str(value) + "<br/>")
        if self.regWindow.ui.horizontalLayout_3.count() == 2:
            removeWidget = self.regWindow.ui.horizontalLayout_3.takeAt(1)
            removeWidget.widget().setParent(None)
            self.regWindow.ui.horizontalLayout_3.removeWidget(removeWidget.widget())
        if (len(self.list_checkBox_true) == 2):
            graphic = MplCanvas(plot=seabornregplot(pandas.DataFrame(yx), self.list_checkBox_true[1], self.list_checkBox_true[0]))
            self.regWindow.ui.horizontalLayout_3.addWidget(graphic)
            self.regWindow.ui.horizontalLayout_3.setStretch(1,0)
        self.set_regression_table(data_copy_dict)
    def statistic_combobox(self):
        headers = self.headers[1:]
        for i in headers:
            self.ui.comboBox.addItem(i)
    def statistic_values(self):
        column_name = self.ui.comboBox.currentText()
        lst = self.data_dict[column_name]
        self.statWindow.ui.count_label.setText(str(len(lst)))
        self.statWindow.ui.mean_label.setText(str(mean(lst)))
        self.statWindow.ui.dispersion_label.setText(str(variance(lst)))
        self.statWindow.ui.mean_quadratic_label.setText(str(pstdev(lst)))
        self.statWindow.ui.asimmetric_label.setText(str(scipy.stats.skew(lst, bias=False)))
        self.statWindow.ui.xx_label.setText(str(scipy.stats.kurtosis(lst, bias=False)))
        alpha = 0.05
        t = scipy.stats.t.isf(alpha / 2, len(lst) - 2)
        interval = [mean(lst) - t * (pstdev(lst) / (len(lst) ** 0.5)),
                    mean(lst) + t * (pstdev(lst) / (len(lst) ** 0.5))]
        self.statWindow.ui.interval_label.setText('(' + str(interval[0])+ ' ; ' + str (interval[1]) + ')')

    def set_dataset(self):
        self.ui.tableWidget.setRowCount(len(self.data))
        self.ui.tableWidget.setColumnCount(len(self.headers))
        self.ui.tableWidget.setHorizontalHeaderLabels(self.headers)

        for i, row in enumerate(self.data):
            for j, item in enumerate(row):
                new_item = QTableWidgetItem(str(item))
                self.ui.tableWidget.setItem(i, j, new_item)

        header = self.ui.tableWidget.horizontalHeader()
        for i in range(0, header.count()):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)

  


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())
