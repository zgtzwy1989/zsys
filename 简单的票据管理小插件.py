import sys,openpyxl,datetime,os.path,pathlib
from PyQt5.QtWidgets import QApplication,QWidget,QLabel,QPushButton,QVBoxLayout,QHBoxLayout,QLineEdit,QInputDialog,QMessageBox,QMenuBar,QAction,QTableWidget,QHeaderView,QTableWidgetItem,QAbstractItemView

class Demo (QWidget):
    def __init__(self):
        super(Demo, self).__init__()
        self.resize(70,70)
        self.main()
        self.new_excel.clicked.connect(self.new1)#创建新的数据表
        self.new.clicked.connect(self.new2)#写入新的数据
        self.old_read.clicked.connect(self.new3)#读取数据
        self.old_write.clicked.connect(self.new4)
        #self.bar.triggered[QAction].connect(self.new_window)

    def main(self):
        self.new = QPushButton("新的数据", self)
        self.old_read = QPushButton("读取数据", self)
        self.new_excel = QPushButton("创建新的数据表", self)
        self.old_write = QPushButton("写入数据", self)
        self.one = QLabel("上一条数据", self)
        self.one_like = QLineEdit(self)
        self.write_data = QLabel("领取数值", self)
        self.write_data_like = QLineEdit(self)
        self.result = QLabel("结果", self)
        self.result_like = QLineEdit(self)
        self.result_bz = QLabel("领取人", self)
        self.result_bz_like = QLineEdit(self)
        self.bar=QMenuBar(self)
        self.file=self.bar.addMenu("其他工具")
        #工具红冲变量名
        self.bar_toolbox_hc=QAction("票据红冲",self)
        self.file.addAction(self.bar_toolbox_hc)
    # 工具箱按列排序
        self.l4 = QHBoxLayout()
        self.l4.addWidget(self.bar)
        # 列排序
        self.l1_px = QVBoxLayout()
        self.l1_px.addWidget(self.new)
        self.l1_px.addWidget(self.old_read)
        self.l1_px.addWidget(self.old_write)
        self.l1_px.addWidget(self.new_excel)
        # self.setLayout(self.l1_px)
        self.l2_px = QVBoxLayout()
        self.l2_px.addWidget(self.one)
        self.l2_px.addWidget(self.write_data)
        self.l2_px.addWidget(self.result)
        self.l2_px.addWidget(self.result_bz)
        # self.setLayout(self.l2_px)
        self.l3_px = QVBoxLayout()
        self.l3_px.addWidget(self.one_like)
        self.l3_px.addWidget(self.write_data_like)
        self.l3_px.addWidget(self.result_like)
        self.l3_px.addWidget(self.result_bz_like)
        # 行排序
        self.h1_px = QHBoxLayout()
        self.h1_px.addLayout(self.l1_px)
        self.h1_px.addLayout(self.l2_px)
        self.h1_px.addLayout(self.l3_px)
        #self.setLayout(self.h1_px)
       #在行排序的基础上加入菜单栏排序
        self.l1_addtools=QVBoxLayout()
        self.l1_addtools.addLayout(self.l4)
        self.l1_addtools.addLayout(self.h1_px)
        self.setLayout(self.l1_addtools)
    def new1(self):

        ss=os.path.isfile(r'登记表.xlsx')
        print(ss)
        if ss==True:
            QMessageBox.information(self, "消息框标题", "已存在", QMessageBox.Yes)
        if ss==False:
            wb = openpyxl.Workbook(r"登记表.xlsx")
            ws = wb.create_sheet(title="登记本", index=0)
            text = ["序号", "时间", "开始号", "领取数值", "结束号", "领取人", "退回时间", '段号']
            ws.append(text)
            wb.save(r"登记表.xlsx")
            wb.close()
            QMessageBox.information(self, "消息框标题", "创建初始数据完成。", QMessageBox.Yes)







    def new2(self):
        value, ok = QInputDialog.getText(self, "提示", "请输入开始票据\n\n请输入文本:", QLineEdit.Normal, "这是默认值")
        wb = openpyxl.load_workbook(r'登记表.xlsx')
        ws = wb["登记本"]
        max_column = ws.max_row
        num=max_column+1
        print(max_column)
        value=value.rjust(9, '0')
        ws["c"+f"{num}"]=value
        wb.save(r'登记表.xlsx')
        wb.close()
        QMessageBox.information(self, "消息框标题", "创建完成。", QMessageBox.Yes )
    def nwe3_1(self):
        if len(self.write_data_like.text()) !=0:
            num=int(self.i[0].value)+int(self.write_data_like.text())
            ss=str(num)
            self.result_new=(ss.rjust(9,'0'))
            self.result_like.setText(self.result_new)

        else:
            self.result_like.clear()

    def new3(self):
        self.result_like.clear()
        self.result_bz_like.clear()
        wb = openpyxl.load_workbook(r'登记表.xlsx')
        ws = wb["登记本"]
        max_column = ws.max_row
        for self.i in ws.iter_rows(min_row=max_column,min_col=3):
            self.one_like.setText(self.i[0].value)
        print(self.i[0].value,max_column)
        self.write_data_like.textChanged[str].connect(self.nwe3_1)
    def new4(self):
        '''
        a份数，b是结果 c领取人
        "序号","时间","开始号", "领取数值", "结束号", "领取人","退回时间",'段号'
        '''

        wb = openpyxl.load_workbook(r'登记表.xlsx')
        ws = wb["登记本"]
        max_column = ws.max_row
        a=self.write_data_like.text()
        b=self.result_like.text()
        c=self.result_bz_like.text()
        print(type(a),type(b),c)
        ws["d"+f"{max_column}"]=str(a)
        ws["e"+f"{max_column}"]=str(b)
        ws["f"+f"{max_column}"]=str(c)
        num=int(b)+1
        num=str(num).rjust(9,"0")
        ws["c"+f"{max_column+1}"]=str(num)
        #拼接短号
        text_new=ws["c" + f"{max_column}"].value + "-" + ws["e"+f"{max_column}"].value
        ws["h"+f'{max_column}']=str(text_new)
        #序号
        ws["a"+f'{max_column}']=int(max_column-1)
        #时间
        ws["b" + f'{max_column}']=str(datetime.date.today())
        wb.save(r'登记表.xlsx')
        wb.save(r'd:/票据登记表备份.xlsx')
        wb.close()
        QMessageBox.information(self,"消息框标题","保存成功。",QMessageBox.Yes )
        self.result_like.clear()
        self.result_bz_like.clear()
        self.write_data_like.clear()

class tool_new(QWidget):
    def __init__(self):
        super(tool_new, self).__init__()
        self.resize(300,300)
        self.id_name=QLabel("本段发票最后一张：")
        self.id_edit=QLineEdit(self)
        #保存按钮
        self.id_save=QPushButton("保存")
        #表格设置
        self.tableWidget = QTableWidget(self)
        #self.tableWidget.setRowCount(1)
        self.tableWidget.setColumnCount(4)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)#设置行
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)#设置行
        self.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)#只能选中一行
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)#不能进行修改
        self.tableWidget.setHorizontalHeaderLabels(['断号', '领用人', '领用日期', '退回时间'])
        #排序
        self.h1=QHBoxLayout()
        self.h1.addWidget(self.id_name)
        self.h1.addWidget(self.id_edit)
        self.h1.addWidget(self.id_save)

        #self.setLayout(self.v1)
        self.layout_1=QVBoxLayout()
        self.layout_1.addLayout(self.h1)
        self.layout_1.addWidget(self.tableWidget)
        self.setLayout(self.layout_1)
        #获取信息
        self.id_edit.textChanged[str].connect(self.read)
        #保存删除信息
        self.id_save.clicked.connect(self.move_id_save)

    def read(self):#表格获取函数
        zo_iterms=[]
        print(self.id_edit.text())
        ss=str(self.id_edit.text())
        id_idex = ss.rjust(9,"0")  # 待查信息
        print(id_idex)
        wb = openpyxl.load_workbook(r'登记表.xlsx')
        ws = wb["登记本"]
        cout = ws.max_row
        for i in ws.iter_rows(min_row=2,max_row=cout):#查找数据

            if i[4].value == id_idex:
                iterms=[i[7].value,i[5].value,i[1].value,i[6].value]
                print(iterms)
                zo_iterms.append(iterms)
            else:
                pass
        self.tableWidget.setRowCount(len(zo_iterms))#根据数量创建表格
        for i in range(len(zo_iterms)):#所有插入信息
            for j in range(4):
                print(i,j,str(zo_iterms[i][j]))
                items = QTableWidgetItem(str(zo_iterms[i][j]))
                self.tableWidget.setItem(i, j, items)
        self.tableWidget.cellClicked.connect(self.lalala)#获取表格的位置信号
    def lalala(self):
        #print(self.tableWidget.currentRow())#获取表格的
        self.tableWidget_str=self.tableWidget.item(int(self.tableWidget.currentRow()),0).text()
    def move_id_save(self):
        wb = openpyxl.load_workbook(r'登记表.xlsx')
        ws = wb["登记本"]
        max_row=0
        for i in ws.iter_rows(min_col=8):
            max_row+=1
            if str(i[0].value)==self.tableWidget_str:
                ws["g"+f'{max_row}']= str(datetime.date.today())
                wb.save(r'登记表.xlsx')
                wb.save(r'd:/票据登记表备份.xlsx')
                wb.close()
                self.tableWidget.clear()
                self.id_edit.clear()
                QMessageBox.information(self, "消息框标题", "保存成功。", QMessageBox.Yes)
            else:
                pass
if __name__ == '__main__':
    app=QApplication(sys.argv)
    tt=Demo()
    ss=tool_new()
    tt.show()
    tt.bar.triggered[QAction].connect(ss.show)
    app.exit(app.exec_())