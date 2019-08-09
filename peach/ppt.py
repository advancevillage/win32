#!/usr/bin/python
# -*- coding: UTF-8 -*-
import win32com
import json
import time
from win32com.client import Dispatch

'''
 定义Shape 类型常量
'''
PPT_SHAPE_TABLE_CODE = 19
PPT_SHAPE_CHART_xlLineMarkers =  65         #数据点折线图
PPT_SHAPE_CHART_xlCylinderBarClustered = 95 #簇状条形圆柱图
PPT_SHAPE_CHART_xlColumnClustered = 51	    #簇状柱形图。
PPT_SHAPE_CHART_xlColumnStacked = 52	    #堆积柱形图。
PPT_SHAPE_CHART_xlPie = 5                   #饼图
PPT_SHAPE_CHART_xlXYScatterLines = 74       #折线散点图
PPT_SHAPE_CHART_xlXYScatterLinesNoMarkers = 75 #无数据点折线散点图
PPT_SHAPE_CHART_xl3DPie	= -4102                #3D饼图
PPT_SHAPE_CHART_xl3DColumnStacked = 55         #三维堆积柱形图
PPT_SHAPE_CHART_xlLineMarkers = 65             #数据点折线图
PPT_SHAPE_CHART_xlLineStacked100 = 64          #百分比堆积折线图
PPT_SHAPE_CHART_xl3DLine = -4101               #三维折线图

class PPT:
    name = ""
    visible = 1
    ppt = object

    def __init__(self):
        self.app = win32com.client.Dispatch('PowerPoint.Application')

    def open(self, name, visible):
        self.name = name
        self.visible = visible
        self.app.Visible = self.visible
        self.ppt = self.app.Presentations.Open(self.name)

    def save(self):
        self.ppt.Save()

    def close(self):
        self.ppt.Close()
        self.app.Quit()

    '''
    @param: slideID  幻灯片ID
    @param: tableID  Shape Table ID
    @param: i        起始行
    @param: j        起始列
    @param: values[][] 数据
    '''
    def write2table(self, slideID, tableID, i, j, values):
        table = self.ppt.Slides(slideID).Shapes(tableID).Table
        rowCount = table.Rows.Count
        colCount = table.Columns.Count
        rc = len(values)
        cc = len(values[0])
        x  = 0
        for m in range(i, rowCount + 1):
            y = 0
            for n in range(j, colCount + 1):
                value = values[x][y]
                if isinstance(value, str):
                    value = value.decode("utf-8")
                else:
                    pass
                if x < rc and y < cc:
                    table.Cell(m, n).Shape.TextFrame.TextRange.Text = value
                else:
                    pass
                y += 1
            x += 1

    '''
    @brief:  报表数据显示一样由于数据过导致分割成多个表显示
    @:param  slideID     幻灯片ID
    @:param  tableIDs    幻灯片ID中多个表ID
    @:param  i,j         表中写入数据的起始位置
    @:param  values      数据集
    '''
    def write2tables(self, slideID, tableIDs, i, j, values):
        x = 0
        rc = len(values)
        cc = len(values[0])
        for tableID in tableIDs:
            table = self.ppt.Slides(slideID).Shapes(tableID).Table
            rowCount = table.Rows.Count
            colCount = table.Columns.Count
            for m in range(i, rowCount + 1):
                y = 0
                for n in range(j, colCount + 1):
                    value = values[x][y]
                    if isinstance(value, str):
                        value = value.decode("utf-8")
                    else:
                        pass
                    if x < rc and y < cc:
                        table.Cell(m, n).Shape.TextFrame.TextRange.Text = value
                    else:
                        pass
                    y += 1
                x += 1
        time.sleep(3)

    '''
    @param: slideID    幻灯片ID
    @param: chartType  图表类型
    @param: left       幻灯片整个画布的横坐标
    @param: top        幻灯片整个画布的纵坐标
    @param: width      图标宽度
    @param: height     图标高度
    @param: values[][] 数据
    '''
    def write2chart(self, slideID, chartType, left, top, width, height, values):
        colums = ['A','B','C','D', 'E', 'F', 'G','H','I','J','K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'S', 'W', 'X', 'Y', 'Z']
        chart = self.ppt.Slides(slideID).Shapes.AddChart2(-1, chartType, left, top, width, height).Chart
        chart.ChartColor = 13
        rc = len(values)
        cc = len(values[0])
        chart.ChartData.Activate()
        workbook = chart.ChartData.Workbook
        sheet = workbook.Worksheets('Sheet1'.decode("utf-8"))
        for i in range(0, rc):
            for j in range(0, cc):
                value = values[i][j]
                if isinstance(value, str):
                    value = value.decode("utf-8")
                else:
                    pass
                sheet.Cells(i + 1, j + 1).Value = value
        chart.ChartWizard(Source="='Sheet1'!$A$1:${}${}".format(colums[cc - 1], rc), Gallery = chartType)
        time.sleep(1)
        workbook.Close()
        time.sleep(1)

    '''
    @:param  slideID  幻灯片ID
    @:param  textID   文本框ID
    @:param  src      需要被替换的串
    @:param  dest     替换成的串
    '''
    def write2text(self, slideID, textID, src, dest):
        text = self.ppt.Slides(slideID).Shapes(textID).TextFrame
        content = text.TextRange.Text
        contentTemp = content.encode("utf-8")
        contentTemp = contentTemp.replace(src, dest)
        content = contentTemp.decode("utf-8")
        text.TextRange.Text = content

    '''
    @:param  返回最后一个幻灯片ID
    '''
    def lastSlide(self):
        lastSlide = self.ppt.Slides.Count
        return lastSlide

    '''
    @brief: 解析PPT 查询Slides, Shapes, Type
    '''
    def parse(self):
        result = {}
        slides_count = self.ppt.Slides.Count
        result['total'] = slides_count
        result['slides'] = []
        for i in range(1, slides_count + 1):
            o1 = {}
            shapes_count = self.ppt.Slides(i).Shapes.Count
            o1['total'] = shapes_count
            o1['shapes'] = []
            for j in range(1, shapes_count + 1):
                o2 = {}
                type = self.ppt.Slides(i).Shapes(j).Type
                name = self.ppt.Slides(i).Shapes(j).Name
                o2['type'] = type
                o2['name'] = name
                o2['slideID'] = i
                o2['shapeID'] = j
                o1['shapes'].append(o2)
            result['slides'].append(o1)
        print json.dumps(result).decode("unicode_escape")
