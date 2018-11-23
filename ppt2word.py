# -*- coding: utf-8 -*-
"""
Created on Sun Feb 25 12:37:48 2018

@author:zhangzhennudt
@email:zhangzhennudt@126.com
"""

import win32com
import os
from win32com.client import Dispatch, constants
ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Visible = 1
print("开始导出")
pptSel = ppt.Presentations.Open("your ppt path")
win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
f = open("your output txt path","w")
slide_count = pptSel.Slides.Count
for i in range(1,slide_count + 1):
  shape_count = pptSel.Slides(i).Shapes.Count
  print(shape_count)
  for j in range(1,shape_count + 1):
    if pptSel.Slides(i).Shapes(j).HasTextFrame:
      s = pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
      f.write(s+ "\n")
f.close()
ppt.Quit()
print("导出完毕")
