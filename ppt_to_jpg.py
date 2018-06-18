# -*- coding: utf-8 -*-
import os
import sys
import win32com
import win32com.client

ppSaveAsJPG = 17

def ppt_to_jpg(ppt_file_name,output_dir_name):
	'''将PPT另存为JPG格式
	arguments:
		ppt_file_name: 要转换的ppt文件，
		output_dir_name：转换后的存放JPG文件的目录
	'''
	# 启动PPT
	ppt_app = win32com.client.Dispatch('PowerPoint.Application')
	# 设置为0表示后台运行，不显示，1则显示
	ppt_app.Visible = 1
	# 打开PPT文件
	ppt = ppt_app.Presentations.Open(ppt_file_name)
	# 另存为
	ppt.SaveAs(output_dir_name, ppSaveAsJPG)
	# 退出
	ppt_app.Quit()

if __name__ == '__main__':
	current_dir = os.sys.path[0]
	dir_list = os.listdir(current_dir)
	# 当前目录下所有的PPT文件,eg: ppt_name.ppt
	ppt_file_names = (fns for fns in dir_list if fns.endswith(('.ppt','.pptx')))
	# 当前目录下所有的PPT文件名，这两者的区别在于有无后缀名,eg: ppt_name
	ppt_names = (os.path.splitext(fns)[0] for fns in dir_list if fns.endswith(('.ppt','.pptx')))
	# ppt_names = (fns.split('.')[0] for fns in dir_list if fns.endswith(('.ppt','.pptx')))
	for ppt_file_name,ppt_name in zip(ppt_file_names,ppt_names):
		# 该PPT的完整路径文件名，eg: F:\\test\\ppt_name.ppt
		ppt_file_name = os.path.join(current_dir,ppt_file_name)
		# 需要新建一个与PPT同名的文件，获取完整路径,eg:  F:\\test\\ppt_name
		ppt_dir_path = os.path.join(current_dir,ppt_name)
		os.mkdir(ppt_dir_path)
		# print ppt_file_name, ppt_dir_path
		ppt_to_jpg(ppt_file_name,ppt_dir_path)
