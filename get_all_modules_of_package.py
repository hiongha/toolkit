#!/usr/bin/python
#coding=utf-8
#xinghe <xingh3223@berryoncology.com>

import os.path
import importlib
import pkgutil
import sys,inspect
'''
从模块(或包)中获取全部类的全部子类
'''

def get_all_modules(moduleStr):

	'''
	功能: 获取模块(包)的全部子模块
	moduleStr: <strings> 模块名称
	'''
	modules = []
	try:
		module = importlib.import_module(moduleStr) #输入的package变量字符串变量,将它的type转换成module

		for filefinder,name,ispkg in pkgutil.iter_modules([os.path.dirname(module.__file__)],module.__name__+"."):
			if ispkg == False: #如果是模块,就导入该模块,并记录模块名称 
				try:
					module = importlib.import_module(name)
					modules.append(module)
				except:
					print("无法导入该模块: %s"%name)
			else:
				modules += get_all_modules(moduleStr = name)
	except:
		print('模块可能不存在,请检查: %s'%module)
	return modules	


def get_all_classes(module):

	'''
	获得模块的全部类.
	module: <module> 模块
	'''
	cls = set()
	for name,obj in inspect.getmembers(module,inspect.isclass):
		if obj.__module__.split('.')[0] == module.__name__.split('.')[0]:
			cls.add(obj)
	return cls


def get_all_subclasses(modules = list()):

	'''
	获取模块列表中每个模块的全部类的全部子类.
	modules: <list> 模块列表
	'''
	class_subclasses_dict = {}
	for i in set(modules):
		classes = get_all_classes(i)
		for each in classes:
			subclasses = each.__subclasses__()
			if subclasses == []:
				class_subclasses_dict[each.__module__+"."+each.__name__] = '' 
			else:
				subclasses_name = [ n.__module__+"."+n.__name__ for n in subclasses ]
				class_subclasses_dict[each.__module__+"."+each.__name__] = subclasses_name 
				try:
					class_subclasses_dict += get_all_subclasses(subclasses)
				except:
					pass
	return class_subclasses_dict

if __name__ == '__main__':
	
	name = 'subprocess'  #test: 输入的是模块,不是包
	#name = 'email.mime'  #test: 输入的是包
	module = importlib.import_module(name)
	if module.__package__ != '':
		modules = get_all_modules(name)
	else:
		modules = [module]

	print("所有的模块:")
	print(modules)
	print()
	print("每个类的子类,第一个为父类:")	
	class_subclasses_dict = get_all_subclasses(modules)
	for key,value in class_subclasses_dict.items():
		print(key)
		print("\n".join(["\t|__"+t for t in value]))
		print()

