'''
######
#	Module for logging and printing log messages
######
'''

import logging
import os

class Log:
	def __init__(self, file_path=None, file_name=None, logging_level=None):
		######
		#	If file path is unspecified, then default to current directory
		######
		if file_path is None:
			self.file_path = os.getcwd()
		elif isinstance(file_path, bool):
			print("Boolean not acceptable as a file path.")
			return False
		elif isinstance(file_path, int):
			print("Integers not acceptable as a file path.")
			return False
		else:
			self.file_path = file_path
		
		######
		#	If logging level is unspecified, default to WARNING
		######
		if logging_level is None:
			self.log_level = "logging.INFO"
		else:
			self.log_level = logging_level
		
		######
		#	If file name is unspecified, default to generic name
		######
		if file_name is None:
			try:
				logging.basicConfig(filename=self.file_path + "/log.txt", level=logging.INFO)
			except IOError:
				print("Trying without '/' at end of file path.")
				try:
					logging.basicConfig(filename=self.file_path + "log.txt", level=logging.INFO)
				except IOError:
					print("Something went wrong with logging.basicConfig() when file_name is None")
					return False
		else:
			try:
				logging.basicConfig(filename=self.file_path + "/" + file_name, level=logging.INFO)
			except IOError:
				print("Trying without '/' at end of file path.")
				try:
					logging.basicConfig(filename=self.file_path + file_name, level=logging.INFO)
				except IOError:
					print("Something went wrong with logging.basicConfig()")
					return False
					
	def log_debug(self, m):
		logging.debug(m)
	
	def log_info(self, m):
		logging.info(m)
	
	def log_warning(self, m):
		logging.warning(m)
		
	def log_error(self, m):
		logging.error(m)
	
	def log_critical(self, m):
		logging.critical(m)
