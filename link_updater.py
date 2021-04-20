import os
import time
import datetime
from win32com.client import Dispatch

def get_workbooks(dir_list):
	#folder = subprocess.Popen(r'explorer /select,"C:\"')
	#get local directory
	ext = ['xlsx', 'xlsm', 'xls', 'csv', 'xml']
	wb_dirs = []
	for dir_root in dir_list:
		try:
			[wb_dirs.append(os.path.join(dir_root,x)) for x in os.listdir(dir_root) if x.endswith(tuple(ext)) and x[0] != "~"]
		except FileNotFoundError as e:
			create_log(str(e))

	return wb_dirs

def create_log(message, console = True):
	file_name = 'log.txt'

	with open(file_name, 'a+') as f:
		if console is True:
			print(message)
		f.write(message+'\n')

def exit_message(message, exit_time):
	create_log(message+'\n') #seperates run instance logs
	time.sleep(exit_time)
	quit()
		
def get_directoires():

	file_name = 'directories.txt'
	dir_list = ['#INCLUDE ALL FOLDER LOCATIONS THAT CONTAIN THE EXCEL FILES WITH LINKED VARIABLES\n',
				'C:\\Users\\Folder Location X\n',
				'C:\\Users\\Folder Location Y']

	try:
		with open(file_name, 'r') as f:
			lines = f.readlines()
			dir_list = []
			for line in lines:
				if line[0] == '#': #allows for comments to be included
					continue
				if line[1:3] != ':\\':
					exit_message('Invalid links contained within directories. Please fix before running the application again.\n', 6)
					
				dir_list.append(line.strip('\n'))

	except FileNotFoundError:
		with open(file_name, 'x') as f:
			f.writelines(dir_list)

		exit_message('Created directories.txt where the application is being ran. Update file with the correct folder locations before running the application again.', 6)
	
	return dir_list

def run_macro(wb_dir, com_instance):

	wb = com_instance.workbooks.open(wb_dir)
	com_instance.AskToUpdateLinks = False

	try:
		#recalculate links
		wb.UpdateLink(Name = wb.LinkSources())

	except:
		create_log(wb_dir + " failed to update", False)

		wb.Close(True)
		return False

	wb.Close(True)
	return True

def run_excel(workbook_dirs):

	xl_app = Dispatch("Excel.Application")
	xl_app.Visible = False
	xl_app.DisplayAlerts = False
	updated_wbs = 0

	for wb_dir in workbook_dirs:
		if run_macro(wb_dir, xl_app) is True:
			updated_wbs += 1
		
	xl_app.Quit()

	return updated_wbs

if __name__ == "__main__":

	dir_list = get_directoires()
	workbook_dirs = get_workbooks(dir_list)
	total_wbs = str(len(workbook_dirs))

	if workbook_dirs == []:
		exit_message('No excel files found within the listed directories.\nClosing application', 4)
	else:
		create_log('Found ' + total_wbs + ' workbook(s) to update')

	if input("Update linked references? Type y or n: ").lower() == 'y':
		updated_wbs = str(run_excel(workbook_dirs))

		create_log(updated_wbs + ' out of ' + total_wbs + ' workbooks updated')
		if updated_wbs != total_wbs:
			print("Some files may not contain a linked reference. Check log.txt to see which files didn't get updated")

	exit_message(str(time.strftime("Completed operations at %m-%d %H:%M",time.localtime())), 5)