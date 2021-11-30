#!/usr/bin/python

from src import generic 
import time 
import logging
import sys
import json
import argparse
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.FileHandler('logs/generic.log')
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)



def main():
	parser = argparse.ArgumentParser()

	parser.add_argument('-f','--path',help='File Absolute Path',required=True)
	parser.add_argument('-s','--store',help='Store to DB',default = 1)
	parser.add_argument('-y','--yara',help='Apply Yara Matcher',required=True)
	parser.add_argument('-e','--extract',help='Extract Features',required=True)
	parser.add_argument('--version', action='version', version='%(prog)s 1.0')
	results = parser.parse_args()
	
	logger.info('Starting Main Process at {}'.format(time.time()))
	gen_obj = generic.GenericParser(results.path)
	
	gen_obj.check_mime()
	if int(results.extract):
		gen_obj.filemeta()
	if int(results.yara):
		gen_obj.yara_match()
	gen_obj.get_stat()
	print json.dumps(gen_obj.file_meta,indent=4,sort_keys=True)
if __name__ == '__main__':
	main()
