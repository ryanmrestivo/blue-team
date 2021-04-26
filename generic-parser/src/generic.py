__author__     = 'Saikiran Uppu'
__copyright__  = 'Copyright ,2017'
__licence__    = ''
__version__    = '1.0.0'
__maintainer__ = 'Saikiran Uppu'
__email__      = 'iamsaikiran.official@gmail.com'
__status__     = 'Development'

import sys
import logging
import time
import json
import os
import magic
from pe_parser import PEFeatureExtractor , test
from pdf_parser import PDFFeatureExtractor , pdf_test
import hashlib
from yara_match import YaraClass
import math
from office_parser import OfficeParser

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.FileHandler('logs/generic.log')
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

class GenericParser:

        def __init__(self, file_path):
                self.file_path = file_path
		self.file_meta = {}
		self.file_meta['features'] = []
		self.macro = 1
		self.pdf_featues = ''
		self.pe_features = []
                self.mime_with_macro_office = {".doc": "application/msword",
                        ".dot": "application/msword",
                        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
                        ".docm": "application/vnd.ms-word.document.macroEnabled.12",
                        ".dotm": "application/vnd.ms-word.template.macroEnabled.12",
                        ".xls": "application/vnd.ms-excel",
                        ".xlt": "application/vnd.ms-excel",
                        ".xla": "application/vnd.ms-excel",
                        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        ".xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
                        ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
                        ".xltm": "application/vnd.ms-excel.template.macroEnabled.12",
                        ".xlam": "application/vnd.ms-excel.addin.macroEnabled.12",
                        ".xlsb": "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
                        ".ppt": "application/vnd.ms-powerpoint",
                        ".pot": "application/vnd.ms-powerpoint",
                        ".pps": "application/vnd.ms-powerpoint",
                        ".ppa": "application/vnd.ms-powerpoint",
                        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        ".potx": "application/vnd.openxmlformats-officedocument.presentationml.template",
                        ".ppsx": "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
                        ".ppam": "application/vnd.ms-powerpoint.addin.macroEnabled.12",
                        ".pptm": "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
                        ".potm": "application/vnd.ms-powerpoint.template.macroEnabled.12",
                        ".ppsm": "application / vnd.ms - powerpoint.slideshow.macroEnabled.12"
                                                                           }
                self.mime_with_macro_pdf = {
                        ".pdf": ["application/pdf", "application/x-pdf", "application/acrobat", "application/vnd.pdf", "text/pdf", "text/x-pdf"]
                }
		self.mime_executable = {
			".exe" : ["application/x-dosexec","application/octet-stream", "application/x-msdownload", "application/exe", "application/x-exe", "application/dos-exe", "vms/exe", "application/x-winexe", "application/msdos-windows", "application/x-msdos-program"]
		}
                self.mime_compressed = {}
                self.mime_packed = {}
                self.mime_no_macro = {}
		self.mime_web = {
			".htm" : "text/html",
			".html": "text/html",
			".swf" : "application/x-shockwave-flash",
			".jar" : "application/java-archive"

		}
	def md5_hash(self,fname):
    		hash_md5 = hashlib.md5()
    		with open(fname, "rb") as f:
        		for chunk in iter(lambda: f.read(4096), b""):
           			 hash_md5.update(chunk)
    		return hash_md5.hexdigest()

	def sha1_hash(self,fname):
                hash_sha1 = hashlib.md5()
                with open(fname, "rb") as f:
                        for chunk in iter(lambda: f.read(4096), b""):
                                 hash_sha1.update(chunk)
                return hash_sha1.hexdigest()
	
	def sha256_hash(self,fname):
                hash_sha256 = hashlib.sha1()
                with open(fname, "rb") as f:
                        for chunk in iter(lambda: f.read(4096), b""):
                                 hash_sha256.update(chunk)
                return hash_sha256.hexdigest()
	def entropy(self,fname):
		f = open(fname ,"rb")
		byteArr = map(ord, f.read())
		f.close()
		fileSize = len(byteArr)
		#print fileSize
		freqList = []
		for b in range(256):
    			ctr = 0
    			for byte in byteArr:
        			if byte == b:
            				ctr += 1
    		freqList.append(float(ctr) / fileSize)
		# Shannon entropy
		ent = 0.0
		for freq in freqList:
    			if freq > 0:
        			ent = ent + freq * math.log(freq, 2)
		ent = -ent
		return ent
	def file_size(self,fname):
		return os.path.getsize(fname)
	def get_stat(self):
		
		file_stats = os.stat(self.file_path)
		self.file_meta['protection_bytes'] = file_stats[0]
		self.file_meta['inode_number'] = file_stats[1]
		self.file_meta['device'] = file_stats[2]
		self.file_meta['no_of_hard_links'] = file_stats[3]
		self.file_meta['user_id_of_owner'] = file_stats[4]
		self.file_meta['group_id_of_owner'] = file_stats[5]
		self.file_meta['access_time'] = file_stats[7]
		self.file_meta['modification_time'] = file_stats[8]
		self.file_meta['meta_data_change_time'] = file_stats[9]
		return file_stats
				
	def filemeta(self):
		self.file_meta['file_path'] 	= self.file_path
		self.file_meta['file_name'] 	= self.file_path.split('/')[-1]
		self.file_meta['md5']           = self.md5_hash(self.file_path)
		self.file_meta['sha1']          = self.sha1_hash(self.file_path)
		self.file_meta['sha256']        = self.sha256_hash(self.file_path)
		self.file_meta['magic_info']    = self.magic_info
		self.file_meta['magic_buffer']  = self.magic_buffer
		self.file_meta['mime'] 		= self.magic_mime
		self.file_meta['macro'] 	= self.macro
		self.file_meta['entropy'] 	= self.entropy(self.file_path)
		self.file_meta['min_possible_file_size'] = self.file_meta['entropy'] * self.file_size(self.file_path)
		self.file_meta['size'] 		= self.file_size(self.file_path)
		self.file_meta['file_size_not_multiple_8'] =  self.file_meta['size'] % 8 
	def yara_match(self):
		self.file_meta['yara'] = []
		ys = YaraClass('src/rules/','.', True , self.file_path)
                ys.compile()
                #ys.scan_all()
                result = ys.scan_single(self.file_path)
                #print result
                final = []
                for a in result:
                	#print a
                        #print str(a)
                        final.append(str(a))
		final = list(set(final))
		self.file_meta['yara'] = final
        def check_mime(self):
                logger.info('GenericParser on file {} starts at {}'.format(self.file_path, time.time()))
                self.magic_info = magic.from_file(self.file_path)
                self.magic_buffer = magic.from_buffer(open(self.file_path).read(1024))
                self.magic_mime = magic.from_file(self.file_path, mime=True)
                #print self.magic_mime
                if self.magic_mime in self.mime_with_macro_office.values():
                        logger.info('Office File {} mime {}'.format(self.file_path, self.magic_mime))
			obj = OfficeParser(self.file_path)
			results =  obj.analysis()
			self.file_meta['features'] = results
                        logger.info('Sending File to office_extractor')
                elif self.magic_mime in self.mime_with_macro_pdf.values()[0]:
                        logger.info('Pdf File {} mime {}'.format(self.file_path, self.magic_mime))
			logger.info('Seding file to pdf_extractor')
			#self.file_meta['features'] = pdf_test(self.file_path)
			self.pdf_features = pdf_test(self.file_path)
			#print self.pdf_features
			self.file_meta['features'] = self.pdf_features
			self.filemeta()
			#print self.file_meta
		elif self.magic_mime in self.mime_compressed:
                        logger.info('Compressed File {} mime {}'.format(self.file_path, self.magic_mime))
                elif self.magic_mime in self.mime_packed:
                        logger.info('Packed File {} mime {}'.format(self.file_path, self.magic_mime))
		elif self.magic_mime in self.mime_executable['.exe'][0]:
                        logger.info('Executable File {} mime {}'.format(self.file_path, self.magic_mime))
			self.pe_features = test(self.file_path)	
			self.file_meta['features'] = self.pe_features
			logger.info('Sending File {} to Exe Extractor'.format(self.file_path))
                elif self.magic_mime in self.mime_web.values():
			logger.info('Web File {} mime {}'.format(self.file_path, self.magic_mime))
			
		elif self.magic_mime in self.mime_no_macro:
                        logger.info('NonMacro File {} mime {}'.format(self.file_path, self.magic_mime))
                else:
			self.macro = 0



def parser():
        parser = GenericParser(file_path)
        parser.check_mime()
	parser.filemeta()
	stats = parser.get_stat(file_path)	
	print stats

def main():
        print 'This is intended for module purpose only'
        logger.info('Checking logs working')
        parser()


if __name__ == '__main__':
        main()

