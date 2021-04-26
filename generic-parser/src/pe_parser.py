#!/usr/bin/python
import pefile
import logging
import sys
import hashlib
import json
import datetime

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.FileHandler('logs/generic.log')
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class PEFeatureExtractor():
	def __init__(self , filepath):
		self.filepath = filepath
		self.pe_features = {}
		self.rare_features = {}
		self.pe_feature_empty = -1
		self.pe_feature_default = 0
		self.pe_section_status = 1
		self.pe_md5 = ''
		self.anti_vm = []
		self.anti_dbg = []
		self.embeded_files = []
		self.urls = []
		self.import_symbol_count = 0
		self.export_symbol_count = 0
		self.bound_import_symbol_count = 0
		self.datadirs = { 0: 'IMAGE_DIRECTORY_ENTRY_EXPORT',
             			  1:'IMAGE_DIRECTORY_ENTRY_IMPORT',
             			  2:'IMAGE_DIRECTORY_ENTRY_RESOURCE',
             			  5:'IMAGE_DIRECTORY_ENTRY_BASERELOC',
             			  12:'IMAGE_DIRECTORY_ENTRY_IAT'} 
		self.section_flags = ['IMAGE_SCN_MEM_EXECUTE', 'IMAGE_SCN_CNT_CODE', 'IMAGE_SCN_MEM_WRITE', 'IMAGE_SCN_MEM_READ']
        	self.rawexecsize = 0
        	self.vaexecsize = 0
	
	def md5(self, fname):
    		hash_md5 = hashlib.md5()
    		with open(fname, "rb") as f:
        		for chunk in iter(lambda: f.read(4096), b""):
        			hash_md5.update(chunk)
    		return hash_md5.hexdigest()
	def antiVM_detector(self, fname):
		
		vm_signatures =  {
        		"Red Pill":"\x0f\x01\x0d\x00\x00\x00\x00\xc3",
        		"VirtualPc trick":"\x0f\x3f\x07\x0b",
        		"VMware trick":"VMXh",
        		"VMCheck.dll":"\x45\xC7\x00\x01",
        		"VMCheck.dll for VirtualPC":"\x0f\x3f\x07\x0b\xc7\x45\xfc\xff\xff\xff\xff",
        		"Xen":"XenVMM",
        		"Bochs & QEmu CPUID Trick":"\x44\x4d\x41\x63",
        		"Torpig VMM Trick": "\xE8\xED\xFF\xFF\xFF\x25\x00\x00\x00\xFF\x33\xC9\x3D\x00\x00\x00\x80\x0F\x95\xC1\x8B\xC1\xC3",
        		"Torpig (UPX) VMM Trick": "\x51\x51\x0F\x01\x27\x00\xC1\xFB\xB5\xD5\x35\x02\xE2\xC3\xD1\x66\x25\x32\xBD\x83\x7F\xB7\x4E\x3D\x06\x80\x0F\x95\xC1\x8B\xC1\xC3"
          		}
		# Reading file in Binary
		with open(fname ,'rb') as f:
			curr_buff = f.read()
			for vm_sig in vm_signatures:
				present = curr_buff.find(vm_signatures[vm_sig])
				if present > -1:
					  self.anti_vm.append("0x%x %s" % (present, vm_sig))
	def extractor(self):
		logger.info('Features for PE file {} being extracted'.format(self.filepath))
		self.pe_md5 = self.md5(self.filepath)
		with open(self.filepath ,'rb') as f:	
			data = f.read()
			self.pe_success , self.pe_error = self.try_with_pe_open(data)
			#print self.pe_success
	def try_with_pe_open(self,file_data ):
		logger.info('Trying to see if PE is able to open {}'.format(self.filepath))
		try:
			pe = pefile.PE(data=file_data, fast_load=False)
		except Exception as e:
			logger.error('Error: on file {}: {}'.format(self.filepath,str(e)))
			error_str =  '(Exception)'
            		return None, error_str
		if (pe.PE_TYPE is None or pe.OPTIONAL_HEADER is None or len(pe.OPTIONAL_HEADER.DATA_DIRECTORY) < 7):
                	logger.error('Error in pe validation : on file: %s' % self.filepath)
                	error_str =  '(Exception):, %s' % (str(self.filepath))
                	return None, error_str
                return pe, None
	def check_for_pe_success(self):
		if self.pe_success:
			logger.info('PE Feature extraction success on file {}'.format(self.filepath))
			return 1
		else:
			logger.info('PE Feature extraction failed on file {}'.format(self.filepath))
			return 0
	def extract_features(self):
		try:
			pass
		except Exception as e:
			logger.error('PE Feature extraction failed on file {}'.format(self.filepath))
	def get_pe_features(self):
		return self.pe_features
	def get_rare_features(self):
		return self.rare_features
	def put_pe_features(self):
		self.pe_features_list = ['check_sum', 'generated_check_sum', 'compile_date', 'debug_size', 'export_size', \
				    'iat_rva', 'major_version','minor_version', 'number_of_bound_import_symbols', \
				    'number_of_bound_imports', 'number_of_export_symbols', \
                                    'number_of_import_symbols', 'number_of_imports', 'number_of_rva_and_sizes', \
				    'number_of_sections', 'pe_warnings','std_section_names', 'total_size_pe', 'virtual_address',\
				    'virtual_size', 'virtual_size_2', \
                                    'datadir_IMAGE_DIRECTORY_ENTRY_BASERELOC_size', 'datadir_IMAGE_DIRECTORY_ENTRY_RESOURCE_size', \
                                    'datadir_IMAGE_DIRECTORY_ENTRY_IAT_size', 'datadir_IMAGE_DIRECTORY_ENTRY_IMPORT_size', \
                                    'pe_char', 'pe_dll', 'pe_driver', 'pe_exe', 'pe_i386', 'pe_majorlink', 'pe_minorlink', \
                                    'sec_entropy_data', 'sec_entropy_rdata', 'sec_entropy_reloc', 'sec_entropy_text',\
				    'sec_entropy_rsrc', \
                                    'sec_rawptr_rsrc', 'sec_rawsize_rsrc', 'sec_vasize_rsrc', 'sec_raw_execsize', \
                                    'sec_rawptr_data', 'sec_rawptr_text', 'sec_rawsize_data', 'sec_rawsize_text', \
				    'sec_va_execsize', 'sec_vasize_data', 'sec_vasize_text', 'size_code', \
				    'size_image', 'size_initdata', 'size_uninit']
	def put_rare_features(self):
		self.rare_features_list = ['imported_symbols', 'section_names', 'pe_warning_strings']
	def pe_sections(self):
		self.pe_sections = ['.text', '.bss', '.rdata', '.data', '.rsrc', '.edata', '.idata', \
                        	    '.pdata', '.debug', '.reloc', '.stab', '.stabstr', '.tls', \
                        	    '.crt', '.gnu_deb', '.eh_fram', '.exptbl', '.rodata']
		self.pe_additional_sections = [ '/' + str(i) for i in range(200)]
		for sec in self.pe_additional_sections:
			self.pe_sections.append(sec)
	def pe_mapping_section(self):
		try:
			self.pe_features['anti_debugging_capabilities'] = self.anti_dbg
			#Initial Assignment
			self.rare_features['section_names'] = []
			for each_sec in self.pe_success.sections:
				section_name = self.convertToAsciiNullTerm(each_sec.Name).lower()
				self.rare_features['section_names'].append(section_name)
				if (section_name not in self.pe_sections):
                			self.pe_section_status = 0
			self.pe_features['std_section_names'] = self.pe_section_status
			self.pe_features['debug_size']         = self.pe_success.OPTIONAL_HEADER.DATA_DIRECTORY[6].Size
        		self.pe_features['major_version']      = self.pe_success.OPTIONAL_HEADER.MajorImageVersion
                	self.pe_features['minor_version']      = self.pe_success.OPTIONAL_HEADER.MinorImageVersion
                	self.pe_features['iat_rva']            = self.pe_success.OPTIONAL_HEADER.DATA_DIRECTORY[1].VirtualAddress
        		self.pe_features['export_size']	       = self.pe_success.OPTIONAL_HEADER.DATA_DIRECTORY[0].Size
        		self.pe_features['check_sum']	       = self.pe_success.OPTIONAL_HEADER.CheckSum
			try:
				self.pe_features['generated_check_sum'] = self.pe_success.generate_checksum()
			except ValueError:
				self.pe_features['generated_check_sum'] = 0
			if (len(self.pe_success.sections) > 0):
            			self.pe_features['virtual_address']         = self.pe_success.sections[0].VirtualAddress
            			self.pe_features['virtual_size']	    = self.pe_success.sections[0].Misc_VirtualSize
        			self.pe_features['number_of_sections']      = self.pe_success.FILE_HEADER.NumberOfSections
        			self.pe_features['compile_date']            = self.pe_success.FILE_HEADER.TimeDateStamp
        			self.pe_features['number_of_rva_and_sizes'] = self.pe_success.OPTIONAL_HEADER.NumberOfRvaAndSizes
        			self.pe_features['total_size_pe']	    = len(self.pe_success.__data__)	
				self.pe_features['import_symbols'] = []
				self.pe_features['import_bound_symbols']  = []
			if hasattr(self.pe_success,'DIRECTORY_ENTRY_IMPORT'):
            			self.pe_features['number_of_imports'] = len(self.pe_success.DIRECTORY_ENTRY_IMPORT)
            			for module in self.pe_success.DIRECTORY_ENTRY_IMPORT:
                			self.import_symbol_count  += len(module.imports)
					self.pe_features['import_symbols'].append(str(module.imports[0].name))
            			self.pe_features['number_of_import_symbols'] = self.import_symbol_count
			
			if hasattr(self.pe_success, 'DIRECTORY_ENTRY_BOUND_IMPORT'):
            			self.pe_features['number_of_bound_imports'] = len(self.pe_success.DIRECTORY_ENTRY_BOUND_IMPORT)
            			for module in self.pe_success.DIRECTORY_ENTRY_BOUND_IMPORT:
                			self.bound_import_symbol_count += len(module.entries)
					self.pe_features['import_bound_symbols'].append(str(module.imports[0].name))
           			self.pe_features['number_of_bound_import_symbols'] = self.bound_import_symbol_count
			if hasattr(self.pe_success,'DIRECTORY_ENTRY_EXPORT'):
            			self.pe_features['number_of_export_symbols'] = len(self.pe_success.DIRECTORY_ENTRY_EXPORT.symbols)
            			symbol_set = set()
            			for symbol in self.pe_success.DIRECTORY_ENTRY_EXPORT.symbols:
                			symbol_info = 'unknown'
                			if (not symbol.name):
                    				symbol_info = 'ordinal=' + str(symbol.ordinal)
                			else:
                    				symbol_info = 'name=' + symbol.name

                		symbol_set.add(convertToUTF8('%s'%(symbol_info)).lower())
            			self.pe_features['ExportedSymbols'] = list(symbol_set)
			self.warnings = self.pe_success.get_warnings()
        		if (self.warnings):
            			self.pe_features['pe_warnings'] = 1
            			self.pe_features['pe_warning_strings'] = self.warnings
        		else:
            			self.pe_features['pe_warnings'] = 0
			
			if hasattr(self.pe_success, 'DIRECTORY_ENTRY_IMPORT'):
            			symbol_set = set()
            			for module in self.pe_success.DIRECTORY_ENTRY_IMPORT:
                			for symbol in module.imports:
                   				symbol_info = 'unknown'
                    				if symbol.import_by_ordinal is True:
                        				symbol_info = 'ordinal=' + str(symbol.ordinal)
                    				else:
                        				symbol_info = 'name=' + symbol.name
                    				if symbol.bound:
                        				symbol_info += ' bound=' + str(symbol.bound)

                    				symbol_set.add(self.convertToUTF8('%s:%s'%(module.dll, symbol_info)).lower())

            			self.pe_features['imported_symbols'] = list(symbol_set)
			if (len(self.pe_success.sections) >= 2):
            			self.pe_features['virtual_size_2']     = self.pe_success.sections[1].Misc_VirtualSize

        		self.pe_features['size_image']             = self.pe_success.OPTIONAL_HEADER.SizeOfImage
        		self.pe_features['size_code']              = self.pe_success.OPTIONAL_HEADER.SizeOfCode
        		self.pe_features['size_initdata']          = self.pe_success.OPTIONAL_HEADER.SizeOfInitializedData
        		self.pe_features['size_uninit']            = self.pe_success.OPTIONAL_HEADER.SizeOfUninitializedData
        		self.pe_features['pe_majorlink']           = self.pe_success.OPTIONAL_HEADER.MajorLinkerVersion
        		self.pe_features['pe_minorlink']           = self.pe_success.OPTIONAL_HEADER.MinorLinkerVersion
        		self.pe_features['pe_driver']              = 1 if self.pe_success.is_driver() else 0
        		self.pe_features['pe_exe']                 = 1 if self.pe_success.is_exe() else 0
        		self.pe_features['pe_dll']                 = 1 if self.pe_success.is_dll() else 0
        		self.pe_features['pe_i386']                = 1
        		if self.pe_success.FILE_HEADER.Machine != 0x014c:
            			self.pe_features['pe_i386'] = 0
        		self.pe_features['pe_char']                = self.pe_success.FILE_HEADER.Characteristics

			
			for idx, datadir in self.datadirs.items():
            			datadir = pefile.DIRECTORY_ENTRY[ idx ]
            			if len(self.pe_success.OPTIONAL_HEADER.DATA_DIRECTORY) <= idx:
                			continue
				directory = self.pe_success.OPTIONAL_HEADER.DATA_DIRECTORY[idx]
            			self.pe_features['datadir_%s_size' % datadir] = directory.Size
				
			for sec in self.pe_success.sections:
            			if not sec:
                			continue

            			for char in self.section_flags:
                			if hasattr(sec, char):
                 		     	   self.rawexecsize += sec.SizeOfRawData
                    			   self.vaexecsize  += sec.Misc_VirtualSize
                    			break		
				secname = self.convertToAsciiNullTerm(sec.Name).lower()
				secname = secname.replace('.','')
				self.pe_features['sec_entropy_%s' % secname ] = sec.get_entropy()
				self.pe_features['sec_rawptr_%s' % secname] = sec.PointerToRawData
				self.pe_features['sec_rawsize_%s' % secname] = sec.SizeOfRawData
				self.pe_features['sec_vasize_%s' % secname] = sec.Misc_VirtualSize	
				self.pe_features['sec_va_execsize'] = self.rawexecsize
				self.pe_features['sec_raw_execsize'] = self.vaexecsize
			self.pe_features['anti_vm_capabilities'] = self.anti_vm
			
			# Extracting anti debug capabilities
			anti_dbg_signatures = ['CheckRemoteDebuggerPresent', 'FindWindow', 'GetWindowThreadProcessId', 'IsDebuggerPresent', 'OutputDebugString', 'Process32First', 'Process32Next', 'TerminateProcess',  'UnhandledExceptionFilter', 'ZwQueryInformation']
			for entry in pe_success.DIRECTORY_ENTRY_IMPORT:			
				 for imp in entry.imports:
            				if (imp.name != None) and (imp.name != ""):
                				for anti in anti_dbg_signatures:
                    					if imp.name.startswith(anti):
                        					self.anti_dbg.append("%s %s" % (hex(imp.address),imp.name))
			self.pe_features['anti_debugging_capabilities'] = self.anti_dbg	
			self.pe_features['compile_time'] = datetime.datetime.fromtimestamp(pe_success.FILE_HEADER.TimeDateStamp)
		except Exception as e:
			logger.error('Error in mapping pe section headers for {} {}'.format(self.filepath, str(e)))
	def convertToAsciiNullTerm(self , name):
		name = name.split('\x00', 1)[0]
    		return name.decode('ascii', 'ignore')
	def convertToUTF8(self,s):
    		if (isinstance(s, unicode)):
        		return s.encode( "utf-8" )
    		try:
        		u = unicode( s, "utf-8" )
    		except:
        		return str(s)
    		utf8 = u.encode( "utf-8" )
    		return utf8
	def createFeatureDict(self):
		for i in self.pe_features_list:
			self.pe_features[i] = self.pe_feature_empty
		for i in self.rare_features_list:
			self.rare_features[i] = self.pe_feature_empty
def main():
	peExtractor = PEFeatureExtractor(sys.argv[1])
	peExtractor.extractor()
	peStatus = peExtractor.check_for_pe_success()
	if peStatus:
		peExtractor.extract_features()
		peExtractor.put_pe_features()
		peExtractor.put_rare_features()
		peExtractor.createFeatureDict()
		peExtractor.pe_sections()
		peExtractor.pe_mapping_section()
		print json.dumps(peExtractor.get_pe_features(),sort_keys=True, indent=4)
		print json.dumps(peExtractor.get_rare_features(),sort_keys=True, indent=4)
	else:
		logger.info('Not able to extract PE on file {} .. Exiting'.format('/home/ransom/exe_example'))

def test(filename):
	peExtractor = PEFeatureExtractor(filename)
        peExtractor.extractor()
        peStatus = peExtractor.check_for_pe_success()
        if peStatus:
                peExtractor.extract_features()
                peExtractor.put_pe_features()
                peExtractor.put_rare_features()
		peExtractor.antiVM_detector(filename)
                peExtractor.createFeatureDict()
                peExtractor.pe_sections()
                peExtractor.pe_mapping_section()
                #print json.dumps(peExtractor.get_pe_features(),sort_keys=True, indent=4)
                #print json.dumps(peExtractor.get_rare_features(),sort_keys=True, indent=4)
		pe_features = {}
		pe_features['pe_features'] = peExtractor.get_pe_features()
		pe_features['pe_rare_features'] = peExtractor.get_rare_features()
		return pe_features
        else:
                logger.info('Not able to extract PE on file {} .. Exiting'.format(filename))


if __name__ == '__main__':
	main()
