#!/usr/bin/python
import pefile
import logging
import sys
import hashlib
import json
from pdf import *

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
handler = logging.FileHandler('logs/generic.log')
handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class PDFFeatureExtractor():
        def __init__(self , filepath):
                self.filepath = filepath
		self.pdf_features = {}
		self.verbose = True
		self.extract = False	
		self.comments = []
		self.xref = []
                self.trailer = []
                self.startXref = []
                self.indirectObjects = []

		self.yara = 0
		self.generateembedded = 0
		self.raw = 1
		self.debug = 1
		self.generate = 1
	def extract_stats(self):
		self.decoders = []
		#LoadDecoders(self.decoders, True)
		selectComment = False
        	selectXref = False
        	selectTrailer = False
        	selectStartXref = False
        	selectIndirectObject = False
		oPDFParser = cPDFParser(self.filepath, self.verbose, self.extract)
		self.object = oPDFParser.GetObject()
		cntComment = 0
        	cntXref = 0
        	cntTrailer = 0
        	cntStartXref = 0
        	cntIndirectObject = 0
        	dicObjectTypes = {}

		if self.object.type == PDF_ELEMENT_COMMENT:
        	        cntComment += 1
                elif self.object.type == PDF_ELEMENT_XREF:
                        cntXref += 1
                elif self.object.type == PDF_ELEMENT_TRAILER:
                        cntTrailer += 1
                elif self.object.type == PDF_ELEMENT_STARTXREF:
                        cntStartXref += 1
                elif self.object.type == PDF_ELEMENT_INDIRECT_OBJECT:
                        cntIndirectObject += 1
                        type1 = self.object.GatType()
                	if not type1 in dicObjectTypes:
                		dicObjectTypes[type1] = [self.object.id]
                	else:
                		dicObjectTypes[type1].append(self.object.id)	
        	self.pdf_features['comment'] = cntComment
		self.pdf_features['xref'] = cntXref
		self.pdf_features['trailer'] = cntTrailer
		self.pdf_features['start_xref'] = cntStartXref
		self.pdf_features['indirect_obj'] = cntIndirectObject
		names = dicObjectTypes.keys()
        	names.sort()
		self.pdf_features['names'] = names
		for key in names:
        		print(' %s %d: %s' % (key, len(dicObjectTypes[key]), ', '.join(map(lambda x: '%d' % x, dicObjectTypes[key]))))
	def get_pdf_features(self):
		self.pdf_features['comments'] = self.comments
		self.pdf_features['xreg'] = self.xref
		self.pdf_features['trailer'] = self.trailer
		self.pdf_features['startXref'] = self.startXref
		self.pdf_features['indirectObjects'] = self.indirectObjects
		self.pdf_f = {}
		self.pdf_f['pdf_features'] = self.pdf_features
		
		return self.pdf_f
	
	def pdf_extract_elements(self):
		self.get_comments()
		self.get_xref()
		self.get_trailer()
		self.get_startXref()
		self.get_indirectObjects()
	def get_comments(self):
		if self.object.type == PDF_ELEMENT_COMMENT:
                        if self.generate:
                            comment = self.object.comment[1:].rstrip()
                            if re.match('PDF-\d\.\d', comment):
                                #print("    oPDF.header('%s')" % comment[4:])
				self.comments.append("oPDF.header({})".format(comment[4:]))
                            elif comment != '%EOF':
                                #print('    oPDF.comment(%s)' % repr(comment))
				self.comments.append("oPDF.comment({})".format(repr(comment)))
                        elif self.yara == None and self.generateembedded == 0:
                            print('PDF Comment %s' % FormatOutput(self.object.comment,1))
                            print('')


	def get_xref(self):
		if self.object.type == PDF_ELEMENT_XREF:
                        if self.generate and self.yara == None and self.generateembedded == 0:
                            if self.debug:
				self.xref.append('xref {}'.format(self.object.content))
                                #print('xref %s' % FormatOutput(self.object.content, self.raw))
                            else:
				pass
                                #print('xref')
                            #print('')

	def get_trailer(self):
		if self.object.type == PDF_ELEMENT_TRAILER:
                        oPDFParseDictionary = cPDFParseDictionary(self.object.content[1:], self.nocanonicalizedoutput)
                        if self.generate:
                            result = oPDFParseDictionary.Get('/Root')
                            if result != None:
                                savedRoot = result
                        elif self.yara == None and self.generateembedded == 0:
                            if not self.search and not self.key or self.search and self.object.Contains(self.search):
                                if oPDFParseDictionary == None:
                                    print('trailer %s' % FormatOutput(object.content, self.raw))
                                else:
                                    print('trailer')
                                    oPDFParseDictionary.PrettyPrint('  ')
                                print('')
                            elif self.key:
                                if oPDFParseDictionary.parsed != None:
                                    result = oPDFParseDictionary.GetNested(self.key)
                                    if result != None:
                                        print(result)

	def get_startXref(self):
		pass
	def get_indirectObjects(self):
		pass

def pdf_test(filepath):
	pdf_parse_obj = PDFFeatureExtractor(filepath)
	pdf_parse_obj.extract_stats()
	pdf_parse_obj.pdf_extract_elements()
	#print json.dumps(pdf_parse_obj.get_pdf_features(),sort_keys=True,indent=4)
	return pdf_parse_obj.get_pdf_features()
def main():
	filepath = sys.argv[1]
	

if __name__ == '__main__':
	main()

