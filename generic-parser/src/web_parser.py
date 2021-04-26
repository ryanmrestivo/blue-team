from html.parser import HTMLParser 
import sys
import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
#handler = logging.FileHandler('logs/generic.log')
#handler.setLevel(logging.INFO)
#formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
#handler.setFormatter(formatter)
#logger.addHandler(handler)


class MyHTMLParser(HTMLParser):
	

    def handle_starttag(self, tag, attrs):
	meta = {}
	meta['start_tag'] = []
        print "Start tag:", tag
	meta['start_tag'].append(tag)
        for attr in attrs:
	    meta['start_tag'][tag] = []
	    meta['start_tag'][tag].append(attr)	    
            print("     attr:", attr)
	print meta
    def handle_endtag(self, tag):
        print("End tag  :", tag)

    def handle_data(self, data):
        print("Data     :", data)

    def handle_comment(self, data):
        print("Comment  :", data)

    def handle_entityref(self, name):
        c = chr(name2codepoint[name])
        print("Named ent:", c)

    def handle_charref(self, name):
        if name.startswith('x'):
            c = chr(int(name[1:], 16))
        else:
            c = chr(int(name))
        print("Num ent  :", c)

    def handle_decl(self, data):
        print("Decl     :", data)
    
  
class WebParser():
	"""
	Find if any hidden exploits or macros hidden
	"""
	def __init__(self):
		self.html_meta = {}
		self.swf_meta = {}
		self.jar_meta = {}
	

	def html_parser(self,file_path):
		"""
		Function to parse html file and find JS,swf traces
	
		Args: File_path
		
		Return : Dictionary html_features 
		"""	
		self.html_meta['javascript'] = {}
		self.html_meta['images'] = 0
		self.html_meta['tags'] = 0
		self.html_meta['no_of_words'] = 0
		self.html_meta['decleration'] = False
		try:
			with open(file_path) as f:
				html_data = f.read()
				parser = MyHTMLParser()
				parser.feed(html_data)
				
		except Exception as e:
			print str(e)
			#logger.error('Error : Parsing HTML FILE {}'.format(str(e)))		
def main():
	web = WebParser()
	web.html_parser(sys.argv[1])
if __name__ :
	main()
