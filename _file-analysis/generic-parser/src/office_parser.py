import sys
from oletools import olevba
 
class OfficeParser():
    def __init__(self,sample):
        self.sample = sample
        self.results = {}
 
    def extract_macro(self):
        vba = olevba.VBA_Parser(self.sample)
        macro_code = ""
 
        if vba.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
                macro_code += olevba.filter_vba(vba_code)
 
            self.results["analysis"] = vba.analyze_macros()
 
            self.results["code"] = macro_code
            vba.close()
            return self.results
 
        vba.close()
        return False
 
    def analysis(self):
        return self.extract_macro()
 
if __name__ == '__main__':
    obj = OfficeParser(sys.argv[1])
    results =  obj.analysis()
    for r in results["analysis"]:
        print r
    print "code: %s" % results["code"]
