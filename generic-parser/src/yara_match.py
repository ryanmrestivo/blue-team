#!/usr/bin/env python
# yara_scan
# usage: python yara_scan.py -y <yara_rule_dir> [-s <scan_files_dir> (optional otherwise current dir is scanned)]

import os
import sys
import argparse
import zipfile
import yara
import shutil
import subprocess
import magic
from oletools.olevba import VBA_Parser
import rarfile
import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# create a file handler

handler = logging.FileHandler('logs/generic.log')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

def parse_arguments():
    """
    Argument parser requires a yara rule folder and optionally a scan directory otherwise the current
    directory is used.  Verbose option will print invalid rules that are not scanned with.
    """
    parser = argparse.ArgumentParser(usage="Scan Files in a Directory with Yara Rules")
    parser.add_argument("-v", "--verbosity", action="store_true", dest="verbose",
                        help="Print verbose information")
    parser.add_argument('-y', '--yara_dir',
                        action='store',
                        help='Path to Yara rules directory')

    parser.add_argument('-s', '--scan_dir',
                        action='store',
                        default=os.getcwd(),
                        help='Path to the directory of files to scan (optional otherwise current dir is scanned)')
    parser.add_argument('-f', '--scan_file',
                        action='store',
                        help='Path to file scan')
    return parser

class YaraClass:
    """
    Main Yara Class that handles walking rule dir, compiling and testing rules, and walking and scanning files.
    """
    def __init__(self, arg_yara_dir, arg_scan_dir, arg_verbose , arg_file):
        """
        YaraClass initialization that sets verbose, scan and yara directory
        """
        try:
            self.verbose = arg_verbose
            self.scan_dir = arg_scan_dir
            self.yara_dir = arg_yara_dir
            self.file_path = arg_file
        except Exception as e:
            logger.error("Init Exception: {}".format(e))

    def compile(self):
        """
        Walks rule dir, tests rules, and compiles them for scanning.
        """
        try:
            logger.info("Compiling rules from {}".format(self.yara_dir))
            all_rules = {}
            for root, directories, files in os.walk(self.yara_dir):
                for file in files:
                    logger.info('Adding Yara file {} to compile list'.format(file))
                    if "yar" in os.path.splitext(file)[1]:
                        rule_case = os.path.join(root,file)
                        if self.test_rule(rule_case):
                            all_rules[file] = rule_case
            self.rules = yara.compile(filepaths=all_rules)
        except Exception as e:
            logger.error("Compile Exception: {}".format(e))

    def test_rule(self, test_case):
        """
        Tests rules to make sure they are valid before using them.  If verbose is set will print the invalid rules.
        """
        try:
            testit = yara.compile(filepath=test_case)
            return True
        except:
            if self.verbose:
            	logger.error("{} is an invalid rule".format(test_case))
            return False

    def scan_all(self):
        """
        Scan all method that recursively walks the directory and calls scan and unpack
        """
        try:
            for root, directories, files in os.walk(self.scan_dir):
                for file in files:
                    work_file = os.path.join(root,file)
                    self.scan(work_file)
                    self.check_unpack(work_file)
        except Exception as e:
            logger.error("Scan Exception: {} for file {}".format(e, work_file))
    def scan_single(self, scan_file):
        """
        Scan method to scan a user submitted file
        """
        try:
            work_file = scan_file
            result_obtained = self.scan(work_file)
            return result_obtained
            #self.check_unpack(work_file)
        except Exception as e:
            return []
            logger.error('Scan Exception {} for file {}'.format(e, work_file))

    def scan(self, scan_file):
        """
        Scan method that uses compiled rules to scan a file
        """
        try:
            matches = self.rules.match(scan_file)
            return matches
            #print "{}\n{}\n".format(scan_file, matches)

        except Exception as e:
            return []
            logger.error("Scan Exception: {}".format(e))

    def check_unpack(self, work_file):
        fc = self.FileClass(work_file)
        fc.check_magic()
        if fc.zip:
            fc.unzip()
        if fc.doc:
            fc.get_macro()
        if fc.rar:
            fc.unrar_file()
        if fc.pe:
            fc.pe_unpack()
        if (
            fc.zip or
            fc.doc or
            fc.rar or
            fc.pe
        ):
            try:
                for root, directories, files in os.walk(fc.tmp_dir):
                    for file in files:
                        check_unpack_file = os.path.join(root,file)
                        self.scan(check_unpack_file)
                fc.rm_tmp_dir()
            except Exception as e:
                logger.error("Check Unpack Loop Exception: {}".format(e))


    class FileClass:
        """
        Subclass that identifies file types and extracts archives and macros
        """
        def __init__(self, file):
            self.file = file
            self.tmp_dir = "{}_tmp".format(self.file)
            self.doc = False
            self.zip = False
            self.rar = False
            self.pe = False

        def mk_tmp_dir(self):
            if not os.path.exists(self.tmp_dir):
                os.makedirs(self.tmp_dir)

        def rm_tmp_dir(self):
            if os.path.exists(self.tmp_dir):
                shutil.rmtree(self.tmp_dir, ignore_errors=True)

        def check_magic(self):
            try:
                file_type = magic.from_file(self.file)
                #print file_type
                if (
                    "Composite Document File" in file_type or
                    "Word 2007+" in file_type or
                    "Excel 2007+" in file_type or
                    "PowerPoint 2007+" in file_type or
                    "Rich Text Format data" in file_type
                     ):
                    self.zip = True
                    self.doc = True
                if (
                    "Java" in file_type or
                    "Macromedia Flash data" in file_type or
                    "Zip" in file_type
                    ):
                    self.zip = True
                if "RAR" in file_type:
                    self.rar = True
                if "PE" in file_type:
                    self.pe = True

            except Exception as e:
                logger.error("Check Magic Exception: {}".format(e))

        def unzip(self):
            try:
                self.mk_tmp_dir()
                logger.info("Unzipping {}".format(self.file))
                with zipfile.ZipFile(self.file) as zf_file:
                    zf_file.setpassword("infected")
                    zf_file.extractall(self.tmp_dir)
            except Exception as e:
                logger.error("Unzip Exception: {}".format(e))

        def get_macro(self):
            """
            Get Macros from an Office file and write them to a text file
            """
            try:
                self.mk_tmp_dir()
                logger.info("Getting Macros from {}".format(self.file))
                vb = VBA_Parser(self.file, relaxed=True)
                if vb.detect_vba_macros():
                    with open("{}{}macros.txt".format(self.tmp_dir, os.sep), "w") as macro_file:
                        for (subfilename, stream_path, vba_filename, vba_code) in vb.extract_all_macros():
                            macro_file.write(vba_code)
            except Exception as e:
                logger.error("get_macro Exception: {}".format(e))

        def unrar_file(self):
            """
            Unrar files into tmp directory
            """
            try:
                self.mk_tmp_dir()
                logger.info("Unraring file from {}".format(self.file))
                rf = rarfile.RarFile(self.file)
                rf.extractall(self.tmp_dir)
            except Exception as e:
                logger.error("unrar Exception: {}".format(e))

        def pe_unpack(self):
            """
            PE Unpacking actions
            """
            try:
                self.mk_tmp_dir()
                logger.info("Unpacking PE from {}".format(self.file))
                subprocess.call(["upx", "-d", "{}".format(self.file), "-o", "{}{}upx_unpacked".format(self.tmp_dir, os.sep)], stderr=subprocess.STDOUT)
            except Exception as e:
                logger.error("PE Unpacking Exception {}".format(e))

def main():
    args = parse_arguments().parse_args()
    ys = YaraClass(args.yara_dir, args.scan_dir, args.verbose , args.scan_file)
    ys.compile()
    #ys.scan_all()
    ys.scan_single(args.scan_file)

if __name__ == "__main__":
    main()

