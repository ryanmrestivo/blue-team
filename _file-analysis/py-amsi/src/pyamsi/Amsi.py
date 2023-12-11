#!/usr/bin/env python3

"""
Name: py-amsi (v1.11)
Description: Scan strings and files using the windows antimalware interface
Author: Olorunfemi-Ojo Tomiwa
URL: https://github.com/Tomiwa-Ot/py-amsi
License: MIT
Copyright (c) @Tomiwa-Ot 2022
"""

import os
import ctypes

# AMSI return codes & their meanings
STATUS = {
    'AMSI_RESULT_CLEAN' : 0,
    'AMSI_RESULT_NOT_DETECTED' : 1,
    'AMSI_RESULT_BLOCKED_BY_ADMIN_START' : 16384,
    'AMSI_RESULT_BLOCKED_BY_ADMIN_END' : 20479,
    'AMSI_RESULT_DETECTED' :  32768
}

# Further description of AMSI return codes
MESSAGE = {
    STATUS['AMSI_RESULT_CLEAN'] : 'File is clean',
    STATUS['AMSI_RESULT_NOT_DETECTED'] : 'No threat detected',
    STATUS['AMSI_RESULT_BLOCKED_BY_ADMIN_START'] : 'Threat is blocked by the administrator',
    STATUS['AMSI_RESULT_BLOCKED_BY_ADMIN_END'] : 'Threat is blocked by the administrator',
    STATUS['AMSI_RESULT_DETECTED'] : 'File is considered malware',
    5 : 'N/A'
}

# Load DLL to access AMSI functions
dll = ctypes.CDLL(os.path.join(
    os.path.dirname(__file__),
    'amsiscanner.dll'
))


def scan_file(path, debug=False):
    """
    Scans a buffer-full content for malware

    Arguments
        path (str):   path of the file to be scanned
        debug (bool): show debug messages (default=False, optional)
    
    Returns
        dictionary : 
            Sample Size = String size
            Risk Level  = Risk level stated by the AMSI provider
            Message     = AMSI response
    """
    if not os.path.exists(path):
        raise Exception(f'No such file: {path}')
    result = dll.scanBytes(path, os.path.basename(path), int(debug))
    file = open(path, 'rb')
    data = file.read()
    file.close
    return {
        "Sample size" : len(bytearray(data)),
        "Risk Level" : result,
        "Message" : MESSAGE[result]
    }
    

def scan_string(text, name, debug=False):
    """
    Scans a string for malware

    Arguments
        text (str):   string to be scanned
        name (str):   a name for the string
        debug (bool): show debug messages (default=False, optional)
    
    Returns
        dictionary : 
            Sample Size = String size
            Risk Level  = Risk level stated by the AMSI provider
            Message     = AMSI response
    """
    result = dll.scanString(text, name, int(debug))
    return {
        "Sample size" : len(text),
        "Risk Level" : result,
        "Message" : MESSAGE[result]
    }
