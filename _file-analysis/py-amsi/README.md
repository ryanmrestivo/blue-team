# py-amsi

py-amsi is a library that scans strings or files for malware using the Windows
Antimalware Scan Interface (AMSI) API. AMSI is an interface native to Windows 
that allows applications to ask the antivirus installed on the system
to analyse a file/string. AMSI is not tied to Windows Defender. Antivirus
providers implement the AMSI interface to receive calls from applications.
This library takes advantage of the API to make antivirus scans in python.
Read more about the Windows AMSI API [here](https://learn.microsoft.com/en-us/windows/win32/amsi/antimalware-scan-interface-portal).

## Installation
- Via pip
  
  ```
  pip install pyamsi
  ```
- Clone repository

  ```bash
  git clone https://github.com/Tomiwa-Ot/py-amsi.git
  cd py-amsi/
  python setup.py install
  ```

## Usage
```python
from pyamsi import Amsi

# Scan a file
Amsi.scan_file(file_path, debug=True) # debug is optional and False by default

# Scan string
Amsi.scan_string(string, string_name, debug=False) # debug is optional and False by default

# Both functions return a dictionary of the format
# {
#     'Sample Size' : 68,         // The string/file size in bytes
#     'Risk Level' : 0,           // The risk level as suggested by the antivirus
#     'Message' : 'File is clean' // Response message
# }
```

<table>
    <tr>
        <th>Risk Level</th>
        <th>Meaning</th>
    </tr>
    <tr>
        <td>0</td>
        <td>AMSI_RESULT_CLEAN (File is clean)</td>
    </tr>
    <tr>
        <td>1</td>
        <td>AMSI_RESULT_NOT_DETECTED (No threat detected)</td>
    </tr>
    <tr>
        <td>16384</td>
        <td>AMSI_RESULT_BLOCKED_BY_ADMIN_START (Threat is blocked by the administrator)</td>
    </tr>
    <tr>
        <td>20479</td>
        <td>AMSI_RESULT_BLOCKED_BY_ADMIN_END (Threat is blocked by the administrator)</td>
    </tr>
    <tr>
        <td>32768</td>
        <td>AMSI_RESULT_DETECTED (File is considered malware)</td>
    </tr>
</table>

## Docs
https://tomiwa-ot.github.io/py-amsi/index.html
