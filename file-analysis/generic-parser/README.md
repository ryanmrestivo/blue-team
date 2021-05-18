# Generic Parser for Analyzing Malware Files to Detect Suspicious Behaviour.

A Single Library Parser to extract meta information,static analysis and detect macros within the files.

# Usage:

## PreRequsite
1. Clone the Repo
2. Create a virutalenv
```
virtualenv pyenv
```
3. Install the requirements.
```
pip install -r requirements.txt
```
### Script Usage

```
(pyenv) admin@cuckoo:~/generic-parser$ python app.py -h
usage: app.py [-h] -f PATH [-s STORE] -y YARA -e EXTRACT [--version]

optional arguments:
  -h, --help            show this help message and exit
  -f PATH, --path PATH  File Absolute Path
  -s STORE, --store STORE
                        Store to DB
  -y YARA, --yara YARA  Apply Yara Matcher
  -e EXTRACT, --extract EXTRACT
                        Extract Features
  --version             show program's version number and exit

```
1. PATH  : This should point to the path of the malware file which you want to analyze.
2. STORE : Enable this flag if you want to store in a database.
3. YARA  : Enable this flag to apply yara to match for suspicious indicators in the file.
4. version : Shows the version of the tool.

### Features:

1. Ability to Identify the Decomposition module selected based on the mime-type.
2. Apply PDF based decomposition to extract features from the pdf file.
3. Apply Office based decomposition to extract features of office files.
4. Web Based files are decomposed to get interesting strings etc.
5. Yara is applied on the entire file to get interesting matches which can help in identifying suspicious behaviour.

### Sample UseCases

 - Please refer to [USECASES](USECASES.md)