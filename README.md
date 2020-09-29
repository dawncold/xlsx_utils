## How to use

### (optional) virtualenv

```
virtualenv venv
source venv/bin/activate
```

### install dependency

```
pip install openpyxl
```

### split xlsx

e.g. split `~/Downloads/large.xlsx` into 200 files, splitted files are stored nearby original file, within a new directory, `~/Downloads/large-splitted`
```
python splitxl.py ~/Downloads/large.xlsx 200
```

### merge xlsx

```
python mergexl.py /path/to/dir/contains/xlsx
```
