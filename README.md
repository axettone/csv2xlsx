---
author: "Paolo Niccolò Giubelli <paoloniccolo.giubelli@gmail.com>"
date: "Last update on April 12, 2020"
---

# csv2xlsx

A simple python3 script to convert CSV files to XLSX

## Install

`pip3 install -r requirements.txt`

## Usage

`csv2xlsx.py [-h] [-d DELIMITER] csv_file`

Arguments

| Argument  | Explaination                              |
|-----------|-------------------------------------------|
| -h        | Show usage info                           |
| -d        | Specify the CSV delimiter (default is ',')|
| -csv_file | The csv file to convert                   |

This script will create a folder `xlsx` into the directory containing the `csv_file`
and will put the xlsx inside it.

## Batch usage

```bash
for f in *csv
do
    ./csv2xlsx.py $f
done
```
