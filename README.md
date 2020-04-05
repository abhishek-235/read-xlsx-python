### read-xlsx-python
read multiple worksheets from xlsx file in python

#### Steps:
1. Create virtual environment
> virtualenv -p python3.6 venv
2. Activate virtual environment
> source venv/bin/activate
3. Install openpyxl
> pip install openpyxl
4. Execute read_xlsx.py
> python read_xlsx.py
5. Change source xlsx file name
> wb = load_workbook(filename = 'sample-ssn.xlsx')

#### You can get a list containing all worksheet names in workbook using "wb.sheetnames"

> print(wb.sheetnames)

> ['Sample SSN numbers', 'Sheet2', 'employement']


#### read_xlsx.py file shows How to:
1. Activate any worksheet by name
2. Get header/column titles
3. Read row wise data
4. Prepare/format data dictionary with header-title as key and row data as value

#### You can modify read_xlsx.py file as per your requirement
