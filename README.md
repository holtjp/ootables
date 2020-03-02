# Office Open Tables

Office Open Tables (ootables) provides an API for reading data in Excel tables
in Python. Eventually, this package plans to support tables in any Office Open
XML document, inlcuding Word and PowerPoint. Get started by cloning this repo
and running `python setup.py install` in your virtual environment. Here is an
example of basic usage:

```python
import ootables as oo


# Load the entire Excel file
b = oo.Book('Book1.xlsx')

# Access a sheet
s = b.sheets[0]
print(s.name)

# Access a table
t = s.tables[0]
print(t.name)
print(t.header)

# Access the table's data as a list of dictionaries
t.data
```

_James Holt, February 2020_
