# CTE Plan 2

Well, currently just a V-Code transformer.

Coming from industry? What will you teach? Certain classes fall in multiple V-Codes.
These scripts transform the [OSPI V-Code Chart xlsx](https://www.k12.wa.us/sites/default/files/public/careerteched/pubdocs/VCodes2016-17.xlsx) into a list of classes which fall in multiple V-Codes.

For example,

- Baking and Pastry Arts/Baker/Pastry Chef: 120501
  - V120505: Food Production and Services
  - V200493: Culinary Arts

run `vcode.py` to generate `vcode.json` from the spreadsheet.

Note: Depends on [openpyxl](https://pypi.org/project/openpyxl/) to read the spreadsheet.

And then since you have python:

1. run: `python -m http.server`
2. visit: `open http://localhost:8000/vcode.html` to see 'em all.
