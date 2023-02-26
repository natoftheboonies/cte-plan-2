from openpyxl import load_workbook
import json

status_sheet = "./VCodes2016-17.xlsx"


# route has vcode has cip
def main(sheet=status_sheet):
    wb = load_workbook(sheet)
    ws = wb.active
    max_row = ws.max_row
    routes = []
    route = None
    for row in range(1, max_row):
        if ws.cell(row, 1).value:
            # next route?
            if not ws.cell(row, 2).value:
                if not route:
                    route = {"name": ws.cell(row, 1).value}
                else:
                    route["type"] = ws.cell(row, 1).value
                    route["vcodes"] = []
                    routes.append(route)
                    route = None
            else:
                if ws.cell(row, 1).value == "CERTIFICATION CODE":
                    continue
                if ws.cell(row, 1).value:
                    vcode = {
                        "vcode": ws.cell(row, 1).value.strip(),
                        "certification": ws.cell(row, 2).value.strip(),
                        "experience": ws.cell(row, 3).value.strip(),
                        "cips": [],
                    }
                    routes[-1]["vcodes"].append(vcode)
        if ws.cell(row, 4).value:
            cip = {
                "code": ws.cell(row, 4).value,
                "name": ws.cell(row, 5).value.strip(),
            }
            vcode["cips"].append(cip)

    business = [route for route in routes if "Industry" in route["type"]]

    with open("vcode.json", "w") as fp:
        json.dump(business, fp)


# then run: python -m http.server
# and open http://localhost:8000/vcode.html

if __name__ == "__main__":
    main()
