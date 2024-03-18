import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
        sheet["A2"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"
        sheet["A2"].value = "Hello xlwings!"


@xw.func
def hello():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.range("B1").value = "32"



if __name__ == "__main__":
    xw.Book("demo.xlsm").set_mock_caller()
    main()
