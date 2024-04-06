import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"

@xw.func
def ret_test():
    return ["min","hwasoo", 100, 20, 30, 50]

@xw.func
def sum_and_product(number1, number2):
    sum_result = number1 + number2
    product_result = number1 * number2
    return sum_result, product_result

@xw.func
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    return 2 * (x + y)


if __name__ == "__main__":
    xw.Book("demo.xlsm").set_mock_caller()
    main()
