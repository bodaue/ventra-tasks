import xlwings as xw


def main() -> None:
    try:
        wb = xw.books["TestTask1.xlsx"]
    except KeyError:
        print("Файл 'TestTask1.xlsx' не открыт в Excel.")
        return

    ws = wb.sheets["Sheet1"]
    used_range = ws.used_range
    last_row = used_range.last_cell.row
    last_col = used_range.last_cell.column

    headers = ws.range((1, 1), (1, last_col)).value

    try:
        status_col_index = headers.index("Status") + 1
    except ValueError:
        print("Столбец 'Status' не найден.")
        return

    for row in range(2, last_row + 1):
        status_value = ws.range((row, status_col_index)).value
        if status_value == "Done":
            color = (0, 255, 0)  # Зеленый
        elif status_value == "In progress":
            color = (255, 0, 0)  # Красный
        else:
            continue

        ws.range((row, 1), (row, last_col)).color = color

    wb.save()
    print("Файл успешно сохранён.")


if __name__ == "__main__":
    main()
