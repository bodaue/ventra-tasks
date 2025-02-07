import smtplib
import time
from email.message import EmailMessage
from typing import List, Tuple, Any

import xlwings as xw
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager


def read_themes_from_excel(excel_path: str, sheet_name: str = "Sheet1") -> List[str]:
    wb: xw.Book = xw.Book(excel_path)
    ws: xw.main.Sheet = wb.sheets[sheet_name]
    used_range = ws.used_range
    last_row: int = used_range.last_cell.row
    last_col: int = used_range.last_cell.column
    headers: List[Any] = ws.range((1, 1), (1, last_col)).value
    try:
        theme_col_index: int = headers.index("Theme") + 1
    except ValueError as e:
        wb.close()
        raise ValueError("Не найден столбец 'Theme' в Excel.") from e
    themes_range = ws.range((2, theme_col_index), (last_row, theme_col_index)).value
    if isinstance(themes_range, str):
        themes_range = [themes_range]
    themes: List[str] = [t for t in themes_range if t and isinstance(t, str)]
    wb.close()
    return themes


def write_results_to_excel(excel_path: str, results: List[Tuple[str, str]], sheet_name: str = "Sheet1") -> None:
    wb: xw.Book = xw.Book(excel_path)
    ws: xw.main.Sheet = wb.sheets[sheet_name]
    used_range = ws.used_range
    last_row: int = used_range.last_cell.row
    headers: List[Any] = ws.range((1, 1), (1, used_range.last_cell.column)).value
    try:
        theme_col_index: int = headers.index("Theme") + 1
    except ValueError:
        theme_col_index = 1
    try:
        sources_col_index: int = headers.index("Sources") + 1
    except ValueError:
        sources_col_index = 2
    for theme, link in results:
        last_row += 1
        ws.range((last_row, theme_col_index)).value = theme
        ws.range((last_row, sources_col_index)).value = link
    wb.save()
    wb.close()


def search_in_yandex(themes: List[str]) -> List[Tuple[str, str]]:
    service = Service(ChromeDriverManager().install())
    driver: webdriver.Chrome = webdriver.Chrome(service=service)
    driver.maximize_window()
    all_results: List[Tuple[str, str]] = []
    for theme in themes:
        driver.get("https://ya.ru/")
        try:
            search_box = driver.find_element(By.ID, "text")
        except NoSuchElementException:
            print(f"Элемент поиска не найден для темы '{theme}', пропускаем её.")
            continue
        search_box.send_keys(theme)
        search_box.send_keys(Keys.ENTER)
        time.sleep(2)
        links = driver.find_elements(By.XPATH, "//a[@href and contains(@class,'OrganicTitle-Link')]")
        top_links: List[str] = [link_el.get_attribute("href") for link_el in links[:3]]
        for ln in top_links:
            all_results.append((theme, ln))
    driver.quit()
    return all_results


def send_email_with_attachment(
        file_path: str,
        subject: str,
        body_text: str,
        from_addr: str,
        to_addr: str,
        login: str,
        password: str
) -> None:
    msg: EmailMessage = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg.set_content(body_text)
    with open(file_path, "rb") as f:
        file_data = f.read()
    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="octet-stream",
        filename=file_path.split("\\")[-1]
    )
    with smtplib.SMTP_SSL("smtp.yandex.ru", 465) as server:
        server.login(login, password)
        server.send_message(msg)
    print("Письмо успешно отправлено.")


def main() -> None:
    excel_path: str = r"TestTask2.xlsx"
    sheet_name: str = "Sheet1"
    themes: List[str] = read_themes_from_excel(excel_path, sheet_name)
    if not themes:
        print("Не найдено ни одной темы в Excel.")
        return
    results: List[Tuple[str, str]] = search_in_yandex(themes)
    if not results:
        print("Не удалось найти ссылки по данным темам.")
    else:
        write_results_to_excel(excel_path, results, sheet_name)
    send_email_with_attachment(
        file_path=excel_path,
        subject="Список тем для доклада",
        body_text="Файл с ссылками.",
        from_addr="your@yandex.ru",
        to_addr="recipient@yandex.ru",
        login="your@yandex.ru",
        password="your_app_password"
    )
    print("Готово. Скрипт завершён.")


if __name__ == "__main__":
    main()
