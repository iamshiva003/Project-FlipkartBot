import os
import sys
import time
import tkinter as tk

import openpyxl
import pandas as pd
from PIL import ImageTk, Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait


# Get absolute path to resource, works for dev and for PyInstaller
# noinspection PyProtectedMember
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    # noinspection PyBroadException
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def get_data(driver, item):
    product_names = []
    product_prices = []
    product_overall_ratings = []
    product_no_of_ratings = []
    product_no_of_reviews = []

    # getting the total no of pages
    wait = WebDriverWait(driver, 10)
    no_of_pages = wait.until(
        ec.element_to_be_clickable((By.XPATH, "// span[contains(text(),\'Page 1 of')]"))).text

    # no_of_pages = driver.find_element(By.XPATH, "// span[contains(text(),\'Page 1 of')]").text
    no_of_pages = [int(s) for s in no_of_pages.split() if s.isdigit()][1]
    print(f"Total no of pages: {no_of_pages}")

    wait = WebDriverWait(driver, 10)
    next_page = wait.until(ec.presence_of_all_elements_located((By.XPATH, '//a[@class="_1LKTO3"]')))[
        -1]
    # next_page = driver.find_elements(By.XPATH, '//a[@class="_1LKTO3"]')[-1]
    print("Got next page")

    for page in range(no_of_pages):
        time.sleep(4)
        # names
        wait = WebDriverWait(driver, 10)
        names = wait.until(ec.presence_of_all_elements_located((By.XPATH, '//div[@class="_4rR01T"]')))
        # names = driver.find_elements(By.XPATH, '//div[@class="_4rR01T"]')
        for item_name in names:
            product_names.append(item_name.text)

        # prices
        wait = WebDriverWait(driver, 10)
        prices = wait.until(
            ec.presence_of_all_elements_located((By.XPATH, '//div[@class="_30jeq3 _1_WHN1"]')))
        # prices = driver.find_elements(By.XPATH, '//div[@class="_30jeq3 _1_WHN1"]')
        for price in prices:
            product_prices.append(price.text)

        # overall ratings
        wait = WebDriverWait(driver, 10)
        overall_ratings = wait.until(
            ec.presence_of_all_elements_located((By.XPATH, '//div[@class="_3LWZlK"]')))
        # overall_ratings = driver.find_elements(By.XPATH, '//div[@class="_3LWZlK"]')
        for overall_rating in overall_ratings:
            product_overall_ratings.append(overall_rating.text)

        # ratings
        wait = WebDriverWait(driver, 10)
        ratings = wait.until(
            ec.presence_of_all_elements_located((By.XPATH, "// span[contains(text(),\'Ratings')]")))
        # ratings = driver.find_elements(By.XPATH, "// span[contains(text(),\'Ratings')]")
        for rating in ratings:
            product_no_of_ratings.append(rating.text)

        # reviews
        wait = WebDriverWait(driver, 10)
        reviews = wait.until(
            ec.presence_of_all_elements_located((By.XPATH, "// span[contains(text(),\'Reviews')]")))
        # reviews = driver.find_elements(By.XPATH, "// span[contains(text(),\'Reviews')]")
        for review in reviews:
            product_no_of_reviews.append(review.text)

        # inserting the data into the Excel file
        df = pd.DataFrame(columns=['Name', 'Price', 'Overall Rating out of 5', 'No.of Ratings', 'No.of Reviews'])
        df.to_excel(resource_path(f"data\\{item}.xlsx"), index=False)
        path = f"data\\{item}.xlsx"
        workbook = openpyxl.load_workbook(resource_path(path))
        sheet = workbook.active
        for pname, price, o_rating, no_rating, review in zip(product_names, product_prices, product_overall_ratings,
                                                             product_no_of_ratings, product_no_of_reviews):
            row_values = [pname, price, o_rating, no_rating, review]
            sheet.append(row_values)
            workbook.save(path)

        print(f"Page {page + 1} loaded")
        if page == no_of_pages - 1:
            break
        # time.sleep(3)
        next_page.click()


def scrape_website(item):
    # selenium Chrome driver setup
    driver = webdriver.Chrome()
    driver.get("https://www.flipkart.com")
    driver.maximize_window()

    time.sleep(5)
    search_box = driver.find_element(By.NAME, 'q')
    search_box.clear()
    search_box.send_keys(item)
    search_box.send_keys(Keys.RETURN)

    time.sleep(5)

    get_data(driver, item)

    time.sleep(5)
    driver.quit()


def submit_data():
    item = user_entry.get()
    print(item)
    scrape_website(item)


# The main function begins from here
# creating the tkinter window
root = tk.Tk()
root.title("Web Scrapping App")
root.geometry("1000x674")
root.resizable(False, False)
user_input = tk.StringVar()

# inserting the background image
frame = tk.Frame(root, width=600, height=400)
frame.pack()
frame.place(anchor='center', relx=0.5, rely=0.5)
img = ImageTk.PhotoImage(Image.open(resource_path("images\\logo.jpg")))
label = tk.Label(frame, image=img)
label.pack()

# creating the widgets
# creating the Heading
heading = tk.Label(root, text="Get Flipkart Products Data", font=('calibre', 30), anchor="center", bg="#007BD8")
heading.pack(pady=15)

# creating and asking the user input of the product
label = tk.Label(root, text="Enter Product Name", font=('calibre', 20), bg="#007BD8")
label.pack(pady=5)

# user input
user_entry = tk.Entry(root, textvariable=user_input, font=('calibre', 10, 'normal'))
user_entry.pack(pady=5)
name = user_input.get()
print(name)
user_entry.focus()

# submitting the input
submit_button = tk.Button(root, text="Fetch Data", command=submit_data, height=2, width=10, fg="#007BD8")
submit_button.pack(pady=5)

root.mainloop()
