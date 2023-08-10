import tkinter as tk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import ImageTk, Image

def scrape_website(item):
    product_names = []
    product_prices = []
    product_ratings = []
    product_delivery = []
    product_no_of_ratings = []

    # selenium Chrome driver setup
    driver = webdriver.Chrome()
    driver.get("https://www.amazon.in")
    driver.maximize_window()

    # inserting the custom input from the user
    search_box = driver.find_element(By.ID, 'twotabsearchtextbox')
    search_box.clear()
    search_box.send_keys(item)

    # clicking the submit button
    driver.find_element(By.ID, "nav-search-submit-button").click()
    products = driver.find_elements(By.XPATH, '//div[@data-component-type="s-search-result"]')

    # iterating through each element
    for product in products:
        # product name
        names = product.find_elements(By.XPATH, ".//span[@class='a-size-medium a-color-base a-text-normal']")
        for item_name in names:
            product_names.append(item_name.text)

        # product price
        prices = product.find_elements(By.XPATH, ".//span[@class='a-price-whole']")
        for price in prices:
            product_prices.append(price.text)

        # product ratings
        ratings = product.find_elements(By.XPATH, ".//span[@class='a-icon-alt']")
        for rating in ratings:
            product_ratings.append(rating.text)

        # product delivery
        deliveries = product.find_elements(By.XPATH, ".//span[@class='a-color-base a-text-bold']")
        for delivery in deliveries:
            product_delivery.append(delivery.text)

        # product no of rating
        no_of_ratings = product.find_elements(By.XPATH, ".//span[@class='a-size-base s-underline-text']")
        for no_of_rate in no_of_ratings:
            product_no_of_ratings.append(no_of_rate.text)

        # displaying the details
        # for names, price, delivery in zip(product_names, product_prices, product_delivery):
        #     print(names, price, delivery)

        # inserting the data into the Excel file
        df = pd.DataFrame(zip(product_names, product_prices, product_ratings, product_no_of_ratings,  product_delivery),
                          columns=['Name', 'Price', 'Ratings', 'Total reviews', 'Delivery Timings'])
        df.to_excel("laptops.xlsx", index=False)

        driver.quit()


def submit_data():
    item = user_entry.get()
    print(item)
    scrape_website(item)

    # tk.messagebox.showinfo("Web Scrapping Result", result)


# creating the tkinter window
root = tk.Tk()
root.title("Web Scrapping App")
root.geometry("800x500")
root.resizable(False, False)
user_input = tk.StringVar()

# inserting the background image
frame = tk.Frame(root, width=600, height=400)
frame.pack()
frame.place(anchor='center', relx=0.5, rely=0.5)
img = ImageTk.PhotoImage(Image.open("./images/amazon-logo.jpg"))
label = tk.Label(frame, image = img)
label.pack()

# creating the widgets
# creating the Heading
tk.Label(root, text="Scrape Amazon Products Data", font=('calibre', 30), anchor="center").pack(pady=20)

# creating and asking the user input of the product
label = tk.Label(root, text="Enter Product Name", font=('calibre', 20))
label.pack()

# user input
user_entry = tk.Entry(root, textvariable=user_input, font=('calibre', 10, 'normal'))
user_entry.pack()
name = user_input.get()
print(name)

# submitting the input
submit_button = tk.Button(root, text="Fetch Data", command=submit_data)
submit_button.pack()

root.mainloop()
