import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

# Set up Excel workbook for test documentation
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Test Cases"
sheet.append(["Test Case", "Description", "Result"])

# Set up WebDriver
driver = webdriver.Chrome()  # Ensure chromedriver is in PATH
driver.maximize_window()
driver.implicitly_wait(10)

# Base URL
url = "https://demo.nopcommerce.com/"

try:
    driver.get(url)

    # Test Case 1: User Registration
    sheet.append(["TC01", "User Registration", "Pending"])
    driver.find_element(By.LINK_TEXT, "Register").click()
    driver.find_element(By.ID, "gender-male").click()
    driver.find_element(By.ID, "FirstName").send_keys("John")
    driver.find_element(By.ID, "LastName").send_keys("Doe")
    driver.find_element(By.ID, "Email").send_keys("testuser123@example.com")
    driver.find_element(By.ID, "Password").send_keys("Test@1234")
    driver.find_element(By.ID, "ConfirmPassword").send_keys("Test@1234")
    driver.find_element(By.ID, "register-button").click()

    # Validate Registration
    success_message = driver.find_element(By.CLASS_NAME, "result").text
    if "Your registration completed" in success_message:
        sheet.append(["TC01", "User Registration", "Passed"])
    else:
        sheet.append(["TC01", "User Registration", "Failed"])

    time.sleep(2)

    # Test Case 2: User Login
    sheet.append(["TC02", "User Login", "Pending"])
    driver.find_element(By.LINK_TEXT, "Log in").click()
    driver.find_element(By.ID, "Email").send_keys("testuser123@example.com")
    driver.find_element(By.ID, "Password").send_keys("Test@1234")
    driver.find_element(By.CLASS_NAME, "login-button").click()

    # Validate Login
    if driver.find_element(By.CLASS_NAME, "account").is_displayed():
        sheet.append(["TC02", "User Login", "Passed"])
    else:
        sheet.append(["TC02", "User Login", "Failed"])

    time.sleep(2)

    # Test Case 3: Product Search and Browsing
    sheet.append(["TC03", "Product Search", "Pending"])
    search_box = driver.find_element(By.ID, "small-searchterms")
    search_box.send_keys("laptop")
    search_box.send_keys(Keys.RETURN)

    # Validate Product Search
    search_results = driver.find_elements(By.CLASS_NAME, "product-item")
    if len(search_results) > 0:
        sheet.append(["TC03", "Product Search", "Passed"])
    else:
        sheet.append(["TC03", "Product Search", "Failed"])

    time.sleep(2)

    # Test Case 4: Add to Cart
    sheet.append(["TC04", "Add to Cart", "Pending"])
    driver.find_elements(By.CLASS_NAME, "product-item")[0].click()
    driver.find_element(By.ID, "add-to-cart-button-1").click()

    # Validate Add to Cart
    cart_message = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.CLASS_NAME, "content"))
    ).text
    if "The product has been added to your shopping cart" in cart_message:
        sheet.append(["TC04", "Add to Cart", "Passed"])
    else:
        sheet.append(["TC04", "Add to Cart", "Failed"])

    time.sleep(2)

    # Test Case 5: Checkout
    sheet.append(["TC05", "Checkout", "Pending"])
    driver.find_element(By.CLASS_NAME, "ico-cart").click()
    driver.find_element(By.ID, "checkout").click()

    # Simulate checkout process
    driver.find_element(By.ID, "BillingNewAddress_FirstName").send_keys("John")
    driver.find_element(By.ID, "BillingNewAddress_LastName").send_keys("Doe")
    driver.find_element(By.ID, "BillingNewAddress_Email").send_keys("testuser123@example.com")
    driver.find_element(By.ID, "BillingNewAddress_CountryId").send_keys("United States")
    driver.find_element(By.ID, "BillingNewAddress_City").send_keys("New York")
    driver.find_element(By.ID, "BillingNewAddress_Address1").send_keys("123 Test St")
    driver.find_element(By.ID, "BillingNewAddress_ZipPostalCode").send_keys("10001")
    driver.find_element(By.ID, "BillingNewAddress_PhoneNumber").send_keys("1234567890")
    driver.find_element(By.CLASS_NAME, "new-address-next-step-button").click()

    # Validate Checkout
    try:
        order_confirmation = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "title"))
        ).text
        if "Thank you" in order_confirmation:
            sheet.append(["TC05", "Checkout", "Passed"])
        else:
            sheet.append(["TC05", "Checkout", "Failed"])
    except:
        sheet.append(["TC05", "Checkout", "Failed"])

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Save results to Excel
    workbook.save("Test_Cases_nopCommerce.xlsx")
    driver.quit()
