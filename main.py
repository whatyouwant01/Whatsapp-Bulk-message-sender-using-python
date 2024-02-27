import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
def preprocess_numbers(numbers_text):
    processed_numbers = []
    for line in numbers_text.split('\n'):
        line = line.strip()
        if line:
            numbers = line.split()
            processed_numbers.extend(numbers)
    return processed_numbers

def send_messages():
    numbers_text = numbers_entry.get("1.0", "end")
    message = message_entry.get("1.0", "end").strip()

    if not numbers_text or not message:
        messagebox.showerror("Error", "Please enter numbers and message")
        return

    numbers = preprocess_numbers(numbers_text)

    driver = None
    try:
        driver = webdriver.Chrome()
        driver.get('https://web.whatsapp.com')
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, "//div[@title='New chat']")))

        batch_size = 30
        pause_duration = 60

        for i in range(0, len(numbers), batch_size):
            batch_numbers = numbers[i:i+batch_size]

            for num in batch_numbers:
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@title='New chat']")))
                    new_chat_button = driver.find_element(By.XPATH, "//div[@title='New chat']")
                    new_chat_button.click()

                    search_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @role='textbox' and @title='Search input textbox']")))
                    search_input.send_keys(num)
                    time.sleep(0.5)

                    loading = driver.find_elements(By.XPATH,"//span[@class='_11JPr' and contains(text(), 'Looking outside your contacts...')]")
                    if loading:
                        try:
                            WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.XPATH, "//span[@class='_11JPr' and contains(text(), 'Looking outside your contacts...')]")))
                        except TimeoutException:
                            print("Loading element did not disappear within the timeout period")

                    in_contact = driver.find_elements(By.XPATH, "//div[@class='_2a-B5 VfC3c' and contains(text(), 'Contacts on WhatsApp')]")
                    not_in_contact = driver.find_elements(By.XPATH, "//div[@class='_2a-B5 VfC3c' and contains(text(), 'Not in your contacts')]")
                    not_in_whatsapp = driver.find_elements(By.XPATH, "//span[@class='_11JPr' and contains(text(), 'No results found')]")
                    time.sleep(0.7)
                    if in_contact:
                        button = driver.find_element(By.XPATH, "//div[@tabindex='-1' and @role='button']")
                        button.click()
                    elif not_in_contact:
                        button = driver.find_element(By.XPATH, "//div[@class='g0rxnol2 g0rxnol2 thghmljt p357zi0d rjo8vgbg ggj6brxn f8m0rgwh gfz4du6o ag5g9lrv bs7a17vp ov67bkzj']/div[@class='_2a-B5 VfC3c' and contains(text(), 'Not in your contacts')]/following-sibling::div[@tabindex='0' and @role='button']")
                        button.click()
                    elif not_in_whatsapp:
                        cancel_button = driver.find_element(By.CSS_SELECTOR, "button[aria-label='Cancel search']")
                        cancel_button.click()
                        continue
                    time.sleep(0.5)
                    message_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@class='to2l77zo gfz4du6o ag5g9lrv bze30y65 kao4egtt' and @contenteditable='true' and @role='textbox' and @title='Type a message']")))
                    for line in message.split('\n'):
                        message_input.send_keys(line + Keys.SHIFT + Keys.ENTER)
                    message_input.send_keys(Keys.RETURN)
                except Exception as e:
                    print(f"Error sending message to {num}: {e}")
                    continue
            if i + batch_size < len(numbers):
                print(f"Sent messages to {i + batch_size} contacts. Taking a break for 1 minute...")
                time.sleep(pause_duration)
    except Exception as e:
        print("An error occurred:", e)
    finally:
        if driver:
            time.sleep(2)
            driver.quit()
            messagebox.showinfo("Success", "Messages sent successfully")
root = tk.Tk()
root.title("WhatsApp Message Sender")
style = ttk.Style()
style.configure('TButton', foreground='blue')
numbers_label = ttk.Label(root, text="Numbers:")
numbers_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)
numbers_entry = tk.Text(root, width=40, height=4)
numbers_entry.grid(row=0, column=1, padx=10, pady=5)
message_label = ttk.Label(root, text="Message:")
message_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
message_entry = tk.Text(root, width=40, height=10)
message_entry.grid(row=1, column=1, padx=10, pady=5)
send_button = ttk.Button(root, text="Send Messages", command=send_messages)
send_button.grid(row=2, columnspan=2, pady=10)
root.mainloop()