import time
import pyautogui
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC




# Initialize the undetected Chrome driver
options = uc.ChromeOptions()
options.add_argument("--remote-debugging-port=5555")
options.add_argument("user-data-dir=selenium") #You only need to login once
options.add_argument("--log-level=3") # I hate useless logs

driver = uc.Chrome(options=options)


#for automatically loggin into google
driver.get("https://accounts.google.com/signin/v2/identifier?service=mail") # Open Gmail login page

# Enter email
email_field = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="email"]'))
)
email_field.send_keys("your_emil_id") # Add your email id inside ""

# Click "Next" button
next_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '#identifierNext button'))
)
next_button.click()

# Enter password
password_field = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="password"]'))
)
password_field.send_keys("your_password_here") # Add your password inside ""

# Click "Next" button
next_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, '#passwordNext button'))
)
next_button.click()



time.sleep(50) # delay added for 2fa



# Open Gmail inbox page
driver.get("https://mail.google.com/mail/u/1/#inbox")

# Wait for the "Compose" button to be clickable to ensure the page has loaded
WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.T-I.T-I-KE.L3'))
)

# Load contacts from the Excel file
contacts_df = pd.read_excel('contacts.xlsx', sheet_name='Sheet1')

# Iterate over each email in the contacts
for index, row in contacts_df.iterrows():
    email = row['User Emails'] # To identify the Emails  
    first_name = row['First Name'] # To identify the first name
  # custom_message = row['Title in Excel']
    
    # Click the "Compose" button
    compose_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.T-I.T-I-KE.L3'))
    )
    compose_button.click()

    # Enter recipient email
    to_field = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[aria-label="To recipients"]'))
    )
    
    for char in email:
        to_field.send_keys(char)
        time.sleep(0.1)  # Simulate typing delay

    # Click the "Subject" text box and enter the subject
    subject_field = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="subjectbox"]'))
    )

  
    #This is where you can add your subject for the email
    subject = "Enter your subject here" # remove
    
    
    
    for char in subject:
        subject_field.send_keys(char)
        time.sleep(0.1)  # Simulate typing delay
    
    subject_field.send_keys("\n")  # Press Enter to move to the body

    # Enter the email body
    body_field = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div[aria-label="Message Body"]'))
    )
    
    
    # This is where you add your email content
    # {first_name} is there to make it more personalized
    # You can make the email more personalized by adding more data to the Excel file.
    # For example, you can add a column with the name "Custom Message".
    # Then, in the body of the email, you can include a placeholder like {custom_message}.
    # At line 70 in the script, replace the placeholder with the actual data from the Excel file.

    body = f"""Hello {first_name}, """
    
    
    
    for char in body:
        body_field.send_keys(char)
        time.sleep(0.1)  # Simulate typing delay  
      
    #print(first_name)


  
  
    # body_field.send_keys(Keys.CONTROL, 'v') # Uncheck this if you want to paste something to the email (eg: an image with a fixed size or something)

  
  
  
    # Attach the file
    attach_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.a1'))
    )
    attach_button.click()
    time.sleep(2)
    pyautogui.typewrite(r"Path to the file that you weant to attach") #btw you can attach multiple file if you copy paste this line and the line below 
    pyautogui.press("enter")
    time.sleep(10)
    
    # Scroll to the "Send" button and click using JavaScript
    send_button = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.T-I.J-J5-Ji.aoO.v7.T-I-atl.L3'))
    )
    driver.execute_script("arguments[0].scrollIntoView();", send_button)
    
    # Retry clicking the "Send" button if intercepted
    for _ in range(3):
        try:
            send_button.click()
            print("Email sent to ", email, " successfully")
            break
        except Exception as e:
            print(f"Retrying due to: {e}")
            time.sleep(4)
    
    # Wait for the confirmation message
    try:
        WebDriverWait(driver, 30).until(
            EC.text_to_be_present_in_element((By.CSS_SELECTOR, '.b8 div.vh'), "Message sent")
        )
        print("Message successfully sent.")
        # Update Excel file with "Message sent" and timestamp
        contacts_df.loc[index, 'Status'] = "Message sent"
        contacts_df.loc[index, 'Time Stamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        contacts_df['Time Stamp'] = pd.to_datetime(contacts_df['Time Stamp'])  # Ensure dtype is datetime
        contacts_df.to_excel('contacts.xlsx', index=False)
    except Exception as e:
        print(f"Failed to confirm message sent for {email}: {e}")
        contacts_df.loc[index, 'Status'] = "Failed to send"
        contacts_df.loc[index, 'Time Stamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        contacts_df['Time Stamp'] = pd.to_datetime(contacts_df['Time Stamp'])  # Ensure dtype is datetime
        contacts_df.to_excel('contacts.xlsx', index=False)
        time.sleep(30)

print("all emails sent successfully")
driver.quit()
