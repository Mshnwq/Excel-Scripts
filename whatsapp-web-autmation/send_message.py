import pywhatkit as kit
import pandas as pd
import pygetwindow as gw
import pyautogui
import time

message = "السلام عليكم ورحمة الله وبركاته\
\n\
\n\
*دعوة خاصة * \
\n\
\n\
 تتقدم الهيئة السعودية للمهندسين ممثلة بفرع المنطقة الغربية بدعوتكم \
للمشاركة في ورشة عمل تعزيز دور المكاتب والشركات الهندسية في المشارع \
الكبرى وذلك يوم الخميس الموافق 7 - 3 - 2024م بمقر فرع الهيئة بجدة \
تبدأ الورشة في تمام الساعة العاشرة صباحاً \
\n\
نامل تاكيد حضورك بالرد على هذه الرسالة \
\n\
وتقبلو فائق التحية والتقدير \
\n \
م. محمد عطار \
\n\
+966593334586 \
\n\
الهيئة السعودية للمهندسين \
\n\
فرع المنطقة الغربية"
# Read contacts from Excel sheet
df = pd.read_excel("data/test.xlsx", usecols=['Phone'])

# Iterate through each row in the DataFrame
for contact in df['Phone']:
    # Send message to contact
    try:
        kit.sendwhatmsg_instantly(f'+{str(int(contact))}', message)
        # Wait for a brief moment to ensure message sending is complete
        time.sleep(25)
        # Check if browser window is in focus
        browser_windows = gw.getWindowsWithTitle("Chrome")  # Replace "Chrome" with your browser title
        if browser_windows:
            browser_window = browser_windows[0]
            if not browser_window.isActive:
                browser_window.activate()
                time.sleep(2)  # Add a brief delay to ensure activation
            print("Focusing Browser Window")
        else:
            print("Browser window not found.")
        # Find the coordinates of the element you want to click on
        x_coordinate = 1200  # Replace with the x-coordinate of the element
        y_coordinate = 900  # Replace with the y-coordinate of the element
        # Move the mouse to the coordinates of the element
        pyautogui.moveTo(x_coordinate, y_coordinate)
        time.sleep(2)  # Add a brief delay to ensure activation
        # Perform a mouse click
        pyautogui.click()
        time.sleep(2)  # Add a brief delay to ensure activation
        # Press the Enter key using pyautogui
        pyautogui.press('enter')
        print("Pressing Enter")
        time.sleep(5)
        print(f"Message sent to {int(contact)} successfully!")
    except Exception as e:
        print(f"Failed to send message to {contact}: {str(e)}")