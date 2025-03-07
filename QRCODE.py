
import cv2
from pyotp import otp
from pyzbar.pyzbar import decode
import pyotp
from datetime import datetime

# Charger l'image contenant le code QR
image = cv2.imread(r'C:\Users\ayman\OneDrive\Bureau\Virtuo Code\Etail\Script_client_SectretKey\AuthAmazonQRCode.png')

decoded_objects = decode(image)
print(decoded_objects)

if decoded_objects:
    qr_data = decoded_objects[0].data.decode('utf-8')
    print(qr_data)

    # Afficher le contenu du code QRpip install opencv-python
    print("OTP code was generated successfully...")

else:
    print("no qr")

otpauth_url = "otpauth://totp/Amazon%3Adev%40etail-agency.com?secret=N6ROBUP7P2YHKZQOVKJDB3YLCSVCUMWOHSR2PZOQPLOAMFB37ASQ&issuer=Amazon"
print(otpauth_url)


def generate(code):
    otp_data = pyotp.parse_uri(code)
    totp = pyotp.TOTP(otp_data.secret)
    otp_code = totp.now()
    return(otp_code )


# storing the current time in the variable
c = datetime.now()

# Displays Time
current_time = c.strftime('%H:%M:%S')
print('Current Time is:', current_time)
# OR
# Displays Date along with Time
print('Current Date and Time is:', c)



