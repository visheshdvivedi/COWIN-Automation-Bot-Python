import xlsxwriter, traceback
import requests
import os, random, time, re, hashlib, json, datetime, platform, getpass, smtplib
import art
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

user_agents = []
email, password = '', ''

PHONENUMBER_REGEX = "^\s*(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})(?: *x(\d+))?\s*$"
SECRET = "U2FsdGVkX1/PCU54Y6nGWMC1Qu7EyRJmbpx0pKs05Pdpf26szah6SEsAjxMGZAb39PznmhSjJlNYwvEc00gQow=="
POST_HEADERS = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}
GET_HEADERS = {"Accept": "application/json"}
USERAGENT = ''
STATES_FILENAME = "states.xlsx"
PLATFORM = platform.system()

GOOGLE = "https://www.google.com/"
BASE_URL = "https://cdn-api.co-vin.in/api"
HOST = "cdn-api.co-vin.in"
GET_OTP_URL = BASE_URL + "/v2/auth/public/generateOTP"
CONFIRM_OTP_URL = BASE_URL + "/v2/auth/public/confirmOTP"
GET_LIST_OF_STATES = BASE_URL + "/v2/admin/location/states"
GET_LIST_OF_DISTRICTS = BASE_URL + "/v2/admin/location/districts/"
FIND_APPOINTMENT_BY_PINCODE = BASE_URL + "/v2/appointment/sessions/public/findByPin"
FIND_APPOINTMENT_BY_DISTRICT = BASE_URL + "/v2/appointment/sessions/public/findByDistrict"
FIND_APPOINTMENT_BY_LATLONG = BASE_URL + "/v2/appointment/centers/public/findByLatLong"

def printResponseData(response):
    '''
    Used for debugging purpose to display HTTP response related data
    '''
    print("\nData For Request")
    print("URL: ", response.request.url)
    print("Header: ", response.request.headers)
    print("Body: ", response.request.body)
    print("\nData For Response")
    print("URL: ", response.url)
    print("Code: ", response.status_code)
    print("Header: ", response.headers)
    print("Body: ", response.text)

def get_user_agents():
    '''
    To read the user agents from user_agents.txt
    '''
    global user_agents
    with open("user_agents.txt", 'r') as file:
        user_agents = file.readlines()

def getPhoneNumberFromUser():
    '''
    Function to get phone number from user.
    This function uses regex to check if phone number format is correct or not
    '''
    while 1:
        number = input("Please enter your mobile number: ")
        if re.match(PHONENUMBER_REGEX, number) != None:
            return number
        else:
            print("[ERROR] Invalid number. Please try again\n")

def welcome_message():
    '''
    Function to print a beautiful welcome message for our script
    It uses art module to print an ASCII art
    And it also displays detals about my different accounts
    on social media.
    '''
    global user_agents
    if PLATFORM == "Windows":
        os.system("cls")
    else:
        os.system("clear")
    print(art.text2art("PYTHON + COWIN"))
    print("\nCreated By: Vishesh Dvivedi [All About Python]\n")
    print("GitHub:    visheshdvivedi")
    print("Youtube:   All About Python")
    print("Instagram: @itsallaboutpython")
    print("Blog Site: itsallaboutpython.blogspot.com\n")

def checkInternetConnectivity():
    '''
    Check for internet connectivity by contacting google site.
    If response status code is 200, we are connected to internet and
    the function return True
    else it returns False
    '''
    print("\n[INFO] Checking for Internet Connectivity...")
    if not requests.get(GOOGLE).status_code == 200:
        print("[ERROR] Internet Connectivity : False")
        print("[ERROR] Exiting program...")
        exit(0)
    else:
        print("[INFO] Internet Connectivity : True\n")

def getOTPFromAPI(number):
    '''
    Function to GET OTP from COWIN API using mobile number
    '''
    while 1:
        USERAGENT = random.choice(user_agents).replace("\n", "")
        POST_HEADERS["User-Agent"] = USERAGENT
        response = requests.post(GET_OTP_URL, json={"secret": SECRET, "mobile":str(number)}, headers=POST_HEADERS)
        if response.status_code == 200:
            print("[INFO] API Has Send OTP to {0}".format(number))
            txnId = response.json()["txnId"]
            otp = input("Enter the OTP you received (If you don't receive it in 2 minutes, click on Enter key): ")
            if otp == "":
                continue
            else:
                return otp, txnId
        elif response.status_code == 403 or response.status_code == 400:
            print("[ERROR] Server is busy....sending another request after 10 seconds")
            time.sleep(10)
        else:
            print("[ERROR] Something went wrong while retrieving OTP from COWIN API")
            print("COde:", response.status_code)
            print("Text:", response.text)
            print("Sending another request after 10 seconds..")
            time.sleep(10)


def confirmOTP(otp, txnId):
    '''
    Function to GET OTP from COWIN API using txnID retrieved using getOTP function
    '''
    while 1:
        POST_HEADERS["User-Agent"] = USERAGENT
        code = hashlib.sha256(otp.encode()).hexdigest()
        response = requests.post(CONFIRM_OTP_URL, json={"otp":code, "txnId":txnId}, headers=POST_HEADERS)
        if response.status_code == 200:
            token = response.json()["token"]
            return token
        elif response.status_code == 403:
            print("[ERROR] Server is busy....sending another request after 10 seconds")
            time.sleep(10)
        else:
            print("[ERROR] Something went wrong while confirming OTP from COWIN API")
            print("COde:", response.status_code)
            print("Text:", response.text)

def getListOfStates(token):
    '''
    To get a list of states from the COWIN API and save it in states.xlsx
    '''
    GET_HEADERS["Host"] = HOST
    GET_HEADERS["User-Agent"] = random.choice(user_agents).replace("\n", "")
    GET_HEADERS["Authorization"] = "Bearer {0}".format(token)
    while 1:
        response = requests.get(GET_LIST_OF_STATES, headers=GET_HEADERS)
        if response.status_code == 200:
            data = response.json()
            saveListOfStates("states.xlsx", data)
            break
        else:
            printResponseData(response)
            break

def getListOfDistricts(state_id, token):
    '''
    To get a list of districts based on the state ID and save if in districts.xlsx
    '''
    GET_HEADERS["Host"] = HOST
    GET_HEADERS["User-Agent"] = random.choice(user_agents).replace("\n", "")
    GET_HEADERS["Authorization"] = "Bearer {0}".format(token)
    while 1:
        response = requests.get(GET_LIST_OF_DISTRICTS + state_id, headers=GET_HEADERS)
        if response.status_code == 200:
            data = response.json()
            saveListOfDistricts("districts.xlsx", data)
            break
        else:
            printResponseData(response)
            break

def saveListOfStates(filename, data):
    '''
    To save the list of states received from the API onto an excel by the name 'filename'
    '''
    if not filename in os.listdir():
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "State ID")
        worksheet.write(0, 1, "State Name")
        row = 1
        for state in data["states"]:
            worksheet.write(row, 0, state["state_id"])
            worksheet.write(row, 1, state["state_name"])
            row += 1
        workbook.close()

def saveListOfDistricts(filename, data):
    '''
    To save list of districts received from the API onto an excel by the name 'filename'
    '''
    if not filename in os.listdir():
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "District ID")
        worksheet.write(0, 1, "District Name")
        row = 1
        for state in data["districts"]:
            worksheet.write(row, 0, state["district_id"])
            worksheet.write(row, 1, state["district_name"])
            row += 1
        workbook.close()

def getUserChoiceMainMenu():
    '''
    Function to ask for user commands
    '''
    if PLATFORM == "Windows":
        os.system("cls")
    else:
        os.system("clear")
    welcome_message()
    while 1:
        try:
            print("="*50)
            print("What do you want to do: ")
            print("1. Get list of districts.")
            print("2. Get list of vaccination slots by pincode.")
            print("3. Get list of vaccination slots by district.")
            print("4. Get list of vaccination slots by latitude and longitude.")
            print("5. Set up an email reminder loop")
            print("6. Exit")
            choice = int(input("Enter your choice (1/2/3/4): "))
            if choice in [1, 2, 3, 4, 5, 6]:
                if choice == 1:
                    print("[INFO] Choice: Get list of districts.")
                    return "DISTRICTS"
                elif choice == 2:
                    print("[INFO] Choice: Get list of vaccination slots by pincode.")
                    return "SLOTS-PINCODE"
                elif choice == 3:
                    print("[INFO] Choice: Get list of vaccination slots by district.")
                    return "SLOTS-DISTRICT"
                elif choice == 4:
                    print("[INFO] Choice: Get list of vaccination slots by latitude and longitude.")
                    return "SLOTS-LATLONG"
                elif choice == 5:
                    print("[INFO] Choice: Set up an email reminder loop.")
                    return "EMAIL"
                elif choice == 6:
                    print("[INFO] Exiting...")
                    exit(0)
                else:
                    print("[ERROR] Please enter number between 1 and 4")
        except:
            print("[ERROR] Invalid input..enter a number between 1 and 4")
        

def processUserChoice(opt, token, isEmailLoop = False):
    '''
    To process user commands and perform correct actions based on the input
    '''
    output = {
        "by": "",
        "param":{}
    }
    curr_time = datetime.datetime.now().strftime("%d-%m-%Y")
    if opt == "DISTRICTS":
        print("[INFO] Select your state from the states.xlsx excel and enter its state ID below")
        id = input("Enter your state ID: ")
        getListOfDistricts(id, token)
    elif opt == "SLOTS-PINCODE":
        pincode = int(input("[INFO] Enter pincode: "))
        output["by"] = "pincode"
        output["param"] = {"pincode": pincode,"date": curr_time}
        data = getDataFromAPI(token, output)
        if len(data["sessions"]) == 0:
            print("[INFO] No data found for this request...")
            return 0
        if not isEmailLoop:
            saveDataInExcel(data)
        else:
            return data
    elif opt == "SLOTS-DISTRICT":
        district_id = int(input("Enter district id: "))
        output["by"] = "district"
        output["param"] = {"district_id": district_id,"date": curr_time}
        data = getDataFromAPI(token, output)
        if len(data["sessions"]) == 0:
            print("[INFO] No data found for this request...")
            return 0
        if not isEmailLoop:
            saveDataInExcel(data)
        else:
            return data
    elif opt == "SLOTS-LATLONG":
        lat = int(input("Enter latitude: "))
        long = int(input("Enter longitude: "))
        output["by"] = "latlong"
        output["param"] = {"lat": lat,"long": long}
        data = getDataFromAPI(token, output)
        if len(data["sessions"]) == 0:
            print("[INFO] No data found for this request...")
            return 0
        if not isEmailLoop:
            saveDataInExcel(data)
        else:
            return data
    elif opt == "EMAIL":
        option = setEmailReminderLoop()
        startEmailReminderLoop(option, token)
        data = getDataFromAPI(token, output)
        if len(data["sessions"]) == 0:
            print("[INFO] No data found for this request...")
            return 0
        if not isEmailLoop:
            saveDataInExcel(data)

def startEmailReminderLoop(opt, token):
    '''
    To ask user for email reminder settings
    '''
    output = {
        "by": "",
        "param": {}
    }
    curr_time = datetime.datetime.now().strftime("%d-%m-%Y")
    try:
        if opt == "SLOTS-PINCODE":
            pincode = int(input("[INFO] Enter pincode: "))
            output["by"] = "pincode"
            output["param"] = {"pincode": pincode,"date": curr_time}
        elif opt == "SLOTS-DISTRICT":
            district_id = int(input("Enter district id: "))
            output["by"] = "district"
            output["param"] = {"district_id": district_id,"date": curr_time}
        elif opt == "SLOTS-LATLONG":
            lat = int(input("Enter latitude: "))
            long = int(input("Enter longitude: "))
            output["by"] = "latlong"
            output["param"] = {"lat": lat,"long": long}
        print("[INFO] Initiating email loop")    
        while 1:
            count = 0
            slots = {"sessions":[]}
            data = getDataFromAPI(token, output)
            if len(data["sessions"]) != 0:
                for session in data["sessions"]:
                    if int(session["available_capacity"]) > 0:
                        count += 1
                        slots["sessions"].append(session)
                print("[INFO] Found {0} sessions. ".format(count))
                filename = saveDataInExcel(slots, True)
                sendMail(os.path.join(os.getcwd(), 'data', filename))
            else:
                print("[INFO] No slots available. ")
                continue

            print("[INFO] Waiting for 30 minutes before sending another request...")
            time.sleep(1800)

    except KeyboardInterrupt:
        print("[INFO] User interrupt....exiting")
        exit(0)
    except Exception as ex:
        print(ex)

def setEmailReminderLoop():
    '''
    To set and start the email reminder loop based on user settings
    '''
    if PLATFORM == "Windows":
        os.system("cls")
    else:
        os.system("clear")
    while 1:
        try:
            print("="*50)
            print("How do you want to get the vacciation slots: ")
            print("1. Get vaccination slots by pincode.")
            print("2. Get vaccination slots by district.")
            print("3. Get vaccination slots by latitude and longitude.")
            print("4. GO back to main menu.")
            print("5. Exit")
            choice = int(input("Enter your choice (1/2/3/4/5): "))
            if choice in [1, 2, 3, 4, 5]:
                if choice == 1:
                    return "SLOTS-PINCODE"
                elif choice == 2:
                    return "SLOTS-DISTRICT"
                elif choice == 3:
                    return "SLOTS-LATLONG"
                elif choice == 4:
                    return "MAINMENU"
                elif choice == 5:
                    print("[INFO] Exiting...")
                    exit(0)
            else:
                print("[ERROR] Please enter one of these integers (1, 2, 3, 4, 5)")
                continue
        except:
            print("[ERROR] Invalid input...Please try again")
            break                
        
def sendMail(filepath):
    '''
    Function to send email reminder mail and attach the required excel at 'filepath'
    '''
    global email, password
    content = '''
    Hey there,
    Here are the available slots found based on your script settings.

    Regards,
    COWIN Automator

    GitHub: visheshdvivedi
    YouTube: All About Python
    Instagram: @itsallaboutpython
    Blog: itsallaboutpython.blogspot.com
    '''
    host =  ""
    port = ""
    while 1:  
        try:      
            if email == "" and password == "":
                print("[INFO] We need an email account to send you mails.")
                email = input("[INFO] Kindly enter a valid email ID: ")
                password = getpass.getpass("[INFO] Kindly enter the password: ")
                receiver_email = input("[INFO] Enter the email ID you want to send emails to: ")
            if "gmail" in email:
                host = "smtp.gmail.com"
                port = 587
            elif "outlook" in email:
                host = "smtp-mail.outlook.com"
                port = 587
            else:
                print("[ERROR] This script only supports gmail and outlook IDs as of now. Please enter a gmail or outlook ID")
                email, password = '', ''

            mail = smtplib.SMTP(host, port)
            mail.starttls()
            try:
                mail.login(email, password)
            except Exception as ex:
                print("[ERROR] Invalid username and password...please try again")
                print(traceback.print_exc())
                email, password = '', ''
                continue

            message = MIMEMultipart()
            message['From'] = email
            message['To'] = receiver_email
            message['Subject'] = "Available Slots: COWIN Automator"
            message.attach(MIMEText(content, 'plain'))
            file_data = open(filepath, 'rb').read()
            payload = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            payload.set_payload(file_data)
            encoders.encode_base64(payload)
            payload.add_header('Content-Decomposition', 'attachment', filename="Available_Slots")
            message.attach(payload)
            text = message.as_string()
            mail.sendmail(email, receiver_email, text)
            mail.quit()
            print("[INFO] Seccessfully send mail to {0}".format(receiver_email))
            break
        except Exception as ex:
            print(traceback.print_exc())
            break
        

def getDataFromAPI(token, choice):
    '''
    To get data from API based on user choices and commands
    '''
    by = choice["by"]
    url = ""
    POST_HEADERS["Authorization"] = "Bearer {0}".format(token)

    if by == "pincode":
        url = FIND_APPOINTMENT_BY_PINCODE
    elif by == "district":
        url = FIND_APPOINTMENT_BY_DISTRICT
    elif by == "latlong":
        url = FIND_APPOINTMENT_BY_LATLONG

    response = requests.get(url, params=choice['param'], headers=POST_HEADERS)
    if response.status_code == 200:
        return response.json()
    else:
        printResponseData(response)

def saveDataInExcel(data, isReturnFilename = False):
    '''
    To save data from API onto an excel file
    '''
    curr_dir = os.getcwd()
    if not os.path.exists(os.path.join(curr_dir, 'data')):
        os.mkdir("data")
    os.chdir(os.path.join(curr_dir, "data"))
    for file in os.listdir(os.path.join(curr_dir, "data")):
        if file.endswith(".xlsx"):
            os.unlink(file)
    filename = datetime.datetime.now().strftime("%d-%m-%Y %H-%M-%S") + ".xlsx"
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    count = 0
    for key in data["sessions"][0].keys():
        worksheet.write(0, count, str(key))
        count += 1
    row = 1
    for session in data["sessions"]:
        column = 0
        for key in session.keys():
            if key != "slots":
                worksheet.write(row, column, session[key])
            else:
                data = ", ".join(session[key])
                worksheet.write(row, column, data)
            column += 1
        row += 1
    workbook.close()
    os.chdir(curr_dir)
    print("[INFO] Saved data within data folder in in file {0}".format(filename))
    if isReturnFilename:
        return filename

def main():
    '''
    main function
    '''
    global user_agents
    welcome_message()
    get_user_agents()
    checkInternetConnectivity()
    phone_number = getPhoneNumberFromUser()
    otp, txnId = getOTPFromAPI(phone_number)
    token = confirmOTP(otp, txnId)
    print("[INFO] Token successfully received")
    
    if not STATES_FILENAME in os.listdir():
        print("[INFO] Retrieving states data.")
        getListOfStates(token)
    else:
        print("[INFO] States data already present ('states.xlsx')")

    while 1:
        opt = getUserChoiceMainMenu()
        processUserChoice(opt, token)
        input("Press Enter to go to main menu...")


if __name__ == "__main__":
    main()
