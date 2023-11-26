# SHOULD install pip from the terminal in the python FOLDER.
# pip install pywin32
import os
import win32com.client
import random
import time
from datetime import datetime


def sendmail(users):
    # Create a new sendingList.txt file and remove the old one if it exists
    if os.path.exists('sendingList.txt'):
        os.remove('sendingList.txt')

    with open('sendingList.txt', 'a') as file:
        # Write the current date and time to the file
        file.write(f"official list, time: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}\n")

        for user in users:
            try:
                outlook = win32com.client.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = user['email']
                mail.Subject = 'Mini surprise for Hanukkah! 🎁🕎✨'
                mail.HTMLBody = f"<p><b style='color: blue;'>Hi dwarf {user['engName']}! </b></p>" \
                                f"<p>Guess what? Your Secret Santa is on a mission to spread some Hanukkah magic! " \
                                f"🕎 Get ready for a brighter gift than a dreidel. 🎁</p>" \
                                f"<p><b style='color: green;'>Your giant name is: {user['engGiant']} </b> 🌟🎅🎉</p>" \
                                f"<p><b style='color: red;'>Please note, maximum budget: 50 NIS </b> 💰</p> <br>" \
                                f"<p>This message is an<b style='color: red;'> official message</b> " \
                                f"for the dwarf and giant system.</p>" \
                                f"<p>P.S. Please send to the \"Perfect farts💨😍\" family group a confirmation message " \
                                f"that you have received this message and you've got your giant name. 📬</p>" \
                                f"<p>Thanks for the collaboration - Secret Santa Team 🎅🤶</p>" \
                                f"<br><br>" \
                                f"<br>" \
                                f"<p><b style='color: blue;'> שלום הגמד/ה {user['hebName']}!</b></p>" \
                                f"<p>נחש מה? סנטה הסודי שלך במשימה להפיץ קצת קסם של חנוכה! 🕎" \
                                f" התכוננו למתנה מדליקה יותר מסביבון. 🎁</p>" \
                                f"<p> 🌟🎅🎉 <b style='color: green;'>שם הענק שלך: {user['hebGiant']}</b> </p>" \
                                f"<p> 💰 <b style='color: red;'>שימו לב, תקציב מקסימלי: 50 ש\"ח</b></p> <br>" \
                                f"<p>הודעה זו היא<b style='color: red;'> הודעה רשמית</b> למערכת הגמד והענק.</p>" \
                                f"<p>אנא שלח/י לקבוצת המשפחה \"פלצנות מושלמת💨😍\" הודעת אישור שאת/ה " \
                                f"קיבלת הודעה זאת וקיבלת את שם הענק שלך 📬</p>" \
                                f"<p>🎅🤶 תודה על שיתוף הפעולה - צוות הגמד והענק</p>"
                mail.Send()
                print(f"Email sent to {user['engName']} giant is: {user['engGiant']}")
                time.sleep(1)

                # Append the details to the file
                file.write(f"Email sent to {user['engName']} giant is: {user['engGiant']} \n")

            except Exception as e:
                print(f"Error sending email to {user['engName']} giant is: {user['engGiant']}: {e}")
                time.sleep(1)


def generateSecretSanta(users):
    # Shuffle the list to randomize the assignment
    random.shuffle(users)

    for i in range(len(users)):
        person = users[i]

        giant_index = (i + 1) % len(users)
        giant = users[giant_index]

        # Update the 'engGiant' and 'hebGiant' keys
        person['engGiant'] = giant['engName']
        person['hebGiant'] = giant['hebName']

    return users


users = [
    {'engName': "Tal", 'hebName': "טל", 'email': 'talfreestyle@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Adina", 'hebName': "עדינה", 'email': 'Adina03@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Batya", 'hebName': "בתיה", 'email': 'Lbatya123@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Devora", 'hebName': "דבורה", 'email': 'devch248@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Dani", 'hebName': "דני", 'email': 'lubindaniel56@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Malka", 'hebName': "מלכה", 'email': '4lymalka@gmail.com', 'hebGiant': "", 'engGiant': ""},
    {'engName': "Tzvi", 'hebName': "צבי", 'email': 'tsaalenu@gmail.com', 'hebGiant': "", 'engGiant': ""},
]

users = generateSecretSanta(users)

sendmail(users)
print("mails are successfully sent!")
