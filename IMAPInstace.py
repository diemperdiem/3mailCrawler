import datetime
import imaplib
import os
import email
import requests
from tkinter import *
from pathlib import Path

# Base
root = Tk()
root.title("Bot Email Crawler")
root.resizable(FALSE, FALSE)

textoLabel = StringVar()
textoLabel.set("Trabajando")

labelStatus = Label(root, textvariable=textoLabel, width=30, borderwidth=10)
labelStatus.pack()

home = str(Path.home())

# asigns the path for folders
pathPrincipal = "".join(home + "\\Desktop\\pdfPrueba\\")
pathPrincipal = Path(pathPrincipal)

# creates de folders if no existent
def createDirs():
    if not os.path.exists(pathPrincipal):
        os.makedirs(pathPrincipal)


def fixNameFile(fn):
    validchars = "-_.() "
    out = ""
    for c in fn:
        if str.isalpha(c) or str.isdigit(c) or (c in validchars):
            out += c
        else:
            out += "_"
    return out


def deletePDF():
    try:
        x = os.listdir(pathPrincipal)
        for file in x:
            if bool(file.__contains__('.pdf')):
                filePath = os.path.join(pathPrincipal, file)
                # os.close(filePath)
                os.remove(filePath)
    except Exception as e:
        textoLabel.set(e)


def emailCrawl():
    try:
        createDirs()
        deletePDF()
        # User
        username = 'XXXXXX'
        # Pass
        password = 'XXXXXX'
        mail = imaplib.IMAP4_SSL('outlook.office365.com', 993)

        mail.login(username, password)

        result, mailboxes = mail.list()
        mail.select("inbox")

        # query for email subject is a contian in string not a exact match, query can accept variables
        result, data = mail.search(None, '(FROM "' + username + '" SUBJECT "T")')

        mail_ids = data[0]
        mails_list = mail_ids.split()
        typ, data = mail.fetch(mails_list[-1], '(RFC822)')
        raw_email = data[0][1]

        # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        # downloading attachments
        for part in email_message.walk():
            # this part comes from the snipped I don't understand yet...
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            if bool(fileName.__contains__('.pdf')):
                # fileName = (fileName + str(datetime.datetime.now()))
                fileName = fixNameFile(fileName)
                if bool(fileName):
                    filePath = os.path.join(pathPrincipal, fileName)
                    if not os.path.isfile(filePath):
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                        url = 'http://localhost:51216/api/File/Upload'
                        files = {'file': open(filePath, 'rb')}
                        r = requests.post(url, files=files)
                        print(r.status_code)
                        print(r.text)
                        textoLabel.set("El documento se a manejado con exito")
    except Exception as e:
        print(e)
        textoLabel.set("Error intente de nuevo")


root.after(500)
emailCrawl()
# deletePDF()
labelStatus.update()
root.after(900)
root.destroy()
root.mainloop()

# labelStatus.grid(row=0, column=0)


#
#
# correoCuenta = Entry(root, width=40, borderwidth=3)
# passwordCorreo = Entry(root, width=40, borderwidth=3)
# correoObjetivo = Entry(root, width=40, borderwidth=3)

# myLabel1.grid(row=0, column=0)
# correoCuenta.grid(row=0, column=0, columnspan=3, padx=10, pady=10)
# correoCuenta.grid(row=0, column=0, padx=10, pady=10, columnspan=2)
# passwordCorreo.grid(row=1, column=0, padx=10, pady=10, columnspan=2)
# correoObjetivo.grid(row=2, column=0, padx=10, pady=10, columnspan=2)

# correoCuenta.grid(row=0, column=0, padx=10, pady=10)
# passwordCorreo.grid(row=1, column=0, padx=10, pady=10)
# correoObjetivo.grid(row=2, column=0, padx=10, pady=10)
#
# correoCuenta.insert(0, "Correo")
# passwordCorreo.insert(0, "Contraseña")
# correoObjetivo.insert(0, "Correo de cuenta a recibir ")

# correoCuenta.pack()
# passwordCorreo.pack()
# correoObjetivo.pack()

# myButton = Button(root, border=2, borderwidth=3, text="Start", command= lambda: ejecutar())
# myButton.grid(row=4, pady=(0, 10))
#
