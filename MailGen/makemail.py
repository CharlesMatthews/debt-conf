#!/usr/bin/python3

import win32com.client as win32

import pythoncom

from datetime import datetime

import os


def makemail(email,yearend,client_name,debtor_email,Count):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = debtor_email
    mail.Subject = "Ernst & Young LLP Audit Request - Confirmation of Debt"
    mail.HtmlBody = ('''Good Afternoon,<br>
    I am writing on behalf of Ernst and Young LLP with regards to the FY19 audit of '''+ client_name +'''<br><br>

    As part of our audit procedures we are required to confirm the debtors balance which is held between yourselves and '''+client_name+''', as at '''+ str(yearend.strftime("%d %B, %Y")) +'''.<br><br>

    The debtors confirmation letter has been attached, this details an amount '''+client_name+ ''' believe is owed at the confirmation date. Could you please read the letter attached and complete the relevant sections as required.<br><br>

    Please sign and date the letter and provide a scanned copy of this directly to the following email addresses, '''+ email +'''<br><br>

    Further instructions are included in the attached. A prompt confirmation is greatly appreciated.<br><br>

    Thank you for your cooperation.<br><br>

    If you have any questions please let me know
    ''')


    mail.Attachments.Add(os.getcwd()+"\Output\PDFs\\" + str(Count) + ".pdf")
    mail.SaveAs(Path=os.getcwd()+"\Output\Emails\\" + str(Count) + ".msg")


    return
