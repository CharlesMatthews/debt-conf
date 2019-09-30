#!/usr/bin/python3

import os
import time
import pandas as pd
import asyncio
import random
from threading import Thread

import shutil


from flask import Flask, render_template, redirect, url_for, flash, current_app, request
from flask_socketio import SocketIO

from flask_executor import Executor
from templates.forms import Form_Generate
from PDFGen.makepdf import make_pdf
from MailGen.makemail import makemail


executor = Executor()
app = Flask(__name__)
app.config['SECRET_KEY'] = 'Sup>.er"3uhedoa()U*yuoehdoSeceteyQacijas92*()!'

executor.init_app(app)
socketio = SocketIO(app)


clients = []



@executor.job
def ProcessData(auditor_name,email,yearend,client_name,client_signatory,samplesize,datafile,letter_head,signature_file):
    if os.path.exists(os.path.join(current_app.root_path,"static","out.zip")):
        os.remove(os.path.join(current_app.root_path,"static","out.zip"))
    df_Debtors = pd.read_excel(os.path.join(current_app.root_path,"DataSource","DebtorsList.xlsx"), sheet_name=0)
    df_Contact = pd.read_excel(os.path.join(current_app.root_path,"DataSource","DebtorsList.xlsx"), sheet_name=1)

    if samplesize > df_Debtors["DebtorID"].nunique():
        print("User Entry issue-- more samples request than is available????")
        samplesize = df_Debtors["DebtorID"].nunique()

    Count = 1
    file_list = []

    while Count < samplesize:
        time.sleep(0.02)

        ROW = random.randint(1,df_Debtors.shape[0]-1)

        #GETS LIST OF ALL ITEMS FROM SINGLE DEBTOR

        debtor_fl = df_Debtors.loc[df_Debtors['DebtorID'] == df_Debtors["DebtorID"][ROW]]
        print(debtor_fl)
        #GETS DEBTORS CONTACT
        debtor_con= df_Contact.loc[df_Contact['DebtorID'] == df_Debtors["DebtorID"][ROW]]
        print(debtor_con)

        invoices = []

        for ix, item in debtor_fl.iterrows():
            invo_item ={
                        "number": item["DocumentID"],
                        "value": item["Value"],
                        "currency": item["Currency"]

                        }

            invoices.append(invo_item)
        debtor_con = debtor_con.values
        address_list = []
        address_list.extend([debtor_con[0][1],debtor_con[0][2],debtor_con[0][3],debtor_con[0][4],debtor_con[0][5],debtor_con[0][6]])

        outfile= str(Count)

        debtor_email = debtor_con[0][8]


        make_pdf(auditor_name,email,address_list,invoices, outfile , letter_head, signature_file, client_signatory)

        makemail(email,yearend,client_name,debtor_email, Count)

        #file_list.append(os.getcwd()+"\Output\PDFs\\" + str(Count) + ".pdf")
        #file_list.append(os.getcwd()+"\Output\Emails\\" + str(Count) + ".msg")

        # Get names of indexes for which column Age has value 30

        # Delete these row indexes from dataFrame

        print("\n")
        print("\n")
        print(df_Debtors)
        print("\n")
        print("\n")
        print("\n")
        # Get names of indexes for which column Age has value 30
        InNL=[]
        indexNames = df_Debtors[ df_Debtors['DebtorID'] == df_Debtors["DebtorID"][ROW] ].index
        for item in indexNames:
            InNL.append(item)

        df_Debtors.drop(InNL, inplace=True)
        df_Debtors.reset_index(drop=True, inplace=True)
        Count +=1
    time.sleep(2)
    shutil.make_archive(os.getcwd()+"/static/out", 'zip', os.getcwd()+"/Output/Emails")
    time.sleep(2)

    for i, client in enumerate(clients):
        print("Contacting Client",i,":   ", client)
        send_message(client, "/static/out.zip")
    print(clients)
    print("Msg Sent.")
    os.remove(os.path.join(current_app.root_path,"Output","PDFs","signature.png"))
    os.remove(os.path.join(current_app.root_path,"Output","PDFs","lhead.pdf"))
    print("CleanedUp")



@socketio.on('connect')
def handle_connect():
    print('Client connected')
    clients.append(request.sid)


def send_message(client_id, data):
    socketio.emit('output', data, room=client_id)
    print('sending message "{}" to client "{}".'.format(data, client_id))


@socketio.on('disconnect')
def handle_disconnect():
    print('Client disconnected')
    clients.remove(request.sid)




@app.route('/generator')
def generator():


    return render_template("generator.html")


@app.route('/dataentry', methods = ["GET", "POST"])
def dataentry():
    form = Form_Generate()
    #if form.validate_on_submit():

    if form.is_submitted():

        auditor_name = form.name.data
        email = form.email.data
        yearend = form.yearend.data
        client_name = form.client_name.data
        client_signatory = form.client_signatory.data
        samplesize = form.samplesize.data
        datafile = form.datafile.data
        letter_head = form.letter_head.data
        signature_file = form.signature.data


        Save_File_Path = os.path.join(current_app.root_path, "DataSource","DebtorsList.xlsx")
        datafile.save(Save_File_Path)
        Save_File_Path = os.path.join(current_app.root_path,"Output","PDFs","signature.png")

        if signature_file is not None:
            signature_file.save(Save_File_Path)

            signature_file = "signature.png"
        Save_File_Path = os.path.join(current_app.root_path, "Output","PDFs","lhead.pdf")
        if letter_head is not None:
            letter_head.save(Save_File_Path)
            letter_head = "lhead.pdf"

        ProcessData.submit(auditor_name,email,yearend,client_name,client_signatory,samplesize,datafile,letter_head,signature_file)
        #ProcessData(auditor_name,email,yearend,client_name,client_signatory,samplesize,datafile,letter_head,signature_file)

        flash(f"Submission received and is being processed. We'll let you know when it's ready! 😎", 'success')
        return redirect(url_for('index'))


    return render_template("dataentry.html", form = form)

@app.route('/')
def index():
    return render_template("index.html")

if __name__ == '__main__':
    socketio.run(app, debug = True)
