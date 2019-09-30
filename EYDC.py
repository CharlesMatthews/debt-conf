#!/usr/bin/python3

import os
import time
import pandas as pd
import asyncio
import random
from threading import Thread

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
    df_Debtors = pd.read_excel(os.getcwd()+'/DataSource/DebtorsList.xlsx', sheet_name=0)
    df_Contact = pd.read_excel(os.getcwd()+'/DataSource/DebtorsList.xlsx', sheet_name=1)

    if samplesize > df_Debtors["Debtor ID"].nunique():
        print("PROBLEM?? debtors than available in sample????")

    Count = 1

    while Count < samplesize:
        ROW = random.randint(1,df_Debtors.shape[0]-1)

        #GETS LIST OF ALL ITEMS FROM SINGLE DEBTOR

        debtor_fl = df_Debtors.loc[df_Debtors['Debtor ID'] == df_Debtors["Debtor ID"][ROW]]
        print(debtor_fl)
        #GETS DEBTORS CONTACT
        debtor_con= df_Contact.loc[df_Contact['Debtor ID'] == df_Debtors["Debtor ID"][ROW]]
        print(debtor_con)

        invoices = []

        for ix, item in debtor_fl.iterrows():
            invo_item ={
                        "number": item["Document ID"],
                        "value": item["Value"],
                        "currency": item["Currency"]

                        }

            invoices.append(invo_item)
        debtor_con = debtor_con.values
        address_list = []
        address_list.extend([debtor_con[0][1],debtor_con[0][2],debtor_con[0][3],debtor_con[0][4],debtor_con[0][5],debtor_con[0][6]])

        outfile= str(Count)

        debtor_email = debtor_con[0][8]


        make_pdf(auditor_name,email,address_list,invoices, outfile , letter_head, signature_file)

        makemail(email,yearend,client_name,debtor_email, Count)

        Count +=1

    time.sleep(3)
    for i, client in enumerate(clients):
        print("Contacting Client",i,":   ", client)
        send_message(client, "http://willmatthews.xyz")
    print(clients)
    print("Msg Sent.")



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
        Save_File_Path = os.path.join(current_app.root_path, "PDFGen","lhead.pdf")
        if signature_file is not None:
            signature_file.save(Save_File_Path)
        Save_File_Path = os.path.join(current_app.root_path, "PDFGen","signature.png")
        if letter_head is not None:
            letter_head.save(Save_File_Path)

        ProcessData.submit(auditor_name,email,yearend,client_name,client_signatory,samplesize,datafile,letter_head,signature_file)
        #ProcessData(auditor_name,email,yearend,client_name,client_signatory,samplesize,datafile,letter_head,signature_file)

        flash(f"Data submitted and being processed. - We'll let you know when it's ready!", 'success')
        return redirect(url_for('index'))


    return render_template("dataentry.html", form = form)

@app.route('/')
def index():
    return render_template("index.html")

if __name__ == '__main__':
    socketio.run(app, debug = True)
