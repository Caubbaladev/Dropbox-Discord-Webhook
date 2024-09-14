#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jan 28 09:05:45 2024

@author: caubbaladev
"""

from flask import Flask, request, jsonify, Response
import dropbox
import pandas as pd
import docx2txt
from discord_webhook import DiscordWebhook, DiscordEmbed

FILEPATH = ''
APP_KEY = ''
APP_SECRET = ''
REFRESH_TOKEN = ''
UPD_TEXT = "Une **mise à jour** a été apportée au document '**Quoi de neuf.docx**' !\n_ _"
UI_TITLE = "Quoi de neuf ?"
SAMPLE_FILE_NAME = "SAMPLE_DOCX"
UPDATED_FILE_NAME = "UPDATED_DOCX"
DISCORD_WEBHOOK_URL = ''

webhook = DiscordWebhook(
    url=DISCORD_WEBHOOK_URL,
    content=UPD_TEXT
    )


def  textFormat(textList):  
    
    newText = ''
    for text in textList:
        if text == '':
            newText = newText + '\n\n'
        else:
            newText = newText + text
    return newText




def dropbox_getFiles(path):
    dbx = dropbox.Dropbox(
                app_key = APP_KEY,
                app_secret = APP_SECRET,
                oauth2_refresh_token = REFRESH_TOKEN
            )
    
    dbx.users_get_current_account()
    
    print("Successfully set up client!")

    
    
    
    try:
        files = dbx.files_list_folder(path).entries
        files_list = []
        for file in files:
            
            if isinstance(file, dropbox.files.FileMetadata):
                metadata = {
                    'name': file.name,
                    'path_display': file.path_display,
                    'client_modified': file.client_modified,
                    'server_modified': file.server_modified
                }
                files_list.append(metadata)
        data, result = dbx.files_download(path=path+FILEPATH)
        
        df = pd.DataFrame.from_records(files_list)
        
        return df.sort_values(by='server_modified', ascending=False),result
    
    except Exception as e:
        print('Error getting list of files from Dropbox: ' + str(e))


app = Flask(__name__)
@app.route('/webhook', methods=['POST'])

def webhook(*args):
    
    try:
        data = request.json  # Extract the JSON payload from the request
        # Process the webhook data as needed
        print("Webhook received:", data,type(data))
        meta,raw_docx = dropbox_getFiles('/Cours')
    
        with open(UPDATED_FILE_NAME, 'wb') as my_data:
            
            my_data.write(raw_docx.content)
            my_data.close()
    
        oldLinesList = docx2txt.process(SAMPLE_FILE_NAME).splitlines()
        newlinesList = docx2txt.process(UPDATED_FILE_NAME).splitlines()
        
        
        if oldLinesList != newlinesList:
            
            with open(SAMPLE_FILE_NAME, 'wb') as my_data:
                
                my_data.write(raw_docx.content)
                my_data.close()
                
            if len(oldLinesList) < len(newlinesList):
            
                addedTextList = newlinesList[len(oldLinesList)+1:]
                
                #print(addedTextList)       
                embed = DiscordEmbed(
                    title=UI_TITLE, 
                    description=textFormat(addedTextList), 
                    color=8411391,
                    )

                webhook.add_embed(embed)
                response = webhook.execute()
                
        return jsonify({"status": "success"})
    except Exception as e:
        print("Error processing webhook:", str(e))
        return jsonify({"status": "error", "error": str(e)})


@app.route('/webhook', methods=['GET'])

def challenge(*args):
    '''Respond to the webhook challenge (GET request) by echoing back the challenge parameter.'''

    resp = Response(request.args.get('challenge'))
    resp.headers['Content-Type'] = 'text/plain'
    resp.headers['X-Content-Type-Options'] = 'nosniff'

    return resp

if __name__ == '__main__':
    app.run() # host='0.0.0.0',port=443
