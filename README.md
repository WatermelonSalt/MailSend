# MailSend - A Python script to send emails to multiple people via [gmail](smtp.gmail.com)

## Introduction

MailSend is a script written in Python which is capable of sending HTML formatted emails
A configuration file is required in order to send emails which will be explained later in this file
For MailSend to work properly [**"Less Secure Apps"**](https://myaccount.google.com/lesssecureapps) must be enabled

## Dependencies

MailSend has only one dependency which is **colorama** which can be installed using pip by executing the following command in the terminal
`pip install colorama`
The above mentioned module is only used for enabling colored output in the terminal and is not optional

## Requirements

* A Python3 installation (This was checked with Python 3.9.4 and it works. Versions below this version have not been checked and it would be great if someone could help me with the compatibility checking)
* A dummy gmail account with **"less secure apps"** enabled

## Writing the configuration file

A sample configuration file with all available options has been included with the repo and should follow the JSON rules
Please take a look at the provided example configuration file to understand better

### Option Explanation:

**Subject** - This key is optional and can contain a string enclosed within double quotes. If not present "No Subject" will be used automatically

**From** - This key must be present and should only contain a valid **gmail** address and should be enclosed within double quotes

**Password** - This key must be present and should contain the password for the above mentioned **gmail** address and should be enclosed within double quotes

**To** - This key must be present and should contain only one valid email address of any email provider and must be enclosed within double quotes

**CC** - This key is optional and can contain any number of valid email addresses to be sent as carbon copy enclosed within double quotes separated by commas. If not present an empty list will be used automatically

**Bcc** - This key is optional and can contain any number of valid email addresses to be sent as blind carbon copy enclosed within double quotes separated by commas. If not present an empty list will be used automatically

**TextContent** - This key is optional and should contain the path to a single plaintext file containing the data to be attached to the top of the mail body and the path should be enclosed within double quotes

**HTMLContent** - This key is optional and should contain the path to a single html file containing the data to be attached to the mail body after the text content (if present) and the path should be enclosed within double quotes

**Attachments** - This key should contain a list of comma separated paths to the attachment files to be added to the email and must fall within 25 Mb in total and the paths must be enclosed within double quotes

## How to use the script

You can do `python MailSend.py -h` to get help regarding the program's options and arguments accepted by the options

## Things you should know

* This is Open-Source and you can modify it to your needs
* Uses a GPL 3.0 License
* An example config.json is also included which serves as a reference when writing your config.json
* This script can also be bundled into an executable binary using [Pyinstaller](https://pypi.org/project/pyinstaller/)

## Updates

Updates will be made to this repo by me if time permits

## Socials

[![Discord](https://img.shields.io/badge/Discord-Visit-ff5500?labelColor=72269e&style=for-the-badge&logo=Discord&logoColor=#ffe600&logoWidth=20&link=https://discordapp.com/users/799197702281494601)](https://discordapp.com/users/799197702281494601)
[![Spotify](https://img.shields.io/badge/Spotify-Listen-ff5500?labelColor=72269e&style=for-the-badge&logo=Spotify&logoWidth=20&link=https://open.spotify.com/playlist/4ngYgsMMKtMenzkyTczRHP)](https://open.spotify.com/playlist/4ngYgsMMKtMenzkyTczRHP)
