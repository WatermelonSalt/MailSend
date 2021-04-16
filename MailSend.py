#! Python
# Send beautiful emails using Python
"""
Port number used for SSL connections is 465

# Build mail
message = MIMEMultipart("alternative")
message["Subject"] = <Subject for the mail>
message["From"] = <Sender mail id>
message["To"] = <Receiver mail id>
message["Bcc"] = <Bulk mail ids as a string>

# Mail content

<variable> = <Content as a string>

# To add attachments to a mail:

# Open PDF file in binary mode
with open(<filename>, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)

# Add message content to build

<variable> = MIMEText(<variable>, "plain") # For Plain Text
<variable> = MIMEText(<variable>, "html") # For Html content

message.attach(<variable>)
message.attach(<variable>)
message.attach(part) # To attach the attachment

# Create a secure SSL connection
context = ssl.create_default_context()

# Sending the email
with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:

    server.login(<sender>, <password>)
    # Send email
    server.sendmail(<sender>, <receiver>, message.as_string())

# How to send Bccs

toaddr = 'buffy@sunnydale.k12.ca.us'
cc = ['alexander@sunydale.k12.ca.us','willow@sunnydale.k12.ca.us']
bcc = ['chairman@slayerscouncil.uk']
fromaddr = 'giles@sunnydale.k12.ca.us'
message_subject = "disturbance in sector 7"
message_text = "Three are dead in an attack in the sewers below sector 7."
message = "From: %s\r\n" % fromaddr
        + "To: %s\r\n" % toaddr
        + "CC: %s\r\n" % ",".join(cc)
        + "Subject: %s\r\n" % message_subject
        + "\r\n"
        + message_text
toaddrs = [toaddr] + cc + bcc
server = smtplib.SMTP('smtp.sunnydale.k12.ca.us')
server.set_debuglevel(1)
server.sendmail(fromaddr, toaddrs, message)
server.quit()
"""

import getopt
import json
from json import JSONDecodeError
import os
import smtplib
from smtplib import SMTPAuthenticationError, SMTPRecipientsRefused
import ssl
import sys
from collections import Counter
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from colorama import Fore, init

# Enable colored output

init()

# Function to split the opts var into its options and arguments


def splitOptsVar(opts_var):

    options = []
    arguments = []

    for option, argument in opts_var:

        options.append(option)
        arguments.append(argument)

    return options, arguments


# Function to check is multiple instances of the same option is present


def checkMultipleSameOptions(optionlist):

    status = True

    allelementscount = Counter(optionlist)

    for elementcount in allelementscount.values():

        if elementcount > 1:

            status = False

            break

    return status


# Function to filter only help


def filterOnlyHelp(optionlist):

    global nohelp

    nohelp = False

    status = True

    if optionlist.count("-h") or optionlist.count("--help") is True:

        for index, option in enumerate(optionlist):

            if index == 0 and option in ("-h", "--help"):

                try:

                    status = False
                    optionlist[index + 1]

                except:

                    status = True

                    break

            if index != 0 and index == len(optionlist) - 1 and option in (
                    "-h", "--help"):

                status = False

                break

            if index != 0 and index != len(optionlist) - 1 and option in (
                    "-h", "--help"):

                status = False

                break

    else:

        status = False
        nohelp = True

    return status


# Function to parse argument to MailBuild


def mailBuildParsing(optionlist, argumentlist):

    global mailbuild
    global MailBuildArgument

    for index, option in enumerate(optionlist):

        if option in ("-c", "--config"):

            print(
                f"{Fore.GREEN}Found option {Fore.MAGENTA}{option}{Fore.GREEN} with argument {Fore.MAGENTA}{argumentlist[index]}"
            )

            mailbuild = True
            MailBuildArgument = argumentlist[index]

            break

        else:

            mailbuild = False


# Function to provide help


def giveHelp():

    helptext = f"""\
{Fore.MAGENTA}MailSend Help
{Fore.YELLOW}-------------

{Fore.GREEN}Usage: MailSend [Options]

{Fore.RED}Order of options does not matter

{Fore.GREEN}Options:


{Fore.MAGENTA}Short        Long        Action
{Fore.YELLOW}-----        ----        ------

{Fore.CYAN}-c           --config    This option takes in a path to the config file which is used to send the emails

{Fore.CYAN}-h           --help      This option shows this help message

{Fore.MAGENTA}Arguments
{Fore.YELLOW}---------

{Fore.CYAN}-c <path/to/your/config.json>    {Fore.RED}Must include the extension

{Fore.CYAN}-h {Fore.RED}No Arguments
"""

    print(helptext)


# Function to handle the options got from the getopt method


def handleOptions(argv):

    print(f"{Fore.GREEN}Analysing the options provided to the program...")

    try:

        opts, errs = getopt.getopt(argv[1:],
                                   shortopts="c:h",
                                   longopts=["config=", "help"])

    except Exception as exp:

        print(f"{Fore.RED}{exp}\n\n{Fore.BLUE}The program will exit now")
        sys.exit()

    options, arguments = splitOptsVar(opts)

    if errs != []:

        print(f"{Fore.RED}Options were not passed correctly!")
        print(f"{Fore.GREEN}Please try again with proper options")
        sys.exit()

    if opts == []:

        print(
            f"{Fore.BLUE}Hmm, seems like you don't know how to use me\nWell, here is some help"
        )
        giveHelp()
        sys.exit()

    if checkMultipleSameOptions(options) is True:

        pass

    else:

        print(
            f"{Fore.RED}You seem to use an option more than once which is not allowed.\nThe program will now exit"
        )
        sys.exit()

    if filterOnlyHelp(options) is True:

        giveHelp()

    elif nohelp == True:

        pass

    else:

        print(
            f"{Fore.RED}You seem to use other options along with the 'help' option which should not be done.\nThe program will now exit"
        )
        sys.exit()

    mailBuildParsing(options, arguments)


# Mail Builder class


class MailBuild:

    # Initialize required variables

    def __init__(self):

        self.bcctypeset = ""
        self.cctypeset = ""
        self.message = MIMEMultipart("alternative")

    # Extract the information from the configuration file

    def decodeConfig(self, configloc):

        print(
            f"{Fore.GREEN}Trying to get the config file from {Fore.MAGENTA}{configloc}"
        )

        try:

            self.configobj = open(configloc, "r")
            self.config = self.configobj.read()
            self.configobj.close()

        except FileNotFoundError:

            print(
                f"{Fore.RED}Seems like I couldn't find the file specified..\nThe program will exit now"
            )
            sys.exit()

        print(f"{Fore.GREEN}Got the config file successfully, yay!")
        print(f"{Fore.CYAN}Trying to load the config file into the program")

        try:

            self.config = json.loads(self.config)

        except JSONDecodeError:

            print(
                f"{Fore.RED}Seems like there is an error in your JSON encoding, please correct it and re-execute the program.\nThe program will exit now"
            )
            sys.exit()

        print(
            f"{Fore.GREEN}Loaded the config file successfully into the program, yay!"
        )
        print(f"{Fore.GREEN}Trying to get the subject from the config file...")

        try:

            self.Subject = self.config["Subject"]
            print(f"{Fore.GREEN}Got 'Subject' successfully, yay!")

        except KeyError:

            print(
                f"{Fore.RED}'Subject' key was not found, proceeding to next key.\n{Fore.CYAN}'Subject' will be automatically set as 'No Subject'"
            )
            self.Subject = "No Subject"

        try:

            self.From = self.config["From"]
            print(f"{Fore.GREEN}Got 'From' successfully, yay!")

            if "@gmail.com" not in self.From:

                print(
                    f"{Fore.RED}Aw, snap! Seems like it's not a Gmail address. I only support sending mails through gmail accounts.\nThe program will exit now"
                )
                sys.exit()

        except KeyError:

            print(
                f"{Fore.RED}'From' key was not found, can't proceed to the next key.\nProgram will exit now"
            )
            sys.exit()

        try:

            self.Password = self.config["Password"]
            print(f"{Fore.GREEN}Got 'Password' successfully, yay!")

        except KeyError:

            print(
                f"{Fore.RED}'Password' key was not found, can't proceed to the next key.\nProgram will exit now"
            )
            sys.exit()

        try:

            self.To = self.config["To"]

        except KeyError:

            print(
                f"{Fore.RED}'To' key was not found, can't proceed to the next key.\nProgram will exit now"
            )
            sys.exit()

        try:

            self.CC = self.config["CC"]

        except KeyError:

            print(
                f"{Fore.RED}'CC' key was not found, proceeding to next key.\n{Fore.CYAN}'CC' will automatically be assigned to an empty list"
            )

            self.CC = []

        try:

            self.Bcc = self.config["Bcc"]

        except KeyError:

            print(
                f"{Fore.RED}'Bcc' key was not found, proceeding to next key.\n{Fore.CYAN}'Bcc' will automatically be assigned to an empty list"
            )

            self.Bcc = []

        try:

            self.TextContent = self.config["TextContent"]

        except KeyError:

            print(
                f"{Fore.RED}'TextContent' key was not found, proceeding to next key.\n{Fore.CYAN}'TextContent' will automatically be assigned to an empty text file with the name 'Dummy.txt' but will be deleted after sending the mail"
            )

            with open("Dummy.txt", "w+") as dummy:

                dummy.write("")

            self.TextContent = "Dummy.txt"

        try:

            self.HTMLContent = self.config["HTMLContent"]

        except KeyError:

            print(
                f"{Fore.RED}'HTMLContent' key was not found, proceeding to next key.\n{Fore.CYAN}'HTMLContent' will automatically be assigned to an empty HTML file with the name 'Dummy.html' but will be deleted after sending the mail"
            )

            with open("Dummy.html", "w+") as dummy:

                dummy.write("")

            self.HTMLContent = "Dummy.html"

        try:

            self.Attachments = self.config["Attachments"]

        except KeyError:

            print(
                f"{Fore.RED}'Attachments' key was not found, proceeding to next key.\n{Fore.CYAN}'Attachments' will automatically be assigned to an empty list"
            )

            self.Attachments = []

    # Build the mail

    def builder(self):

        self.message["Subject"] = self.Subject
        self.message["From"] = self.From
        self.message["To"] = self.To

        for element in self.CC:

            self.cctypeset += element + ","

        self.cctypeset = self.cctypeset[:len(self.cctypeset) - 1]
        self.message["Cc"] = self.cctypeset

        for element in self.Bcc:

            self.bcctypeset += element + ","

        self.bcctypeset = self.bcctypeset[:len(self.bcctypeset) - 1]
        self.message["Bcc"] = self.bcctypeset

        self.message.attach(MIMEText(open(self.TextContent).read(), "plain"))
        self.message.attach(MIMEText(open(self.HTMLContent).read(), "html"))
        self.attachmentAdder()

    # Add attachments to the mail

    def attachmentAdder(self):

        for attachment in self.Attachments:

            try:

                with open(attachment, "rb") as attfile:

                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attfile.read())

                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename = {os.path.basename(attachment)}")
                self.message.attach(part)

            except FileNotFoundError:

                print(
                    f"{Fore.RED}Seems like I couldn't find the file specified..\nThe program will exit now"
                )
                sys.exit()

    # Use the MailBuilder class through this method

    def execute(self, configloc):

        self.decodeConfig(configloc)
        self.builder()

        return self.message, self.From, self.Password, [self.To
                                                        ] + self.CC + self.Bcc

    # Delete unnecessary stuff

    def __del__(self):

        try:

            os.remove("Dummy.txt")
            print(f"{Fore.YELLOW}Deleted temporarily created file 'Dummy.txt'")

        except FileNotFoundError:

            pass

        try:

            os.remove("Dummy.html")
            print(
                f"{Fore.YELLOW}Deleted temporarily created file 'Dummy.html'")

        except FileNotFoundError:

            pass


# Function to send the mail generated by MailBuilder class


def mailSender(message, sender, password, receivers):

    port = 465
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:

        server.login(sender, password)
        server.sendmail(sender, receivers, message.as_string())


# If the script is not used as a module, call MailBuild
if __name__ == "__main__":

    handleOptions(sys.argv)

    if mailbuild is True:

        print(
            f"{Fore.YELLOW}Trying to build the mail from the data provided in the config file..."
        )

        builder = MailBuild()

        message, sender, password, receivers = builder.execute(
            MailBuildArgument)

        print(f"{Fore.YELLOW}Mail built successfully, yay!")
        print(f"{Fore.CYAN}Trying to send the built mail...")

        try:

            mailSender(message, sender, password, receivers)

        except SMTPAuthenticationError:

            print(
                f"{Fore.RED}Oh, seems like the credentials were wrong, please check again and re-execute the program.\nThe program will exit now"
            )
            sys.exit()

        except SMTPRecipientsRefused:

            print(
                f"{Fore.RED}Oh, seems like the recepients' addresses were wrong, please check again and re-execute the program.\nThe program will exit now"
            )
            sys.exit()

        print(
            f"{Fore.GREEN}Seems like everything went as planned, sent mails successfully, yay!"
        )
