#! Python
# Send beautiful emails using Python

import getopt
import json
import os
import smtplib
import ssl
import sys
from collections import Counter
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from json import JSONDecodeError
from smtplib import SMTPAuthenticationError, SMTPRecipientsRefused

from rich.console import Console
from rich.panel import Panel
from rich import box

# Setting up the console dimensions

os.system("mode con: cols=125 lines=30")

# Initializing the console

console = Console(color_system="auto")

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

            console.print(
                f"[green]Found option[/green] [magenta]{option}[/magenta] [green]with argument[/green] [magenta]{argumentlist[index]}[/magenta]"
            )

            mailbuild = True
            MailBuildArgument = argumentlist[index]

            break

        else:

            mailbuild = False


# Function to provide help


def giveHelp():

    helptext = """\
[magenta]MailSend Help[/magenta]
[yellow]-------------[/yellow]

[green]Usage: MailSend[/green] [cyan][Options][/cyan]

[red]Order of options does not matter[/red]

[green]Options:[/green]


[magenta]Short        Long        Action[/magenta]
[yellow]-----        ----        ------[/yellow]

[cyan]-c           --config    This option takes in a path to the config file which is used to send the emails[/cyan]

[cyan]-h           --help      This option shows this help message[/cyan]

[magenta]Arguments[/magenta]
[yellow]---------[/yellow]

[cyan]-c <path/to/your/config.json>[/cyan]    [red]Must include the extension[/red]

[cyan]-h[/cyan]    [red]No Arguments[/red]
"""

    console.print(Panel.fit(helptext, box=box.DOUBLE), justify="center")


# Function to handle the options got from the getopt method


def handleOptions(argv):

    console.print(
        "[green]Analysing the options provided to the program...[/green]")

    try:

        opts, errs = getopt.getopt(argv[1:],
                                   shortopts="c:h",
                                   longopts=["config=", "help"])

    except Exception as exp:

        console.print(
            f"[red]{exp}[/red]\n\n[blue]The program will exit now[/blue]")
        sys.exit()

    options, arguments = splitOptsVar(opts)

    if errs != []:

        console.print("[red]Options were not passed correctly![/red]")
        console.print("[green]Please try again with proper options[/green]")
        sys.exit()

    if opts == []:

        console.print(
            "[blue]Hmm, seems like you don't know how to use me\nWell, here is some help[/blue]"
        )
        giveHelp()
        sys.exit()

    if checkMultipleSameOptions(options) is True:

        pass

    else:

        console.print(
            "[red]You seem to use an option more than once which is not allowed.\nThe program will now exit[/red]"
        )
        sys.exit()

    if filterOnlyHelp(options) is True:

        giveHelp()

    elif nohelp == True:

        pass

    else:

        console.print(
            "[red]You seem to use other options along with the 'help' option which should not be done.\nThe program will now exit[/red]"
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

        console.print(
            f"[green]Trying to get the config file from[/green] [magenta]{configloc}[/magenta]"
        )

        try:

            self.configobj = open(configloc, "r")
            self.config = self.configobj.read()
            self.configobj.close()

        except FileNotFoundError:

            console.print(
                "[red]Seems like I couldn't find the file specified..\nThe program will exit now[/red]"
            )
            sys.exit()

        console.print("[green]Got the config file successfully, yay![/green]")
        console.print(
            "[cyan]Trying to load the config file into the program[/cyan]")

        try:

            self.config = json.loads(self.config)

        except JSONDecodeError:

            console.print(
                "[red]Seems like there is an error in your JSON encoding, please correct it and re-execute the program.\nThe program will exit now[/red]"
            )
            sys.exit()

        console.print(
            "[green]Loaded the config file successfully into the program, yay![/green]"
        )
        console.print(
            "[green]Trying to get the subject from the config file...[/green]")

        try:

            self.Subject = self.config["Subject"]
            console.print("[green]Got 'Subject' successfully, yay![/green]")

        except KeyError:

            console.print(
                "[red]'Subject' key was not found, proceeding to next key.[/red]\n[cyan]'Subject' will be automatically set as 'No Subject'[/cyan]"
            )
            self.Subject = "No Subject"

        try:

            self.From = self.config["From"]
            console.print("[green]Got 'From' successfully, yay![/green]")

            if "@gmail.com" not in self.From:

                console.print(
                    "[red]Aw, snap! Seems like it's not a Gmail address. I only support sending mails through gmail accounts.\nThe program will exit now[/red]"
                )
                sys.exit()

        except KeyError:

            console.print(
                "[red]'From' key was not found, can't proceed to the next key.\nProgram will exit now[/red]"
            )
            sys.exit()

        try:

            self.Password = self.config["Password"]
            console.print("[green]Got 'Password' successfully, yay![/green]")

        except KeyError:

            console.print(
                "[red]'Password' key was not found, can't proceed to the next key.\nProgram will exit now[/red]"
            )
            sys.exit()

        try:

            self.To = self.config["To"]

        except KeyError:

            console.print(
                "[red]'To' key was not found, can't proceed to the next key.\nProgram will exit now[/red]"
            )
            sys.exit()

        try:

            self.CC = self.config["CC"]

        except KeyError:

            console.print(
                "[red]'CC' key was not found, proceeding to next key.[/red]\n[cyan]'CC' will automatically be assigned to an empty list[/cyan]"
            )

            self.CC = []

        try:

            self.Bcc = self.config["Bcc"]

        except KeyError:

            console.print(
                "[red]'Bcc' key was not found, proceeding to next key.[/red]\n[cyan]'Bcc' will automatically be assigned to an empty list[/cyan]"
            )

            self.Bcc = []

        try:

            self.TextContent = self.config["TextContent"]

        except KeyError:

            console.print(
                "[red]'TextContent' key was not found, proceeding to next key.[/red]\n[cyan]'TextContent' will automatically be assigned to an empty text file with the name 'Dummy.txt' but will be deleted after sending the mail[/cyan]"
            )

            with open("Dummy.txt", "w+") as dummy:

                dummy.write("")

            self.TextContent = "Dummy.txt"

        try:

            self.HTMLContent = self.config["HTMLContent"]

        except KeyError:

            console.print(
                "[nred]'HTMLContent' key was not found, proceeding to next key.[/red]\n[cya]'HTMLContent' will automatically be assigned to an empty HTML file with the name 'Dummy.html' but will be deleted after sending the mail[/cyan]"
            )

            with open("Dummy.html", "w+") as dummy:

                dummy.write("")

            self.HTMLContent = "Dummy.html"

        try:

            self.Attachments = self.config["Attachments"]

        except KeyError:

            console.print(
                "[red]'Attachments' key was not found, proceeding to next key.[/red]\n[cyan]'Attachments' will automatically be assigned to an empty list[/cyan]"
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

                console.print(
                    "[red]Seems like I couldn't find the file specified..\nThe program will exit now[/red]"
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
            console.print(
                "[yellow]Deleted temporarily created file 'Dummy.txt'[/yellow]")

        except FileNotFoundError:

            pass

        try:

            os.remove("Dummy.html")
            console.print(
                "[yellow]Deleted temporarily created file 'Dummy.html'[/yellow]")

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

        console.print(
            "[yellow]Trying to build the mail from the data provided in the config file...[/yellow]"
        )

        builder = MailBuild()

        message, sender, password, receivers = builder.execute(
            MailBuildArgument)

        console.print("[yellow]Mail built successfully, yay![/yellow]")
        console.print("[cyan]Trying to send the built mail...[/cyan]")

        try:

            mailSender(message, sender, password, receivers)

        except SMTPAuthenticationError:

            console.print(
                "[red]Oh, seems like the credentials were wrong, please check again and re-execute the program.\nThe program will exit now[/red]"
            )
            sys.exit()

        except SMTPRecipientsRefused:

            console.print(
                "[red]Oh, seems like the recepients' addresses were wrong, please check again and re-execute the program.\nThe program will exit now[/red]"
            )
            sys.exit()

        console.print(
            "[green]Seems like everything went as planned, sent mails successfully, yay![/green]"
        )
