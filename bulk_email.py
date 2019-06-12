"""
Inspiration: https://github.com/cfedermann/mass-mailer

Resources:

mimetypes
http://naelshiab.com/tutorial-send-email-python/
https://www.anomaly.net.au/blog/constructing-multipart-mime-messages-for-sending-emails-in-python/
https://stackoverflow.com/questions/3902455/mail-multipart-alternative-vs-multipart-mixed/28833772

attachments
https://stackoverflow.com/questions/8456181/python-cant-send-attachment-files-through-email
https://stackoverflow.com/questions/28821487/multipart-email-pdf-attachment-blank
https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
https://stackoverflow.com/questions/38825943/mimemultipart-mimetext-mimebase-and-payloads-for-sending-email-with-file-atta

smtp startls authentication
https://www.afternerd.com/blog/how-to-send-an-email-using-python-and-smtplib/
https://stackabuse.com/how-to-send-emails-with-gmail-using-python/
"""

import argparse
import csv
import getpass
import smtplib
import subprocess
import time

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from pathlib import Path


def load_config(filename):
    _config = {}
    assert (filename.endswith(".ini"))
    with open(filename) as f:
        for line in f:
            if not line.startswith("#"):
                content = line.split("=")
                if len(content) == 2:
                    key = content[0].strip()
                    value = content[1].strip()
                    _config[key] = value

        # Check that all required keys are available.
        # for key in ("SMTP", "FROM", "SUBJECT"):
        #     assert(key in _config.keys())

        # # Fallback to default values for REPLY-TO and FIRST_LASTNAME if they
        # # are not available from the given configuration dictionary.
        # if not "REPLY-TO" in _config.keys():
        #     _config["REPLY-TO"] = _config["FROM"]

        if "FIRST_LASTNAME" not in _config.keys():
            _config["FIRST_LASTNAME"] = "Sir or Madam"
    return _config


def create_pdf_file(name, tex_template, output_folder):
    tex_template_text = Path(tex_template).read_text()

    # customize tex template
    tex_named = tex_template_text.replace('<NAME>', name)

    # create tex file in output folder
    tex_out = output_folder / "{}.tex".format(name.replace(" ", ""))
    tex_out.write_text(tex_named)

    # compile latex to pdf
    subprocess.call(["pdflatex", "-output-directory", output_folder, tex_out])


def create_certificate(recipient, tex_template, output_folder, _config):

    name = recipient.first_name + " " + recipient.last_name
    create_pdf_file(name, tex_template, output_folder)

    output_path = output_folder / "{}.tex".format(name.replace(" ", ""))
    recipient.attachment = output_path
    print("Created .pdf file in {}".format(output_path))


class recipient:
    def __init__(self, first_name, last_name, email):
        self.last_name = last_name
        self.first_name = first_name
        self.email = email
        self.attachment = None


def load_recipients(filename):
    """Read email recipients from .csv file.

    The file should contain a first name, last name and email column.
    Comma separation.

    Parameters
    ----------
    filename : str
        Path to the .csv file containing names and email addresses.

    Returns
    -------
    list
        List of recipient objects.
    """

    with open(filename) as f:
        reader = csv.reader(f, delimiter=",")
        next(reader, None)
        recipient_list = [
            recipient(line[0], line[1], line[2]) for line in reader
        ]
    return recipient_list


def load_template(filename):
    # assert(filename.endswith(".html"))
    with open(filename) as f:
        template = f.read()
    return template


def smtp_login(login, port):
    # connect to smtp mail server
    # https://stackoverflow.com/questions/44761676/smtp-ssl-sslerror-ssl-unknown-protocol-unknown-protocol-ssl-c590?rq=1
    if port == 465:
        server = smtplib.SMTP_SSL(_config["HOST"], port)
        server.ehlo()
    else:
        server = smtplib.SMTP(_config["HOST"], port)
        server.ehlo()
        server.starttls()
        server.ehlo()

    for i in range(3):
        try:
            pw = getpass.getpass(prompt="Password: ", stream=None)
            server.login(login, pw)
            return server
        except Exception as e:
            print(
                "Failed to login to smtp server. Please retype password or check port number and credentials."
            )
            lastException = e
            continue
    else:
        raise lastException


def send_email(_config, recipient_list, template, alt_template=None):
    port = int(_config["PORT"])
    login = _config["LOGIN"]

    server = smtp_login(login, port)

    # loop through all found recipients
    for recipient in recipient_list:

        # avoid overloading smtp server
        time.sleep(5)

        # start multipart message
        msg = MIMEMultipart()

        # set MIME information
        msg["Subject"] = _config["SUBJECT"]
        msg["From"] = _config["SENDER"]
        msg["To"] = recipient.email
        # msg.preamble = "RSG Belgium Student Symposium Certificate"

        # attach html body
        if alt_template is None:
            template_named = template.replace(
                "{{FIRST_LASTNAME}}",
                recipient.first_name + " " + recipient.last_name)
            msg.attach(MIMEText(template_named, "html"))
        else:
            # create multipart to store plain and html templates
            body = MIMEMultipart("alternative")

            alt_templated_named = alt_template.replace(
                "{{FIRST_LASTNAME}}",
                recipient.first_name + " " + recipient.last_name)
            body.attach(MIMEText(alt_templated_named, "plain"))

            template_named = template.replace(
                "{{FIRST_LASTNAME}}",
                recipient.first_name + " " + recipient.last_name)
            body.attach(MIMEText(template_named, "html"))

            # attach to main mixed multipart message
            msg.attach(body)

        # add attachment
        assert (recipient.attachment)
        with recipient.attachment.open("rb") as f:
            attachment = MIMEApplication(f.read(), "pdf")
            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename="certificate.pdf")
        msg.attach(attachment)

        # Encode email
        encoded_email = msg.as_string()
        server.sendmail(msg["From"], msg["To"], encoded_email)
        print("Message sent to {}.".format(recipient.email))

    server.quit()
    print("Closed connection to mail server.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-c", "--config", help=".ini file with configuration values to use")
    parser.add_argument(
        "-e",
        "--emails",
        help=".csv file (, separated) with names (col 1) and emails (col 2)")
    parser.add_argument(
        "-t", "--template", help="email template to use (.html)")
    parser.add_argument(
        "-a",
        "--alt_template",
        help="alternative (textual) template to use (.txt)",
        required=False)
    parser.add_argument(
        "-x", "--tex", help="LaTeX template to use (.tex)")
    parser.add_argument(
        "-o",
        "--out",
        help="output folder where to store generated pdfs",
        required=True)
    args = parser.parse_args()

    # retrieve configuration values, templates
    _config = load_config(args.config)
    email_template = load_template(args.template)
    alt_email_template = load_template(
        args.alt_template) if args.alt_template else None
    recipient_list = load_recipients(args.emails)

    if args.out:
        output_folder = Path(args.out).resolve()
        output_folder.mkdir(exist_ok=True)

        for recipient in recipient_list:
            create_certificate(recipient, args.tex, output_folder, _config)

    send_email(_config, recipient_list, email_template, alt_email_template)
