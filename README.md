Simple python script that can send emails in bulk while generating pdf attachments. One of the intended scenarios is sending out certificates of attendence to a list of people at an event.

```
usage: bulk_email.py [-h] [-c CONFIG] [-e EMAILS] [-t TEMPLATE]
                     [-a ALT_TEMPLATE] [-o OUT]

optional arguments:
  -h, --help            show this help message and exit
  -c CONFIG, --config CONFIG
                        .ini file with configuration values to use
  -e EMAILS, --emails EMAILS
                        .csv file (, separated) with names (col 1) and emails
                        (col 2)
  -t TEMPLATE, --template TEMPLATE
                        email template to use (.html)
  -a ALT_TEMPLATE, --alt_template ALT_TEMPLATE
                        alternative (textual) template to use (.txt)
  -o OUT, --out OUT     output folder where to store generated pdfs
```

The `config.ini` file ([example provided](example_config.ini)) should hold the necessary details (e.g. outgoing smtp server, details for the pdf attachment, etc.).

---

### To do list

- generate pdf from template instead of hard coding
- pylatex instead of fpdf
- option to skip attachments
- option to attach existing files instead of generating (referenced in csv or based on name)
- allow smtp without login
- switch to yaml / configparser to allow comments inside config file
