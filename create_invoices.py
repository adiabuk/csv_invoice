#!/usr/bin/env python
# pylint: disable=attribute-defined-outside-init,protected-access,too-many-instance-attributes

"""
Generate pdf invoices from csv spreadsheet and hardcoded config values
"""

import csv
import locale
import datetime
from configparser import ConfigParser
import pypandoc
from mailmerge import MailMerge

locale.setlocale(locale.LC_ALL, '')

class AttrDict(dict):
    """Class to allow dictinary keys to be specified as attributes"""

    def __init__(self, *args, **kwargs):
        super(AttrDict, self).__init__(*args, **kwargs)
        self.__dict__ = self

def get_config():
    """Get static config from config file"""

    parser = ConfigParser()
    parser.read("config.ini")
    return dict(parser._sections['main'])

def main():
    """ Main method """

    config_dict = {}
    config_dict.update(get_config())
    config = AttrDict(config_dict)

    filename = 'invoices.csv'
    data = open(filename, 'r')
    reader = csv.reader(data.readlines()[1:], delimiter=',')

    template = "template.docx"
    for row in reader:
        config.invoice_date, config.invoicename, config.total, config.flatno, \
                config.invoiceno, config.invoiceaddress, config.invoicecity, \
                config.invoicepostcode = row
        config.monthyear = datetime.datetime.strptime(config.invoice_date, '%d/%m/%Y') \
                .strftime("%b %Y")
        config.total = locale.currency(float(config.total), grouping=True)

        document = MailMerge(template)
        document.merge(**dict(config))
        document.write('{}.docx'.format(config.invoiceno))
        pypandoc.convert_file('{}.docx'.format(config.invoiceno), to='pdf', format='docx',
                              outputfile='{}.pdf'.format(config.invoiceno))
        print("Creating {}".format(config.invoiceno))
        #os.remove('{}.docx'.format(config.invoiceno))

if __name__ == '__main__':
    main()
