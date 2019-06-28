#!/usr/bin/env python
#pylint: disable=no-member,attribute-defined-outside-init,protected-access

"""
Create docx VAT invoice from template and given values
"""

from __future__ import print_function

import locale
from configparser import ConfigParser
from datetime import date
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

def prompt_for_config():
    """
    Prompt for input variables with will be substituted into template
    """
    results = {}
    fields = {'invoiceno': 'Invoice Number',
              'days': 'Days',
              'term': 'Term'}

    for name, text in fields.items():
        result = input(text + ": ")
        results[name] = result
    return results


def main():
    """Main function"""

    config_dict = {}
    config_dict.update(get_config())
    config_dict.update(prompt_for_config())
    config = AttrDict(config_dict)
    template = "template.docx"
    document = MailMerge(template)
    print(document.get_merge_fields())
    print(config.sort)
    config.subtotal = float(float(config.rate) * float(config.days))
    config.vatamount = float((float(config.vatrate) * float(config.subtotal))/100)
    config.total = float(float(config.subtotal) + float(config.vatamount))

    document.merge(
        sort=config.sort,
        acctname=config.acctname,
        rate=locale.currency(float(config.rate), grouping=True),
        vatamount=locale.currency(config.vatamount, grouping=True),
        days=config.days,
        subtotal=locale.currency(config.subtotal, grouping=True),
        email=config.email,
        term=config.term,
        bank=config.bank,
        services=config.services,
        date='{:%d-%b-%Y}'.format(date.today()),
        acctno=config.acctno,
        vatrate=config.vatrate,
        address=config.address,
        client=config.client,
        phone=config.phone,
        total=locale.currency(config.total, grouping=True),
        terms=config.terms,
        vatno=config.vatno,
        invoiceno=config.invoiceno,
        companyno=config.companyno


        )

    document.write('{}.docx'.format(config.invoiceno))
if __name__ == '__main__':
    main()
