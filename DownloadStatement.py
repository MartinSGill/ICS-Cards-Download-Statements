# coding=utf-8
"""
A script that is able to parse the ICS Cards ABN Amro site
and download statements for a specific month and save them
as CSV.

It does some simple parsing of the statement entries to
make the CSV data more useful for importing into a
financial planning application.

For options/help:::

    > python DownloadStatement.py --help

.. Note::
   * Only tested with the ABN Amro "theme" for icscards.nl
   * Only tested with an account with a single card
"""

__author__ = 'Martin Gill'
__copyright__ = 'Copyright 2014, Martin Gill'
__license__ = "MIT"
__version__ = "1.0.0"

import csv
import re
import sys
import datetime
import argparse

from bs4 import BeautifulSoup
import requests


class StatementReader:
    """
    Logs into the ICS Website and retrieves statements,
    processes them and saves them as CSV.
    """
    __baseUrl = "https://www.icscards.nl/abnamro/mijn/accountstatements?period="
    __loginUrl = "https://www.icscards.nl/pkmslogin.form"
    __session = None
    __username = ""
    __password = ""
    __headers = []
    # Datum,*,Omschrijving,Card-nummer,Debet / Credit,Valuta,Bedrag,payee,out,in
    __entries = []

    def __init__(self, username, password):
        self.__username = username
        self.__password = password

    def get_statement(self, period):
        """
        Download the statement for the specified time period.

        :parameter period: Time period of statement to retrieve.
        """
        self.__login()
        url = self.__baseUrl + period.__str__()
        print("Downloading page for {0:04d}-{1:02d}".format(period.year, period.month))
        r = self.__session.get(url)
        data = r.content
        soup = BeautifulSoup(data)
        tables = soup.select("table.expander-table")
        self.__parse_table(tables)

    def to_csv(self, filename):
        """
        Write the downloaded statement to CSV format.

        :param filename: Name of the file to create.
        """
        print('Writing CSV to: "{0}"'.format(filename))
        with open(filename, 'w', newline='') as csv_file:
            writer = csv.writer(csv_file,
                                delimiter=',',
                                quotechar='"',
                                quoting=csv.QUOTE_MINIMAL)
            writer.writerow(self.__headers)
            writer.writerows(self.__entries)

    def to_qif(self, filename):
        """
        Write the downloaded statement to QIF format.

        :param filename: Name of the file to create.
        """

        print('Writing QIF to: "{0}"'.format(filename))
        with open(filename, 'w') as qif_file:
            qif_file.write("!Account\n")
            qif_file.write("N ICS Credit Card\n")
            qif_file.write("^\n")
            qif_file.write("!Type:CCard\n")

            for line in self.__entries:
                qif_file.write("D{0}\n".format(line[0]))
                qif_file.write("T{0}\n".format(line[6]))
                qif_file.write("M{0}\n".format(line[5]))
                qif_file.write("P{0}\n".format(line[7]))
                qif_file.write("^\n")

    def __get_headers(self, table):
        print('Getting Headers')
        headers = table.select("th")
        for th in headers:
            self.__headers.append(th.getText())

        self.__headers.append('payee')
        self.__headers.append('out')
        self.__headers.append('in')

    def __get_entry(self, row):
        cols = row.select("td")
        row_items = []
        for col in cols:
            # Replace newlines with whitespace
            text = col.getText().replace("\n", " ").strip()
            row_items.append(text)

        regex = re.compile(r"(.+?)\s\w{3}\s+Land:")
        # skips the "header" row
        if len(row_items) > 0:
            # Replace commas with dots (Dutch decimal to English)
            row_items[5] = row_items[5].replace(',', '.')
            # Fix comma and eliminate currency symbol and other crap
            row_items[6] = re.sub(r"^.+\s(\d+),(\d+).*", r"\1.\2", row_items[6])

            # Extract the Payee
            match = regex.search(row_items[2])
            if match:
                row_items.append(match.group(1))
            else:
                row_items.append(row_items[2])

            # Credit/Debit handling
            if re.match(r"Debet", row_items[4]):
                # Purchase
                row_items.append(row_items[6])
                row_items.append('')
                row_items[6] = '-' + row_items[6]
            else:
                # Credit
                row_items.append('')
                row_items.append(row_items[6])

            # Store the entry
            self.__entries.append(row_items)
            print('.', end='')

    def __parse_table(self, tables):
        print('Parsing Table')
        if len(tables) < 1:
            raise Exception('No statement table found')
        elif len(tables) > 1:
            raise Exception('Too many statement tables found')

        table = tables[0]
        if len(self.__headers) == 0:
            self.__get_headers(table)

        rows = table.select("tr")
        print('Getting Entries')
        for row in rows:
            self.__get_entry(row)

        print()
        print("All entries retrieved. {0:d} Found.".format(len(self.__entries)))

    def __login(self):
        if self.__session is None:
            print('Logging "{0}" into site'.format(self.__username))
            self.__session = requests.Session()
            payload = {'username': self.__username, 'password': self.__password, 'login-form-type': 'pwd'}
            r = self.__session.post(self.__loginUrl, payload, verify=False)
            data = BeautifulSoup(r.content).getText().strip()
            if data.strip() == "login_success":
                print("Login Succeeded")
            else:
                raise Exception("Login Failure")


class Period:
    """
    Represents a statement period.
    """

    def increment_month(self):
        """
        Increments the month by one, incrementing the year if needed.
        """
        if self.__month == 12:
            self.__month = 1
            self.__year += 1
        else:
            self.__month += 1

    def __init__(self, month, year):
        """
        Create a new statement period.

        :param: month: Period month.
        :param: year: Period year
        """
        self.__year = year
        self.__month = month

    @property
    def month(self):
        """
        The month this Period represents.
        """
        return self.__month

    @property
    def year(self):
        """
        The year this Period represents.
        """
        return self.__year

    def __str__(self):
        """
        Override standard method to provide format required by statement URL.
        """
        return str.format("{0:04d}{1:02d}", self.year, self.month)


def main():
    """
    Main method for the script.
    """
    month = datetime.date.today().month
    year = datetime.date.today().year
    end_month = None
    end_year = None

    parser = argparse.ArgumentParser(description="Downloads statements from ICS-Cards website and outputs them as csv.")
    parser.add_argument("username", help="Login Username.")
    parser.add_argument("password", help="Login Password.")
    parser.add_argument("-m", "--month", help="The statement month required. Current month if omitted.", type=int)
    parser.add_argument("-y", "--year", help="The statement year required. Current year if omitted.", type=int)

    end_group = parser.add_argument_group()
    end_group.add_argument("--end-month", help="The last statement month required.", type=int)
    end_group.add_argument("--end-year", help="The last statement year required.", type=int)

    file_group = parser.add_argument_group()
    file_group.add_argument("-f", "--filename", help="Output filename. 'yyyy-mm' if omitted.")
    file_group.add_argument("-t",
                            "--format",
                            help="The output format of the file. Default is QIF",
                            default="qif",
                            type=str,
                            choices=["csv", "qif"])
    args = parser.parse_args()

    if args.month is not None:
        month = args.month

    if args.year is not None:
        year = args.year

    if args.end_month is not None:
        end_month = args.end_month

    if args.end_year is not None:
        end_year = args.end_year

    if args.filename is None:
        if month == end_month and year == end_year:
            filename = "{0:04d}-{1:02d}.{2}".format(year, month, args.format)
        else:
            filename = "{0:04d}-{1:02d}_to_{2:04d}-{3:02d}.{4}".format(year, month, end_year, end_month, format)
    else:
        if args.filename.endswith(".{0}".format(args.format)):
            filename = args.filename
        else:
            filename = args.filename + ".{0}".format(args.format)

    # #  noinspection PyBroadException
    try:
        my_statement = StatementReader(args.username, args.password)
        period = Period(month, year)
        my_statement.get_statement(period)

        if end_month is not None and end_year is not None:
            while not (period.month == end_month and period.year == end_year):
                period.increment_month()
                my_statement.get_statement(period)

        if args.format == "qif":
            my_statement.to_qif(filename)
        else:
            my_statement.to_csv(filename)

    except Exception as ex:
        print("ERROR: {0}", str(ex), file=sys.stderr)
        exit(-1)


if __name__ == "__main__":
    main()
