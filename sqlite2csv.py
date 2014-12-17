#    SQLITE3 to CSV Converter using Pandas
#    Copyright (C) 2014 Chapin Bryce
#	   You can find more code at github.com/chapinb
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

#    Version 20141216

__author__ = 'cb'

"""
#README

## Installation

This script can run from any directory at the commandline considering the 
following dependencies are met.

### Dependencies

Binaries
* python2.7

Python Libraries
* openpyxl
* pandas

"""

import os
import sqlite3
import pandas as pd
import datetime


def table_reader(table, con):
    """
    Reads data from a SQLITE3 Database Tables

    :param table: string, name of table to parse
    :param con:   object, sqlite3 object for table
    :return:      object, Pandas DataFrame Object
    """
    return pd.read_sql("SELECT * FROM "+table, con)


def xlsx_writer(writer, df, table):
    """
    Writes data to an XLSX spreadsheet

    :param writer: object, Pandas XLSX Writer instance
    :param df:     object, Pandas DataFrame object with table data
    :param table:  string, name of table to name spreadsheet
    :return:       None
    """
    try:
        df.to_excel(writer, sheet_name=table)
        writer.save()
    except:
        print "Error writing data to " + table + " likely due to a value error"

if __name__ == '__main__':
    import argparse

    # Set up arguments
    parser = argparse.ArgumentParser(description="Create XLSX Documents from SQLITE3 Databases",
                                     epilog="Created by CBRYCE")
    parser.add_argument('fin', help="Input SQLITE DB File")
    parser.add_argument('fout', help="Output XLSX File")

    # setup input and output files
    args = parser.parse_args()
    db = args.fin
    filename = args.fout
    writer = pd.ExcelWriter(filename, engine='openpyxl')

    # gather information about database
    fn = os.path.basename(db)
    con = sqlite3.connect(db)
    cursor = con.cursor()
    raw_tables = cursor.execute("SELECT name FROM sqlite_master where type='table';")
    tables = cursor.fetchall()

    # process tables
    for t in tables:
        for i in t:
            xlsx_writer(writer, table_reader(i, con), i)

    # Completed
    print "Successfully wrote " + str(len(tables)) + " tables to " + filename
