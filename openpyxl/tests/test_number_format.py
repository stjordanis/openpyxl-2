from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

# Python stdlib imports
from datetime import datetime, date, timedelta

# 3rd party imports
from nose.tools import eq_, assert_almost_equal, assert_raises

# package imports
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.cell import Cell
from openpyxl.style import NumberFormat
from openpyxl.shared.date_time import SharedDate, CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

import time

# strptime fallback, thanks to coderfi
# http://stackoverflow.com/questions/5585706/datetime-datetime-strptime-not-present-in-python-2-4-1/7226819#7226819
if hasattr(datetime, 'strptime'):
    #python 2.6
    strptime = datetime.strptime
else:
    #python 2.4 equivalent
    strptime = lambda date_string, format: datetime(*(time.strptime(date_string, format)[0:6]))


class TestNumberFormat(object):

    @classmethod
    def setup_class(cls):
        cls.workbook = Workbook()
        cls.worksheet = Worksheet(cls.workbook, 'Test')
        cls.sd = SharedDate()

    def test_convert_date_to_julian(self):
        eq_(40167, self.sd.to_julian(2009, 12, 20))

    def test_convert_date_from_julian(self):

        def test_date_equal(julian, datetime):

            eq_(self.sd.from_julian(julian), datetime)

        date_pairs= (
                        (40167, datetime(2009, 12, 20)),
                        (21980, datetime(1960,  3,  5)),
                    )

        for count, dt in date_pairs:
            yield test_date_equal, count, dt

    def test_convert_datetime_to_julian(self):
        eq_(40167, self.sd.datetime_to_julian(datetime(2009, 12, 20)))
        eq_(40196.5939815, self.sd.datetime_to_julian(datetime(2010, 1, 18, 14, 15, 20, 1600)))

    def test_convert_timedelta_to_julian(self):
        eq_(1.125, self.sd.datetime_to_julian(timedelta(days=1, hours=3)))

    def test_insert_float(self):
        self.worksheet.cell('A1').value = 3.14
        eq_(Cell.TYPE_NUMERIC, self.worksheet.cell('A1')._data_type)

    def test_insert_percentage(self):
        self.worksheet.cell('A1').value = '3.14%'
        eq_(Cell.TYPE_NUMERIC, self.worksheet.cell('A1')._data_type)
        assert_almost_equal(0.0314, self.worksheet.cell('A1').value)

    def test_insert_datetime(self):
        self.worksheet.cell('A1').value = date.today()
        eq_(Cell.TYPE_NUMERIC, self.worksheet.cell('A1')._data_type)

    def test_insert_date(self):
        self.worksheet.cell('A1').value = datetime.now()
        eq_(Cell.TYPE_NUMERIC, self.worksheet.cell('A1')._data_type)

    def test_internal_date(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.worksheet.cell('A3').value = dt
        eq_(40372.27616898148, self.worksheet.cell('A3')._value)

    def test_datetime_interpretation(self):
        dt = datetime(2010, 7, 13, 6, 37, 41)
        self.worksheet.cell('A3').value = dt
        eq_(dt, self.worksheet.cell('A3').value)

    def test_date_interpretation(self):
        dt = date(2010, 7, 13)
        self.worksheet.cell('A3').value = dt
        eq_(datetime(2010, 7, 13, 0, 0), self.worksheet.cell('A3').value)

    def test_number_format_style(self):
        self.worksheet.cell('A1').value = '12.6%'
        eq_(NumberFormat.FORMAT_PERCENTAGE, \
                self.worksheet.cell('A1').style.number_format.format_code)

    def test_date_format_on_non_date(self):
        cell = self.worksheet.cell('A1')

        def check_date_pair(count, date_string):
            cell.value = strptime(date_string, '%Y-%m-%d')
            eq_(count, cell._value)

        date_pairs = (
            (15, '1900-01-15'),
            (59, '1900-02-28'),
            (61, '1900-03-01'),
            (367, '1901-01-01'),
            (2958465, '9999-12-31'), )
        for count, date_string in date_pairs:
            yield check_date_pair, count, date_string

    def test_1900_leap_year(self):
        assert_raises(ValueError, self.sd.from_julian, 60)
        assert_raises(ValueError, self.sd.to_julian, 1900, 2, 29)

    def test_bad_date(self):

        def check_bad_date(year, month, day):
            assert_raises(ValueError, self.sd.to_julian, year, month, day)

        bad_dates = ((1776,  7,  4), (1899, 12, 31), )
        for year, month, day in bad_dates:
            yield check_bad_date, year, month, day

    def test_bad_julian_date(self):
        assert_raises(ValueError, self.sd.from_julian, -1)

    def test_mac_date(self):
        self.sd.excel_base_date = CALENDAR_MAC_1904

        datetuple = (2011, 10, 31)

        dt = date(datetuple[0],datetuple[1],datetuple[2])
        julian = self.sd.to_julian(datetuple[0],datetuple[1],datetuple[2])
        reverse = self.sd.from_julian(julian).date()
        eq_(dt,reverse)
        self.sd.excel_base_date = CALENDAR_WINDOWS_1900
