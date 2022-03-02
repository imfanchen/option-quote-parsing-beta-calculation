import os
import re
import copy
import pandas as pd
from datetime import datetime
from enum import Enum


class OptionType(Enum):
    Call = 1
    Put = 2


class OptionQuote:
    def __init__(self, date, time, firm, expiration, option_type, strike_px, strike_spd, bid_px, ask_px, delta, implied_vol_spd, implied_vol_bps, implied_vol_px, ref_px):
        self.date = date
        self.time = time
        self.firm = firm
        self.expiration = expiration
        self.option_type = option_type
        self.strike_px = strike_px
        self.strike_spd = strike_spd
        self.bid_px = bid_px
        self.ask_px = ask_px
        self.delta = delta
        self.implied_vol_spd = implied_vol_spd
        self.implied_vol_bps = implied_vol_bps
        self.implied_vol_px = implied_vol_px
        self.ref_px = ref_px


class EmailParser:
    """Set up regular expressions"""
    # use https://pythex.org/ to visualize these if required
    _reg_from = re.compile('From: (.*) At: (.*) (.*) (.*)')

    def __init__(self, line):
        # check whether line has a positive match with all of the regular expressions
        self.matched_from = self._reg_from.match(line)
        self.matched_subject = None
        self.matched_contract = None
        self.matched_quote = None

    def create_company_quote(self):
        if self.matched_from:
            firm = self.matched_from.group(1)
            date = datetime.strptime(
                self.matched_from.group(2), '%m/%d/%y')
            time = datetime.strptime(
                self.matched_from.group(3), '%H:%M:%S')
            return OptionQuote(date, time, firm, None, None, None, None, None, None, None, None, None, None, None)
        else:
            return None


class EmailParserXXX(EmailParser):
    """Set up regular expressions"""
    # use https://pythex.org/ to visualize these if required
    _reg_subject = re.compile('Subject: (.*) - Ref (.*) \((.*)\)')
    _reg_contract = re.compile('Expiry (.*) \((.*) (.*)\)')
    _reg_quote = re.compile(
        '(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)\|( +)([0-9.-]*)/([0-9.-]*)( +)([0-9.-]*)( +)([0-9.-]*)/([0-9.-]*)( +)([0-9.-]*)( +)([0-9.-]*)( +)([0-9.-]*)( +)([0-9.-]*)')

    def __init__(self, line):
        super().__init__(line)
        self.matched_subject = self._reg_subject.match(line)
        self.matched_contract = self._reg_contract.match(line)
        self.matched_quote = self._reg_quote.match(line)

    def update_company_quote(self, company_quote: OptionQuote):
        if self.matched_subject:
            subject = self.matched_subject.group(1)
            ref_px = float(self.matched_subject.group(2))
            company_quote.ref_px = ref_px

    def create_contract_quote(self, company_quote: OptionQuote):
        if self.matched_contract:
            expiry = datetime.strptime(
                self.matched_contract.group(1), '%d%b%y')
            return OptionQuote(company_quote.date, company_quote.time, company_quote.firm, expiry, None, None, None, None, None, None, None, None, None, company_quote.ref_px)
        else:
            return None

    def create_price_quotes(self, contract_quote: OptionQuote):
        if (self.matched_quote):
            strike_px = float(self.matched_quote.group(1))
            strike_spd = float(self.matched_quote.group(3))
            bid_price_put = None if self.matched_quote.group(
                6) == '--' else float(self.matched_quote.group(6))
            ask_price_put = None if self.matched_quote.group(
                7) == '--' else float(self.matched_quote.group(7))
            delta = float(self.matched_quote.group(9))
            bid_price_call = None if self.matched_quote.group(
                11) == '--' else float(self.matched_quote.group(11))
            ask_price_call = None if self.matched_quote.group(
                12) == '--' else float(self.matched_quote.group(12))
            implied_vol_spread = float(self.matched_quote.group(14))
            implied_vol_bps = float(self.matched_quote.group(18))
            put_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Put,
                                           strike_px, strike_spd, bid_price_put, ask_price_put, delta, implied_vol_spread, implied_vol_bps, None, contract_quote.ref_px)
            call_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Call,
                                            strike_px, strike_spd, bid_price_call, ask_price_call, delta, implied_vol_spread, implied_vol_bps, None, contract_quote.ref_px)
            return put_option_quote, call_option_quote
        else:
            return None, None


class EmailParserYYY(EmailParser):
    """Set up regular expressions"""
    # use https://pythex.org/ to visualize these if required
    _reg_subject = re.compile('Subject: (.*) - REF (-*\d*\.?\d+)')
    _reg_contract = re.compile('EXPIRY: (.*) Fwd (.*)/(.*) Dv01 (.*)')
    _reg_quote = re.compile(
        '(-*\d*\.?\d+)( +)\[(-*\d*\.?\d+)\]( *)\|( *)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)%( *)\|( *)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)%( +)\|( +)(-*\d*\.?\d+)%( +)\[( +)(-*\d*\.?\d+)%\]( +)(\+*)(-*\d*\.?\d+)%( +)(-*\d*\.?\d+)')

    def __init__(self, line):
        super().__init__(line)
        self.matched_subject = self._reg_subject.match(line)
        self.matched_contract = self._reg_contract.match(line)
        self.matched_quote = self._reg_quote.match(line)

    def update_company_quote(self, company_quote: OptionQuote):
        if self.matched_subject:
            subject = self.matched_subject.group(1)
            ref_px = float(self.matched_subject.group(2))
            company_quote.ref_px = ref_px

    def create_contract_quote(self, company_quote: OptionQuote):
        if self.matched_contract:
            expiry = datetime.strptime(
                self.matched_contract.group(1), '%d-%b-%Y')
            return OptionQuote(company_quote.date, company_quote.time, company_quote.firm, expiry, None, None, None, None, None, None, None, None, None, company_quote.ref_px)
        else:
            return None

    def create_price_quotes(self, contract_quote: OptionQuote):
        if (self.matched_quote):
            strike_px = float(self.matched_quote.group(1))
            strike_spd = float(self.matched_quote.group(3))
            bid_price_put = None if self.matched_quote.group(
                6) == '--' else float(self.matched_quote.group(6))/100
            ask_price_put = None if self.matched_quote.group(
                8) == '--' else float(self.matched_quote.group(8))/100
            delta_put = float(self.matched_quote.group(10))
            bid_price_call = None if self.matched_quote.group(
                13) == '--' else float(self.matched_quote.group(13))/100
            ask_price_call = None if self.matched_quote.group(
                15) == '--' else float(self.matched_quote.group(15))/100
            delta_call = float(self.matched_quote.group(17))
            implied_vol_spread = float(self.matched_quote.group(23))
            implied_vol_bps = float(self.matched_quote.group(28))
            put_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Put,
                                           strike_px, strike_spd, bid_price_put, ask_price_put, delta_put, implied_vol_spread, implied_vol_bps, None, contract_quote.ref_px)
            call_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Call,
                                            strike_px, strike_spd, bid_price_call, ask_price_call, delta_call, implied_vol_spread, implied_vol_bps, None, contract_quote.ref_px)
            return put_option_quote, call_option_quote
        else:
            return None, None


class EmailParserZZZ(EmailParser):
    """Set up regular expressions"""
    # use https://pythex.org/ to visualize these if required
    _reg_subject = re.compile('Subject: (.*)')
    _reg_contract = re.compile('Exp: (.*) Swaptions Ref: (-*\d*\.?\d+)')

    # when considering the tabs (works on pythex.org but not on my machine)
    #_reg_quote = re.compile('(-*\d*\.?\d+)([^\t])\|([^\t])([^\t])([^\t])(-*\d*\.?\d+)([^\t])\/([^\t])([^\t])(-*\d*\.?\d+)([^\t])([^\t])([^\t])(-*\d*\.?\d+)([^\t])\|([^\t])([^\t])([^\t])(-*\d*\.?\d+)([^\t])\/([^\t])([^\t])(-*\d*\.?\d+)([^\t])([^\t])(-*\d*\.?\d+)([^\t])\|([^\t])([^\t])(-*\d*\.?\d+)([^\t])([^\t])([^\t])(\+?)(-*\d*\.?\d+)([^\t])\|([^\t])([^\t])([^\t])([^\t])(-*\d*\.?\d+)')

    # after replacing tabs with spaces (works on pythex.org but not on my machine)
    #_reg_quote = re.compile('(-*\d*\.?\d+)( +)\|( +)(-*\d*\.?\d+)( +)\/( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)\|( +)(-*\d*\.?\d+)( +)\/( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)\|( +)(-*\d*\.?\d+)( +)(\+?)(-*\d*\.?\d+)( +)\|( +)(-*\d*\.?\d+)')

    # After removing the tabs (works on pythex.org but not on my machine)
    #_reg_quote = re.compile('(-*\d*\.?\d+)\|(-*\d*\.?\d+)\/(-*\d*\.?\d+)\|(-*\d*\.?\d+)\/(-*\d*\.?\d+)(-*\d*\.?\d+)\|(-*\d*\.?\d+)(\+?)(-*\d*\.?\d+)( *)\|(-*\d*\.?\d+)')

    # Works here: https://onecompiler.com/python/3xuwu4wmu but NOT ON MY MACHINE (VSCode 1.64.2 and Python 3.10)!!!
    _reg_quote = re.compile(
        '(-*\d*\.?\d+)(\s+)\|(\s+)(-*\d*\.?\d+)(\s+)\/(\s+)(-*\d*\.?\d+)(\s+)(-*\d*\.?\d+)(\s+)\|(\s+)(-*\d*\.?\d+)(\s+)\/(\s+)(-*\d*\.?\d+)(\s+)(-*\d*\.?\d+)(\s+)\|(\s+)(-*\d*\.?\d+)(\s+)(\+?)(-*\d*\.?\d+)(\s+)\|(\s+)(-*\d*\.?\d+)')

    def __init__(self, line):
        line = re.sub('Â\xa0', '', line)  # remove non-breaking space
        super().__init__(line)
        self.matched_subject = self._reg_subject.match(line)
        self.matched_quote = self._reg_quote.match(line)
        self.matched_contract = self._reg_contract.match(line)

    def update_company_quote(self, company_quote: OptionQuote):
        pass

    def create_contract_quote(self, company_quote: OptionQuote):
        if self.matched_contract:
            expiry = datetime.strptime(
                self.matched_contract.group(1), '%d-%b-%y')
            ref_px = float(self.matched_contract.group(2))
            return OptionQuote(company_quote.date, company_quote.time, company_quote.firm, expiry, None, None, None, None, None, None, None, None, None, ref_px)
        else:
            return None

    def create_price_quotes(self, contract_quote: OptionQuote):
        if self.matched_quote:
            strike_px = float(self.matched_quote.group(1))
            bid_price_put = None if self.matched_quote.group(
                4) == '--' else float(self.matched_quote.group(4))/100
            ask_price_put = None if self.matched_quote.group(
                7) == '--' else float(self.matched_quote.group(7))/100
            delta_put = float(self.matched_quote.group(9))
            bid_price_call = None if self.matched_quote.group(
                12) == '--' else float(self.matched_quote.group(12))/100
            ask_price_call = None if self.matched_quote.group(
                15) == '--' else float(self.matched_quote.group(15))/100
            delta_call = float(self.matched_quote.group(17))
            implied_vol_spread = float(self.matched_quote.group(20))
            implied_vol_px = float(self.matched_quote.group(26))
            put_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Put,
                                           strike_px, None, bid_price_put, ask_price_put, delta_put, implied_vol_spread, None, implied_vol_px, contract_quote.ref_px)
            call_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Call,
                                            strike_px, None, bid_price_call, ask_price_call, delta_call, implied_vol_spread, None, implied_vol_px, contract_quote.ref_px)
            return put_option_quote, call_option_quote
        else:
            return None, None


class EmailParserWWW(EmailParser):
    """Set up regular expressions"""
    # use https://pythex.org/ to visualize these if required
    _reg_subject = re.compile('Subject: (.*)\[ref (-*\d*\.?\d+)\]')
    _reg_contract = re.compile(
        'CDX Options: HY \((.*)\) (.*) \*\* Fwd @(-*\d*\.?\d+), Delta @(-*\d*\.?\d+)')
    _reg_quote_put_only = re.compile(
        '  -  \|     -       -    -    -   - \|( *)(-*\d*\.?\d+)( *)\|( *)(-*\d*\.?\d+)/(-*\d*\.?\d+)( +)(-*\d*\.?\d+)%( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)')
    _reg_quote_call_put = re.compile(
        '(-*\d*\.?\d+)( *)\|( *)(-*\d*\.?\d+)/(-*\d*\.?\d+)( +)(-*\d*\.?\d+)%( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)\|( +)(-*\d*\.?\d+)( *)\|( *)(-*\d*\.?\d+)/(-*\d*\.?\d+)( +)(-*\d*\.?\d+)%( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)( +)(-*\d*\.?\d+)')

    def __init__(self, line):
        super().__init__(line)
        self.matched_subject = self._reg_subject.match(line)
        self.matched_contract = self._reg_contract.match(line)
        self.matched_quote_put_only = self._reg_quote_put_only.match(line)
        self.matched_quote_call_put = self._reg_quote_call_put.match(line)
        self.matched_quote = self.matched_quote_put_only or self.matched_quote_call_put

    def update_company_quote(self, company_quote: OptionQuote):
        if self.matched_subject:
            subject = self.matched_subject.group(1)
            ref_px = float(self.matched_subject.group(2))
            company_quote.ref_px = ref_px

    def create_contract_quote(self, company_quote: OptionQuote):
        if self.matched_contract:
            expiry = datetime.strptime(
                self.matched_contract.group(2), '%d-%b-%y')
            return OptionQuote(company_quote.date, company_quote.time, company_quote.firm, expiry, None, None, None, None, None, None, None, None, None, company_quote.ref_px)
        else:
            return None

    def create_price_quotes(self, contract_quote: OptionQuote):
        if (self.matched_quote_call_put):
            strike_px_call = float(self.matched_quote.group(1))
            bid_price_call = None if self.matched_quote.group(
                4) == '--' else float(self.matched_quote.group(4))/100
            ask_price_call = None if self.matched_quote.group(
                5) == '--' else float(self.matched_quote.group(5))/100
            delta_call = float(self.matched_quote.group(7))
            implied_vol_spread_call = float(self.matched_quote.group(9))
            implied_vol_bps_call = float(self.matched_quote.group(13))

            strike_px_put = float(self.matched_quote.group(15))
            bid_price_put = None if self.matched_quote.group(
                18) == '--' else float(self.matched_quote.group(18))/100
            ask_price_put = None if self.matched_quote.group(
                19) == '--' else float(self.matched_quote.group(19))/100
            delta_put = float(self.matched_quote.group(21))
            implied_vol_spread_put = float(self.matched_quote.group(23))
            implied_vol_bps_put = float(self.matched_quote.group(27))
            put_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Put,
                                           strike_px_put, None, bid_price_put, ask_price_put, delta_put, implied_vol_spread_put, implied_vol_bps_put, None, contract_quote.ref_px)
            call_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Call,
                                            strike_px_call, None, bid_price_call, ask_price_call, delta_call, implied_vol_spread_call, implied_vol_bps_call, None, contract_quote.ref_px)
            return put_option_quote, call_option_quote
        elif (self.matched_quote_put_only):
            strike_px_put = float(self.matched_quote.group(2))
            bid_price_put = None if self.matched_quote.group(
                5) == '--' else float(self.matched_quote.group(5))/100
            ask_price_put = None if self.matched_quote.group(
                6) == '--' else float(self.matched_quote.group(6))/100
            delta_put = float(self.matched_quote.group(8))
            implied_vol_spread_put = float(self.matched_quote.group(10))
            implied_vol_bps_put = float(self.matched_quote.group(14))
            put_option_quote = OptionQuote(contract_quote.date, contract_quote.time, contract_quote.firm, contract_quote.expiration, OptionType.Put,
                                           strike_px_put, None, bid_price_put, ask_price_put, delta_put, implied_vol_spread_put, implied_vol_bps_put, None, contract_quote.ref_px)
            call_option_quote = copy.deepcopy(contract_quote)
            return put_option_quote, call_option_quote
        else:
            return None, None


def process_option_quote(line, parent_quotes, price_quotes):
    parser = None
    if line == '\n':
        return
    if(parent_quotes[0] is None):
        parser = EmailParser(line)
    elif (parent_quotes[0].firm == 'XXX'):
        parser = EmailParserXXX(line)
    elif (parent_quotes[0].firm == 'YYY'):
        parser = EmailParserYYY(line)
    elif (parent_quotes[0].firm == 'ZZZ'):
        parser = EmailParserZZZ(line)
    elif (parent_quotes[0].firm == 'WWW'):
        parser = EmailParserWWW(line)

    if parser.matched_from:
        parent_quotes[0] = parser.create_company_quote()
    elif parser.matched_subject:
        parser.update_company_quote(parent_quotes[0])
    elif parser.matched_contract:
        parent_quotes[1] = parser.create_contract_quote(parent_quotes[0])
    elif parser.matched_quote:
        put_price_quote, call_price_quote = parser.create_price_quotes(
            parent_quotes[1])
        price_quotes.append(put_price_quote)
        price_quotes.append(call_price_quote)


def convert_to_dataframe(price_quotes):
    data = {'Date': [], 'Time': [], 'Firm': [], 'Expiration': [], 'Option Type': [], 'Strike Px': [], 'Bid Price': [
    ], 'Ask Price': [], 'Delta': [], 'Implied Vol Spd': [], 'Implied Vol Bps': [], 'Implied Vol Px': [], 'Ref Px': []}
    for quote in price_quotes:
        data['Date'].append(quote.date.strftime('%d-%b-%y'))
        data['Time'].append(quote.time.strftime('%H:%M:%S'))
        data['Firm'].append(quote.firm)
        data['Expiration'].append(quote.expiration.strftime('%d-%b-%y'))
        data['Option Type'].append(
            'P' if quote.option_type == OptionType.Put else 'C')
        data['Strike Px'].append(quote.strike_px)
        data['Bid Price'].append(quote.bid_px)
        data['Ask Price'].append(quote.ask_px)
        data['Delta'].append(quote.delta)
        data['Implied Vol Spd'].append(quote.implied_vol_spd)
        data['Implied Vol Bps'].append(quote.implied_vol_bps)
        data['Implied Vol Px'].append(quote.implied_vol_px)
        data['Ref Px'].append(quote.ref_px)
    df = pd.DataFrame(data)
    return df


if __name__ == '__main__':
    directory = os.getcwd()
    price_quotes = []
    print('The current direcotry is: ' + directory)
    for file in os.listdir(directory):
        if file.startswith('hycdx_option_quotes_') and file.endswith('.txt'):
            print('Reading file: ' + file)
            with open(file, 'r') as file:
                line = file.readline()
                company_quote = None
                contract_quote = None
                parent_quotes = [company_quote, contract_quote]
                while line:
                    process_option_quote(line, parent_quotes, price_quotes)
                    line = file.readline()
    if len(price_quotes) > 0:
        df = convert_to_dataframe(price_quotes)  # pip install pandas
        df.to_excel('task1_output_actual.xlsx',
                    index=False)  # pip install openpyxl
        print('Successfully converted to excel file.')
