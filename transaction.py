import datetime
import helperMethods as Helper


class Transaction:
    def __init__(self, row, transactionDate, type,
                 quantity, symbol, month, day, year, strike,
                 optionType, premium, trade=None, leg=None,
                 transactionWords=None, openPremium=None, closeDate=None):
        self.row = int(row)
        try:
            self.transactionDate = datetime.datetime.strptime(
                str(transactionDate), '%Y-%m-%d %H:%M:%S').date()
        except:
            self.transactionDate = datetime.datetime.strptime(
                str(transactionDate), '%m/%d/%Y').date()
        self.type = str(type)
        self.quantity = int(quantity)
        self.symbol = str(symbol)
        self.month = str(month)
        self.day = int(day)
        self.year = int(year)
        self.strike = float(strike)
        self.optionType = str(optionType)
        self.premium = float(premium)
        if trade:
            self.trade = int(trade)
        else:
            self.trade = None
        if leg:
            self.leg = int(leg)
        else:
            self.leg = None
        if transactionWords:
            self.transactionWords = str(transactionWords)
        else:
            self.transactionWords = None
        if openPremium:
            self.openPremium = float(openPremium)
        else:
            self.openPremium = None
        if self.optionType.lower() == 'put':
            self.putType = 'SP'
        else:
            self.putType = 'LC'
        if closeDate:
            try:
                self.closeDate = datetime.datetime.strptime(
                    str(closeDate), '%Y-%m-%d %H:%M:%S').date()
            except:
                self.closeDate = datetime.datetime.strptime(
                    str(closeDate), '%m/%d/%Y').date()
        else:
            self.closeDate = None
        expirationDate = str(Helper.getMonthNumber(month)
                             ) + "/" + str(day) + "/" + str(year)
        self.expirationDate = datetime.datetime.strptime(
            str(expirationDate), '%m/%d/%Y').date()

    def equals(self, other):
        return self.lightEquals(other) \
            and self.type == other.type \
            and self.transactionDate == other.transactionDate \
            and self.premium == other.premium

    def closingEquals(self, other):
        return self.lightEquals(other) \
            and self.trade == other.trade

    def lightEquals(self, other):
        return self.symbol == other.symbol \
            and self.month == other.month \
            and self.day == other.day \
            and self.year == other.year \
            and self.strike == other.strike \
            and self.optionType == other.optionType
