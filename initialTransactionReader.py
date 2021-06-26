import datetime
import helperMethods as Helper
from transaction import Transaction

TRANSACTIONSINFOSHEETNAME = 'transactions'
YELLOBARINOVERALLDATASHEETROW = 7
OVERALLDATASHEETNAME = 'DATA'


def readOldTransactions(portfolioWorkbook):
    processingSheet = portfolioWorkbook[OVERALLDATASHEETNAME]
    sheetRows = processingSheet.rows
    return turnProcessingRowsToTransactions(sheetRows)


def addOldTransactionToArray(row, array):
    splitTransactionInfo = row[0].value.split(' ')
    if len(splitTransactionInfo) == 10 and boughtOrSold(splitTransactionInfo[0]):
        transaction = Transaction(row=row[0].row, transactionDate=row[4].value, type=splitTransactionInfo[0],
                                  quantity=splitTransactionInfo[1], symbol=splitTransactionInfo[2], month=splitTransactionInfo[3],
                                  day=splitTransactionInfo[4], year=splitTransactionInfo[5], strike=splitTransactionInfo[6],
                                  optionType=splitTransactionInfo[7], premium=splitTransactionInfo[9], trade=row[1].value,
                                  leg=row[2].value, transactionWords=row[0].value, openPremium=row[10].value,
                                  closeDate=row[5].value)
    if len(splitTransactionInfo) == 12 and boughtOrSold(splitTransactionInfo[0]):
        transaction = Transaction(row=row[0].row, transactionDate=row[4].value, type=splitTransactionInfo[0],
                                  quantity=splitTransactionInfo[3], symbol=splitTransactionInfo[4], month=splitTransactionInfo[5],
                                  day=splitTransactionInfo[6], year=splitTransactionInfo[7], strike=splitTransactionInfo[8],
                                  optionType=splitTransactionInfo[9], premium=splitTransactionInfo[11], trade=row[1].value,
                                  leg=row[2].value, transactionWords=row[0].value, openPremium=row[10].value,
                                  closeDate=row[5].value)
    array.append(transaction)


def turnProcessingRowsToTransactions(rows):
    alreadyReadTransactions = []
    openTransactions = []
    maximumTransactionNumber = 0
    for count, row in enumerate(rows):
        if count > YELLOBARINOVERALLDATASHEETROW and row[0].value != None:
            addOldTransactionToArray(row, alreadyReadTransactions)
            closeDate = row[5].value
            if closeDate == None:
                addOldTransactionToArray(row, openTransactions)
            tradeNumber = row[1].value
            if tradeNumber > maximumTransactionNumber:
                maximumTransactionNumber = tradeNumber
    return alreadyReadTransactions, openTransactions, maximumTransactionNumber


def getNewTransactions(transactionWorkbook, alreadyReadTransactions):
    return turnTransactionRowsToTransactions(transactionWorkbook, alreadyReadTransactions)


def turnTransactionRowsToTransactions(rows, alreadyReadTransactions):
    newTransactionArray = []
    for count, row in enumerate(rows):
        if len(row) > 1:
            try:
                splitTransactionInfo = row[2].split(' ')
                if len(splitTransactionInfo) == 10 and boughtOrSold(splitTransactionInfo[0]):
                    transaction = Transaction(row=count, transactionDate=row[0], type=splitTransactionInfo[0],
                                              quantity=splitTransactionInfo[1], symbol=splitTransactionInfo[
                        2], month=splitTransactionInfo[3],
                        day=splitTransactionInfo[4], year=splitTransactionInfo[
                                                  5], strike=splitTransactionInfo[6],
                        optionType=splitTransactionInfo[7], premium=splitTransactionInfo[9], trade=None, leg=None,
                        transactionWords=row[2], openPremium=row[7])
                    checkForExistingTransactions(
                        transaction, alreadyReadTransactions, newTransactionArray)
                if len(splitTransactionInfo) == 12 and boughtOrSold(splitTransactionInfo[0]):
                    transaction = Transaction(row=count, transactionDate=row[0], type=splitTransactionInfo[0],
                                              quantity=splitTransactionInfo[3], symbol=splitTransactionInfo[
                        4], month=splitTransactionInfo[5],
                        day=splitTransactionInfo[6], year=splitTransactionInfo[
                                                  7], strike=splitTransactionInfo[8],
                        optionType=splitTransactionInfo[9], premium=splitTransactionInfo[11], trade=None, leg=None,
                        transactionWords=row[2], openPremium=row[7])
                    checkForExistingTransactions(
                        transaction, alreadyReadTransactions, newTransactionArray)
                if len(splitTransactionInfo) == 7 and splitTransactionInfo[0] == 'REMOVAL':
                    symbol = getSymbolFromRemoval(splitTransactionInfo)
                    transaction = Transaction(row=count, transactionDate=row[0], type='REMOVAL',
                                              quantity=1, symbol=symbol, month='Sep',
                                              day=16, year=2020, strike=1,
                                              optionType='optionType', premium=1, transactionWords=row[2])
                    newTransactionArray.append(transaction)
            except Exception as e:
                print("WARNING FOR ROW NUMBER WITH INVALID INPUT " + str(count))
                print(e)
                pass
    return newTransactionArray


def getSymbolFromRemoval(splitTransactionInfo):
    symbolInfo = splitTransactionInfo[6]
    symbolInfoSplit = symbolInfo.split('.')
    symbolLeftSide = symbolInfoSplit[0]
    return symbolLeftSide.split('0')[1]


def checkForExistingTransactions(transaction, alreadyReadTransactions, newTransactionArray):
    if findExistingTransactionInArray(transaction, alreadyReadTransactions) == None:
        duplicateNewTransaction = findExistingTransactionInArray(
            transaction, newTransactionArray)
        if duplicateNewTransaction:
            updateDuplicateTransactionWithCombinedValues(
                duplicateNewTransaction, transaction)
        else:
            newTransactionArray.append(transaction)


def updateDuplicateTransactionWithCombinedValues(duplicateTransaction, transaction):
    duplicateTransaction.quantity += transaction.quantity
    duplicateTransaction.openPremium += transaction.openPremium
    transactionWords = duplicateTransaction.transactionWords
    wordSplit = transactionWords.split(" ")
    if len(wordSplit) == 12:
        wordSplit[3] = str(int(wordSplit[3]) + transaction.quantity)
        separator = " "
        duplicateTransaction.transactionWords = separator.join(wordSplit)
    elif len(wordSplit) == 10:
        wordSplit[1] = str(int(wordSplit[1]) + transaction.quantity)
        separator = " "
        duplicateTransaction.transactionWords = separator.join(wordSplit)


def findExistingTransactionInArray(transaction, array):
    for item in array:
        if transaction.equals(item):
            return item
    return None


def boughtOrSold(startOfInfo):
    if startOfInfo.lower() == 'bought' or startOfInfo.lower() == 'sold':
        return True
    return False


def openWorksbooks():
    print("Select the transactions workbook")
    transactionWorkbook = Helper.openTransactionWorkbookByDialog()
    print("Select the result portfolio workbook")
    portfolioWorkbook = Helper.openPortfolioWorkbookByDialog()
    return transactionWorkbook, portfolioWorkbook
