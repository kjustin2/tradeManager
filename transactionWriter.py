import helperMethods as Helper
from os import close


def writeTransactions(newTransactions, openTransactions, alreadyReadTransactions,
                      portfolioWorkbook, sheetName, maximumTransactionNumber):
    writeToLog = open('log.txt', 'w')
    sheet = portfolioWorkbook[sheetName]
    lastCurrentRow = alreadyReadTransactions[-1].row
    currentWriteRow = lastCurrentRow + 1
    for transaction in newTransactions:
        if transaction.type == 'REMOVAL':
            writeRemovalTransaction(
                transaction, openTransactions, sheet, writeToLog, alreadyReadTransactions, maximumTransactionNumber)
        else:
            maximumTransactionNumber = writeNewTransactionGetNewMaxTransaction(transaction, openTransactions,
                                                                               sheet, currentWriteRow, maximumTransactionNumber, writeToLog, alreadyReadTransactions)
            currentWriteRow += 1
    writeToLog.close()


def writeRemovalTransaction(transaction, openTransactions, sheet, writeToLog, alreadyReadTransactions, maximumTransactionNumber):
    matchingOpenTransactions = []
    for openTransaction in openTransactions:
        if transaction.symbol == openTransaction.symbol:
            matchingOpenTransactions.append(openTransaction)
    if len(matchingOpenTransactions) == 1:
        transaction.transactionDate = getRemovalTransactionDate(
            matchingOpenTransactions)
        updateClosedOpenTransactions(transaction, matchingOpenTransactions,
                                     openTransactions, writeToLog, sheet, alreadyReadTransactions)
    elif len(matchingOpenTransactions) == 0:
        return
    else:
        closeableTransactions = getCloseableTransactionsForRemoval(
            matchingOpenTransactions, transaction, maximumTransactionNumber)
        transaction.transactionDate = getRemovalTransactionDate(
            closeableTransactions)
        updateClosedOpenTransactions(transaction, closeableTransactions,
                                     openTransactions, writeToLog, sheet, alreadyReadTransactions)


def getRemovalTransactionDate(closeableTransactions):
    return closeableTransactions[0].expirationDate


def getCloseableTransactionsForRemoval(matchingOpenTransactions, newTransaction, maximumTransactionNumber):
    tradeNumber = getRemovalTradeNumber(
        matchingOpenTransactions, newTransaction, maximumTransactionNumber)
    closeableTransactions = []
    for transaction in matchingOpenTransactions:
        if transaction.trade == tradeNumber:
            closeableTransactions.append(transaction)
    return closeableTransactions


def getRemovalTradeNumber(matchingOpenTransactions, newTransaction, maximumTransactionNumber):
    print("REMOVAL MATCHING OPEN TRADES ---------")
    for transaction in matchingOpenTransactions:
        printTransactionNumberHelpfulSummary(transaction)
    print("END ---------")
    print("YOUR CURRENT REMOVAL ---------")
    print("REMOVAL SYMBOL " + newTransaction.symbol)
    print("REMOVAL WORDS " + newTransaction.transactionWords)
    print("END ---------")
    tryInput = True
    while tryInput:
        try:
            transactionNumber = int(input("Enter the correct trade number: "))
            if transactionNumber > maximumTransactionNumber:
                raise Exception()
            tryInput = False
        except:
            print("Invalid input for transaction number")
    print('\n\n')
    return transactionNumber


def writeNewTransactionGetNewMaxTransaction(newTransaction, openTransactions,
                                            currentSheet, currentRow, maximumTransactionNumber, writer, alreadyReadTransactions):
    newTransaction.row = int(currentRow)
    performSimpleCopy(newTransaction, currentSheet, currentRow, writer)
    newTransaction.trade, newTransaction.leg = findTransactionNumberAndLatestLeg(newTransaction, openTransactions,
                                                                                 currentSheet, currentRow, maximumTransactionNumber)
    Helper.writeToFile(writer, 'newTransactionTrade', currentSheet,
                       20, currentRow, newTransaction.trade)
    Helper.writeToFile(writer, 'newTransactionLeg', currentSheet,
                       21, currentRow, newTransaction.leg)
    closeableTransactions = getCloseableTransactions(
        newTransaction, openTransactions)
    if closeableTransactions == None:
        openTransactions.append(newTransaction)
    else:
        updateClosedOpenTransactions(newTransaction, closeableTransactions,
                                     openTransactions, writer, currentSheet, alreadyReadTransactions)
        Helper.writeToFile(writer, 'closedDate', currentSheet,
                           22, newTransaction.row, newTransaction.transactionDate, dateFormat=True)
    if newTransaction.trade > maximumTransactionNumber:
        maximumTransactionNumber = newTransaction.trade
    return maximumTransactionNumber


def getCloseableTransactions(newTransaction, openTransactions):
    closeableTransactions = []
    overallQuantity = 0
    for transaction in openTransactions:
        if transaction.closingEquals(newTransaction):
            overallQuantity = updateOverallQuantity(
                overallQuantity, transaction)
            closeableTransactions.append(transaction)
    if newTransaction.quantity == overallQuantity and newTransaction.type == 'Bought':
        return closeableTransactions
    if newTransaction.quantity == -1 * overallQuantity and newTransaction.type == 'Sold':
        return closeableTransactions
    return None


def updateOverallQuantity(overallQuantity, transaction):
    if transaction.type == 'Bought':
        overallQuantity -= transaction.quantity
    else:
        overallQuantity += transaction.quantity
    return overallQuantity


def updateClosedOpenTransactions(newTransaction, closeableTransactions, openTransactions, writer, currentSheet, alreadyReadTransactions):
    for transaction in closeableTransactions:
        Helper.writeToFile(writer, 'closedDate', currentSheet,
                           22, transaction.row, newTransaction.transactionDate, dateFormat=True)
        transaction.closeDate = newTransaction.transactionDate
        closeAlreadyReadTransaction(
            transaction, alreadyReadTransactions, newTransaction.transactionDate)
    newTransaction.closeDate = newTransaction.transactionDate
    openTransactions = getNewOpenTransactionsAfterClosing(
        closeableTransactions, openTransactions)


def closeAlreadyReadTransaction(compareTransaction, alreadyReadTransactions, transactionCloseDate):
    for transaction in alreadyReadTransactions:
        if transaction.equals(compareTransaction):
            transaction.closeDate = transactionCloseDate


def getNewOpenTransactionsAfterClosing(closeableTransactions, openTransactions):
    newOpenTransactions = []
    for transaction in openTransactions:
        if transaction not in closeableTransactions:
            newOpenTransactions.append(transaction)
    return newOpenTransactions


def findTransactionNumberAndLatestLeg(newTransaction, openTransactions, currentSheet, currentRow, maximumTransactionNumber):
    transactionNumber = -1
    transactionLeg = 0
    matchingOpenTransactions = getMatchingOpenTransanctions(
        newTransaction, openTransactions)
    if len(matchingOpenTransactions) > 0:
        transactionNumber = getNewTransactionNumberAfterMatching(
            matchingOpenTransactions, newTransaction)
        if transactionNumber <= maximumTransactionNumber:
            transactionLeg = getTransactionLegByTransactionNumber(
                matchingOpenTransactions)
    else:
        transactionNumber = maximumTransactionNumber + 1
    return int(transactionNumber), int(transactionLeg + 1)


def getNewTransactionNumberAfterMatching(matchingOpenTransactions, newTransaction):
    print("MATCHING OPEN TRADES ---------")
    for transaction in matchingOpenTransactions:
        printTransactionNumberHelpfulSummary(transaction)
    print("END ---------")
    print("YOUR CURRENT TRADE ---------")
    printTransactionNumberHelpfulSummary(newTransaction)
    print("END ---------")
    tryInput = True
    while tryInput:
        try:
            transactionNumber = int(input("Enter the correct trade number: "))
            tryInput = False
        except:
            print("Invalid input for transaction number")
    print('\n\n')
    return transactionNumber


def printTransactionNumberHelpfulSummary(transaction):
    if transaction.leg:
        print("TRADE LEG " + str(transaction.leg))
    else:
        print("NO TRADE LEG YET")
    if transaction.trade:
        print("TRADE NUMBER " + str(transaction.trade))
    else:
        print("NO TRADE NUMBER YET")
    print("TRANSACTION DATE " +
          transaction.transactionDate.strftime("%m/%d/%Y"))
    print("TRANSACTION WORDS " + transaction.transactionWords)


def getTransactionLegByTransactionNumber(matchingOpenTransactions):
    maxLegNumber = -1
    for transaction in matchingOpenTransactions:
        if transaction.leg > maxLegNumber:
            maxLegNumber = transaction.leg
    return maxLegNumber


def getMatchingOpenTransanctions(newTransaction, openTransactions):
    matchingTransactions = []
    for transaction in openTransactions:
        if transaction.symbol == newTransaction.symbol:
            matchingTransactions.append(transaction)
    return matchingTransactions


def performSimpleCopy(newTransaction, currentSheet, currentRow, writer):
    Helper.writeToFile(writer, 'newTransactionTransactionDate', currentSheet,
                       1, currentRow, newTransaction.transactionDate, dateFormat=True)
    Helper.writeToFile(writer, 'newTransactionId', currentSheet,
                       2, currentRow, newTransaction.id)
    Helper.writeToFile(writer, 'newTransactionType', currentSheet,
                       3, currentRow, newTransaction.type)
    Helper.writeToFile(writer, 'newTransactionQuantity', currentSheet,
                       4, currentRow, newTransaction.quantity)
    Helper.writeToFile(writer, 'newTransactionSymbol', currentSheet,
                       5, currentRow, newTransaction.symbol)
    Helper.writeToFile(writer, 'newTransactionMonth', currentSheet,
                       6, currentRow, newTransaction.month)
    Helper.writeToFile(writer, 'newTransactionDay', currentSheet,
                       7, currentRow, newTransaction.day)
    Helper.writeToFile(writer, 'newTransactionYear', currentSheet,
                       8, currentRow, newTransaction.year)
    Helper.writeToFile(writer, 'newTransactionStrike', currentSheet,
                       9, currentRow, newTransaction.strike)
    Helper.writeToFile(writer, 'newTransactionOptionType', currentSheet,
                       10, currentRow, newTransaction.optionType)
    Helper.writeToFile(writer, '@', currentSheet,
                       11, currentRow, '@')
    Helper.writeToFile(writer, 'newTransactionPremium', currentSheet,
                       12, currentRow, newTransaction.premium)
    Helper.writeToFile(writer, 'newTransactionQuantity', currentSheet,
                       14, currentRow, newTransaction.quantity)
    Helper.writeToFile(writer, 'newTransactionTransactionWords', currentSheet,
                       15, currentRow, newTransaction.transactionWords)
    Helper.writeToFile(writer, 'newTransactionPremium', currentSheet,
                       16, currentRow, newTransaction.premium)
    Helper.writeToFile(writer, 'newTransactionCommission', currentSheet,
                       17, currentRow, newTransaction.commission)
    Helper.writeToFile(writer, 'newTransactionOpenPremium', currentSheet,
                       18, currentRow, newTransaction.openPremium)
    Helper.writeToFile(writer, 'newTransactionFee', currentSheet,
                       19, currentRow, newTransaction.fee)
    Helper.writeToFile(writer, 'newTransactionMonthNumber', currentSheet,
                       23, currentRow, Helper.getMonthNumber(newTransaction.month))
    Helper.writeToFile(writer, 'transactionOptionTypePutType', currentSheet,
                       24, currentRow, newTransaction.putType)
    Helper.writeToFile(writer, 'newTransactionSymbolFormula', currentSheet,
                       25, currentRow, Helper.getSymbolFormula(newTransaction.symbol))
    Helper.writeToFile(writer, 'newTransactionDateFormula', currentSheet,
                       26, currentRow, Helper.getDateFormula(newTransaction))
