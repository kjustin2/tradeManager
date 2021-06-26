import helperMethods as Helper
from os import close


def updateTransactions(newTransactions, openTransactions, alreadyReadTransactions,  maximumTransactionNumber, yellowBarRow):
    lastCurrentRow = yellowBarRow + 1
    if len(alreadyReadTransactions) > 0:
        lastCurrentRow = alreadyReadTransactions[-1].row
    currentUpdateRow = lastCurrentRow + 1
    for transaction in newTransactions:
        if transaction.type == 'REMOVAL':
            updateRemovalTransaction(
                transaction, openTransactions, alreadyReadTransactions, maximumTransactionNumber)
        else:
            maximumTransactionNumber = updateNewTransactionGetNewMaxTransaction(transaction, openTransactions,
                                                                                currentUpdateRow, maximumTransactionNumber, alreadyReadTransactions)
            currentUpdateRow += 1


def updateRemovalTransaction(transaction, openTransactions, alreadyReadTransactions, maximumTransactionNumber):
    matchingOpenTransactions = []
    for openTransaction in openTransactions:
        if transaction.symbol == openTransaction.symbol:
            matchingOpenTransactions.append(openTransaction)
    if len(matchingOpenTransactions) == 1 or transactionsSameTradeNumber(matchingOpenTransactions):
        transaction.transactionDate = matchingOpenTransactions[0].expirationDate
        updateClosedOpenTransactions(transaction, matchingOpenTransactions,
                                     openTransactions, alreadyReadTransactions)
    elif len(matchingOpenTransactions) == 0:
        return
    else:
        closeableTransactions = getCloseableTransactionsForRemoval(
            matchingOpenTransactions, transaction, maximumTransactionNumber)
        transaction.transactionDate = closeableTransactions[0].expirationDate
        updateClosedOpenTransactions(transaction, closeableTransactions,
                                     openTransactions, alreadyReadTransactions)


def transactionsSameTradeNumber(transactions):
    tradeNumber = transactions[0].trade
    for transaction in transactions:
        if transaction.trade != tradeNumber:
            return False
    return True


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


def updateNewTransactionGetNewMaxTransaction(newTransaction, openTransactions,
                                             currentRow, maximumTransactionNumber, alreadyReadTransactions):
    newTransaction.row = int(currentRow)
    newTransaction.trade, newTransaction.leg = findTransactionNumberAndLatestLeg(newTransaction, openTransactions,
                                                                                 maximumTransactionNumber)
    closeableTransactions = getCloseableTransactions(
        newTransaction, openTransactions)
    if closeableTransactions == None:
        openTransactions.append(newTransaction)
    else:
        updateClosedOpenTransactions(newTransaction, closeableTransactions,
                                     openTransactions, alreadyReadTransactions)
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


def updateClosedOpenTransactions(newTransaction, closeableTransactions, openTransactions, alreadyReadTransactions):
    for transaction in closeableTransactions:
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


def findTransactionNumberAndLatestLeg(newTransaction, openTransactions, maximumTransactionNumber):
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
