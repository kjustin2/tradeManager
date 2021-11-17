from os import write
import helperMethods as Helper
import initialTransactionReader as InitialTransactionReader
import transactionProcessor as TransactionProcessor
import dataSheetWriter as DataSheetWriter
import sys


def checkExceptionMessaging(exceptionMessage):
    exceptionMessage = str(exceptionMessage)
    if "[Errno 13] Permission denied" in exceptionMessage:
        Helper.printWithNewLines("CLOSE YOUR PORTFOLIO FILE")
    elif "charmap' codec can't decode byte" in exceptionMessage:
        Helper.printWithNewLines("ENSURE TRANSACTIONS FILE IS A CSV FILE")
    else:
        print(exceptionMessage)


def getTransactions(transactionWorkbook, portfolioWorkbook):
    alreadyReadTransactions, openTransactions, maximumTradeNumber = InitialTransactionReader.readOldTransactions(
        portfolioWorkbook)
    newTransactions = InitialTransactionReader.getNewTransactions(
        transactionWorkbook, alreadyReadTransactions)
    writeableNewTransactions = Helper.getWriteableTransactions(newTransactions)
    allWriteableTransactions = alreadyReadTransactions + writeableNewTransactions
    return alreadyReadTransactions, openTransactions, newTransactions, maximumTradeNumber, allWriteableTransactions


def checkDebug():
    if len(sys.argv) > 1 and sys.argv[1] == 'DEBUG':
        Helper.DEBUG = True


def main():
    checkDebug()

    try:
        transactionWorkbook, portfolioWorkbook = InitialTransactionReader.openWorksbooks()
        alreadyReadTransactions, openTransactions, newTransactions, maximumTradeNumber, allWriteableTransactions = getTransactions(
            transactionWorkbook, portfolioWorkbook)

        TransactionProcessor.updateTransactions(
            newTransactions, openTransactions, alreadyReadTransactions, maximumTradeNumber, InitialTransactionReader.YELLOBARINOVERALLDATASHEETROW,
            allWriteableTransactions)

        DataSheetWriter.clearEditableData(allWriteableTransactions, portfolioWorkbook[InitialTransactionReader.OVERALLDATASHEETNAME], InitialTransactionReader.YELLOBARINOVERALLDATASHEETROW)
        DataSheetWriter.writeTransactions(
            allWriteableTransactions, portfolioWorkbook[InitialTransactionReader.OVERALLDATASHEETNAME], InitialTransactionReader.YELLOBARINOVERALLDATASHEETROW)

        portfolioWorkbookSavedName = input(
            "Enter the name to save your portfolio as: ")
        portfolioWorkbook.save(portfolioWorkbookSavedName + ".xlsx")

        print("SUCCESS")
        input()
    except Exception as e:
        checkExceptionMessaging(e)
        print('-------------------')
        print("FAILED - Press Enter to Complete")
        input()


if __name__ == '__main__':
    main()
