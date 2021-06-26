from os import write
import helperMethods as Helper
import initialTransactionReader as InitialTransactionReader
import transactionProcessor as TransactionProcessor
import dataSheetWriter as DataSheetWriter


def checkExceptionMessaging(exceptionMessage):
    exceptionMessage = str(exceptionMessage)
    if "[Errno 13] Permission denied" in exceptionMessage:
        Helper.printWithNewLines("CLOSE YOUR PORTFOLIO FILE")
    elif "charmap' codec can't decode byte" in exceptionMessage:
        Helper.printWithNewLines("ENSURE TRANSACTIONS FILE IS A CSV FILE")
    else:
        print(exceptionMessage)


def getTransactions(transactionWorkbook, portfolioWorkbook):
    alreadyReadTransactions, openTransactions, maximumTransactionNumber = InitialTransactionReader.readOldTransactions(
        portfolioWorkbook)
    newTransactions = InitialTransactionReader.getNewTransactions(
        transactionWorkbook, alreadyReadTransactions)
    writeableNewTransactions = Helper.getWriteableTransactions(newTransactions)
    allWriteableTransactions = alreadyReadTransactions + writeableNewTransactions
    return alreadyReadTransactions, openTransactions, newTransactions, maximumTransactionNumber, allWriteableTransactions


def main():
    try:
        transactionWorkbook, portfolioWorkbook = InitialTransactionReader.openWorksbooks()
        alreadyReadTransactions, openTransactions, newTransactions, maximumTransactionNumber, allWriteableTransactions = getTransactions(
            transactionWorkbook, portfolioWorkbook)

        TransactionProcessor.updateTransactions(
            newTransactions, openTransactions, alreadyReadTransactions, maximumTransactionNumber)
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
