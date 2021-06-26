import helperMethods as Helper
from os import close
from openpyxl.worksheet.table import Table, TableStyleInfo


def writeTransactions(allWriteableTransactions, dataSheet, yellowBarRow):
    writeToLog = open('log.txt', 'w')
    allWriteableTransactions.sort(key=lambda x: (x.trade, x.leg))
    currentRow = yellowBarRow + 2
    startRow = yellowBarRow + 1
    endRow = currentRow + len(allWriteableTransactions) - 1
    addTable(dataSheet, startRow, endRow)
    for transaction in allWriteableTransactions:
        writeTransactionToData(
            transaction, dataSheet, currentRow, writeToLog)
        writeFormulaInformation(transaction, dataSheet, currentRow, writeToLog)
        currentRow += 1
    writeToLog.close()


def writeTransactionToData(transaction, dataSheet, currentRow, writeToLog):
    Helper.writeToFile(writeToLog, 'transactionWords', dataSheet,
                       1, currentRow, transaction.transactionWords)
    Helper.writeToFile(writeToLog, 'transactionTrade', dataSheet,
                       2, currentRow, transaction.trade)
    Helper.writeToFile(writeToLog, 'transactionLeg', dataSheet,
                       3, currentRow, transaction.leg)
    Helper.writeToFile(writeToLog, 'transactionPutType', dataSheet,
                       4, currentRow, transaction.putType)
    Helper.writeToFile(writeToLog, 'transactionOpenDate', dataSheet,
                       5, currentRow, transaction.transactionDate, dateFormat=True)
    Helper.writeToFile(writeToLog, 'transactionCloseDate', dataSheet,
                       6, currentRow, transaction.closeDate, dateFormat=True)
    Helper.writeToFile(writeToLog, 'transactionExpirationDate', dataSheet,
                       7, currentRow, transaction.expirationDate, dateFormat=True)
    Helper.writeToFile(writeToLog, 'transactionStrike', dataSheet,
                       8, currentRow, transaction.strike)
    Helper.writeToFile(writeToLog, 'transactionContracts', dataSheet,
                       10, currentRow, transaction.quantity)
    Helper.writeToFile(writeToLog, 'transactionOpenPremium', dataSheet,
                       11, currentRow, transaction.openPremium)
    Helper.writeToFile(writeToLog, 'transactionSymbol', dataSheet,
                       13, currentRow, transaction.symbol)


def writeFormulaInformation(transaction, dataSheet, currentRow, writeToLog):
    Helper.writeToFile(writeToLog, 'currentValue', dataSheet,
                       14, currentRow, Helper.getSymbolFormula(transaction.symbol))
    Helper.writeToFile(writeToLog, 'currentAsk', dataSheet,
                       15, currentRow, Helper.getDateFormula(transaction))

    Helper.writeToFile(writeToLog, 'NetPrem', dataSheet,
                       16, currentRow, """=IF([TransType]=\"LS\", [OpnPrem]+[ClsPrem],
                                            IF([TransType]=\"AS\", [OpnPrem]+[ClsPrem],
                                               [OpnPrem]-[ClsPrem]))""")
    Helper.writeToFile(writeToLog, 'TotPrem', dataSheet,
                       17, currentRow, """=IF([SYM]="","",SUMIFS([NetPrem],[Trade'#],[[Trade'#]],[Leg],"<="&[Leg]))""", moneyFormat=True)
    Helper.writeToFile(writeToLog, 'Days', dataSheet,
                       18, currentRow, """=IF(INDIRECT("RC[-14]", FALSE)="DEP",0,
IF(INDIRECT("RC[-14]", FALSE)="DIV",0,
IF(INDIRECT("RC[-14]", FALSE)="LS",TODAY() - INDIRECT("RC[-13]", FALSE),
IF(INDIRECT("RC[-14]", FALSE)="AS",TODAY() - INDIRECT("RC[-13]", FALSE),
IF(INDIRECT("RC[-12]", FALSE)>0,
IF(INDIRECT("RC[-12]", FALSE) - INDIRECT("RC[-13]", FALSE)>0,INDIRECT("RC[-12]", FALSE) - INDIRECT("RC[-13]", FALSE),1),
INDIRECT("RC[-11]", FALSE) - INDIRECT("RC[-13]", FALSE))))))""")
    Helper.writeToFile(writeToLog, 'Cap', dataSheet,
                       19, currentRow, """=IF([TransType]="SP",100*[Strike]*[['#Contracts]],
IF([TransType]="LS",[Strike]*[['#Contracts]],
IF([TransType]="BP",100*([Strike]-[Strike2])*[['#Contracts]],
IF([TransType]="BC",100*([Strike2]-[Strike])*[['#Contracts]],
IF([TransType]="NC",100*[Strike]*[['#Contracts]],
IF([TransType]="AS",100*[Strike]*[['#Contracts]],0))))))""")
    Helper.writeToFile(writeToLog, 'OpenCap', dataSheet,
                       20, currentRow, """=IF([CloseDate]>0,"",[Cap])""")
    Helper.writeToFile(writeToLog, 'BEcap', dataSheet,
                       21, currentRow, """=IF([CloseDate]>0,"",
IF([TransType]="BC","",
IF([TransType]="BP",100*[Strike]*[['#Contracts]],
[Cap])))""")
    Helper.writeToFile(writeToLog, 'CapDays', dataSheet,
                       22, currentRow, """=IF([SYM]="","",[Cap]*[Days])""")
    Helper.writeToFile(writeToLog, 'TotCapDays', dataSheet,
                       23, currentRow, """=IF([SYM]="","",SUMIFS([CapDays],[Trade'#],[[Trade'#]],[Leg],"<="&[Leg]))""")
    Helper.writeToFile(writeToLog, 'AROI', dataSheet,
                       24, currentRow, """=IF([TotCapDays],365*[TotPrem]/[TotCapDays],"")""", percentFormat=True, color='FFFF00')
    Helper.writeToFile(writeToLog, 'BreakEven', dataSheet,
                       25, currentRow, """=IF([SYM]="","",
IF([TransType]="LS",[Strike]-[TotPrem]/[['#Contracts]],
IF([['#Contracts]],[Strike]-[TotPrem]/[['#Contracts]]/100,"")))""", twoDecimalsFormat=True)
    Helper.writeToFile(writeToLog, 'ActShares', dataSheet,
                       26, currentRow, """=IF([CloseDate]>0,"",
IF([TransType]="LS",['#Contracts],
IF([TransType]="AS",100*[['#Contracts]],
IF([TransType]="SP",100*[['#Contracts]],
IF([TransType]="BP",100*[['#Contracts]],"")))))""")
    Helper.writeToFile(writeToLog, 'ActDate', dataSheet,
                       27, currentRow, """=IF([CloseDate]>0,[CloseDate],
IF([ExpDate]>0,[ExpDate],
TODAY()))""", dateFormat=True)
    Helper.writeToFile(writeToLog, 'Inception', dataSheet,
                       28, currentRow, """=IF(PERFORMANCE!D8>0,
IF(PERFORMANCE!D8<[ActDate],"",[NetPrem]),
IF(TODAY()<[ActDate],"",[NetPrem]))""")
    Helper.writeToFile(writeToLog, '1YR', dataSheet,
                       29, currentRow, """=IF(PERFORMANCE!D8>0,
IF(PERFORMANCE!D8<[ActDate],"",
IF(PERFORMANCE!D8 - [ActDate]<366,[NetPrem],"")),
IF(TODAY()<[ActDate],"",
IF(TODAY() - [ActDate]<366,[NetPrem],"")))""")
    Helper.writeToFile(writeToLog, 'YTD', dataSheet,
                       30, currentRow, """=IF(PERFORMANCE!D8>0,
IF(PERFORMANCE!D8<[ActDate],"",
IF(YEAR(PERFORMANCE!D8)=YEAR([ActDate]),[NetPrem],"")),
IF(TODAY()<[ActDate],"",
IF(YEAR(TODAY())=YEAR([ActDate]),[NetPrem],"")))""")
    Helper.writeToFile(writeToLog, 'UnBooked', dataSheet,
                       31, currentRow, """=IF(PERFORMANCE!D8>0,
IF(PERFORMANCE!D8<[ActDate],[NetPrem],""),
IF(TODAY()<[ActDate],[NetPrem],""))""")
    Helper.writeToFile(writeToLog, 'TotPrem', dataSheet,
                       33, currentRow, """=IF(Table1[SYM]="","",SUMIFS(Table1[NetPrem],Table1[Trade'#],Table1[[Trade'#]],Table1[Leg],"<="&Table1[Leg]))""")


def addTable(dataSheet, startingRow, endRow):
    del dataSheet.tables["Table1"]
    rowsAndColumns = "A" + str(startingRow) + ":AE" + str(endRow)
    tab = Table(displayName="Table1", ref=rowsAndColumns)

    style = TableStyleInfo(name="TableStyleMedium7", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style

    dataSheet.add_table(tab)
