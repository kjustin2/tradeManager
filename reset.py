import os
import shutil

os.remove('testingPortfolio.xlsx')
shutil.copyfile('reset.xlsx', 'reset2.xlsx')
os.rename('reset2.xlsx', 'testingPortfolio.xlsx')
