#  Date:   Aug 2014
#  Author: Britton J. Olson
#  Desc:   Convert sqlite database into an excel spreadsheet
#          (xls) file.
#  Notes:  Requires sqlite3 and xlwt python libraries
import sqlite3 as lite
import xlwt
import xlrd
import sys
import math
import csv

class sqlite2xls():
    def __init__(self,db):
        self.con = lite.connect(db)
        self.cur = self.con.cursor()
        self.tables = ''
        
    def getTables(self):
        # Get all the tables in this database
        cmd = "select name from sqlite_master where type = 'table' "
        self.cur.execute(cmd)
        self.tables = self.cur.fetchall()
        self.workbook = None
        
    def getWorkbook(self):
        self.getTables()
        self.workbook = xlwt.Workbook()
        for table in self.tables:
            # Get column heads
            self.cur.execute('pragma table_info(%s)' % table[0] )
            head = self.cur.fetchall()
            # Get row entries
            self.cur.execute('select * from %s' % table[0] )
            players = self.cur.fetchall()

            Np = len(players)
            cmax = 30000     # Max character per cell
            Rmax = 64000     # Max number of rows per sheet
            NS = 1
            if ( Np > Rmax):
                NS = int(math.ceil( float(Np)/float(Rmax) ) )
            for ss in range(NS):
                ips = ss*(Rmax)
                ipe = min( (ss+1)*Rmax, Np)
                # Open workbook and save head/body
                print table[0]
                if (ss < 1 ):
                    sheet = self.workbook.add_sheet(table[0])
                else:
                    sheet = self.workbook.add_sheet(table[0] + '_%s' % (ss) )
                # head
                for col,item in enumerate(head):
                    sheet.write(0,col,item[1])
                # body
                for row,player in enumerate(players[ips:ipe]):
                    for col,item in enumerate(player):
                        if ( type(item) == type(u'') ):
                            imax = min(cmax,len(item))
                            sheet.write(row+1,col,item[0:imax] )
                        else:
                            sheet.write(row+1,col,item )

    def writeXLS(self,out):
        self.getWorkbook()
        self.workbook.save(out)

    def writeCSV(self,out,sheetName):
        out_xls = out
        out_csv = out
        print out[-4:]
        if ( '.csv' not in out[-4:] ):
            out_csv = out + '.csv'
        if ( '.xls' not in out[-4:] ):
            out_xls = out + '.xls'
        self.getWorkbook()
        self.writeXLS(out_xls)
        self.csv_from_excel(out_xls,sheetName,out_csv)



    def csv_from_excel(self,xlsfile,sheetName,csvfile):
        wb = xlrd.open_workbook(xlsfile)
        sh = wb.sheet_by_name(sheetName)
        your_csv_file = open(csvfile, 'wb')
        wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
        
        for rownum in xrange(sh.nrows):
            wr.writerow(sh.row_values(rownum))

        your_csv_file.close()
    


class csv2sqlite():
    def __init__(self,csvFile):
        self.csvFile = csvFile
        self.delimiter = ','
        self.table = 'CSVFile'
        
    def initSQL(self):
        self.con = lite.connect(self.db)
        self.cur = self.con.cursor()
        
    def writeSQL(self,db):
        self.db = db
        self.initSQL()
        
        parsedCSV = []
        with open(self.csvFile) as csvfile:
            lines = csvfile.readlines()
            legend = lines[0].replace('\n','').split(self.delimiter)
            for row in lines[1:] :
                
                item = {}
                entries = row.replace('\n','').split(self.delimiter)
                for (cat,entry) in zip(legend,entries):
                    item[cat] = entry            
                
                parsedCSV.append(item)

            self.addSchema(legend)
            import pdb
            pdb.set_trace()
            for row in parsedCSV:
                self.addColumns(self.table,row)
            self.con.commit()

    def addSchema(self,titleRow):
        print "Adding new table: %s" % self.table
        cmd = 'CREATE TABLE %s (' % self.table
        for col in titleRow:
            print col.strip()
            cmd += "%s TEXT," % col.replace(' ','').strip()
        cmd += ")"
        cmd = cmd.replace(',)',')')  # Fix last comma
        print cmd
        self.cur.execute(cmd)

    def addColumns(self,table,columns):

        SQinsert = self.dictToTuple(columns)
        
        # Insert many to database
        key = '('
        sqsr = '('
        vals = []
        for item in SQinsert:
            key += '%s,' % (item[0].strip())
            sqsr += '? ,'
            vals.append( item[1].strip() )
            
        key += ')'
        key = key.replace(',)',')')
        sqsr += ')'
        sqsr = sqsr.replace(',)',')')
        cmd = "INSERT INTO %s %s VALUES %s" % (table,key,sqsr)
        try:
            self.cur.executemany(cmd,[tuple(vals)])
        except:
            print "Warning: SQL write Error"
            import pdb;pdb.set_trace()

    # Simple helper to convert a dictionary to a tuple
    # suitable for using in sqlite execute all command
    def dictToTuple(self,sqdict):
        tup = []
        for d in sqdict:
            tup.append( [d,sqdict[d]] )
        return tup


