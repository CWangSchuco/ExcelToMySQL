import pdb
import MySQLdb
import xlrd
import numbers
import decimal


def doTheThing():
    
    print("STARTING \n")

    print("ESTABLISHING MYSQL CONNECTION \n")

    # Establish a MySQL connection
    database = MySQLdb.connect(
        host="localhost", user="root", passwd="", db="")

    print("MYSQL CONNECTION ESTABLISHED \n")

    # Get the cursor, which is used to traverse the database, line by line
    cursor = database.cursor()

    # Create the INSERT INTO sql query
    query = """INSERT INTO AcousticReport (TestNumber, ProductName, GlassLite1, GlassAirSpace, GlassLite2, TransmissionLoss50, TransmissionLoss63, TransmissionLoss80, TransmissionLoss100, TransmissionLoss125, TransmissionLoss160, TransmissionLoss200, TransmissionLoss250, TransmissionLoss315, TransmissionLoss400, TransmissionLoss500, TransmissionLoss630, TransmissionLoss800, TransmissionLoss1000, TransmissionLoss1250, TransmissionLoss1600, TransmissionLoss2000, TransmissionLoss2500, TransmissionLoss3150, TransmissionLoss4000, TransmissionLoss5000, TransmissionLoss6300, TransmissionLoss8000, TransmissionLoss10000) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

    print("GETTING EXCEL SHEET \n")

    # Open the workbook and define the worksheet
    book = xlrd.open_workbook("data.xlsx")
    sheet = book.sheet_by_name("Sheet1")

    print("PROCESSING EXCEL SHEET \n")

    # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
    for r in range(1, sheet.nrows):

            print("PROCESSING ROW " + str(r) + " \n")

            TestNumber = sheet.cell(r, 0).value
            ProductName = sheet.cell(r, 1).value
            GlassLite1 = sheet.cell(r, 4).value
            GlassAirSpace = sheet.cell(r, 5).value
            GlassLite2 = sheet.cell(r, 6).value

            if(isinstance(sheet.cell(r, 7).value, numbers.Number)):
                TransmissionLoss50 = sheet.cell(r, 7).value
            else:
                TransmissionLoss50 = None

            if (isinstance(sheet.cell(r, 8).value, numbers.Number)):
                TransmissionLoss63 = sheet.cell(r, 8).value
            else:
                TransmissionLoss63 = None

            if (isinstance(sheet.cell(r, 9).value, numbers.Number)):
                TransmissionLoss80 = sheet.cell(r, 9).value
            else:
                TransmissionLoss80 = None

            TransmissionLoss100 = sheet.cell(r, 10).value
            TransmissionLoss125 = sheet.cell(r, 11).value
            TransmissionLoss160 = sheet.cell(r, 12).value
            TransmissionLoss200 = sheet.cell(r, 13).value
            TransmissionLoss250 = sheet.cell(r, 14).value
            TransmissionLoss315 = sheet.cell(r, 15).value
            TransmissionLoss400 = sheet.cell(r, 16).value
            TransmissionLoss500 = sheet.cell(r, 17).value
            TransmissionLoss630 = sheet.cell(r, 18).value
            TransmissionLoss800 = sheet.cell(r, 19).value
            TransmissionLoss1000 = sheet.cell(r, 20).value
            TransmissionLoss1250 = sheet.cell(r, 21).value
            TransmissionLoss1600 = sheet.cell(r, 22).value
            TransmissionLoss2000 = sheet.cell(r, 23).value
            TransmissionLoss2500 = sheet.cell(r, 24).value
            TransmissionLoss3150 = sheet.cell(r, 25).value
            TransmissionLoss4000 = sheet.cell(r, 26).value
            TransmissionLoss5000 = sheet.cell(r, 27).value

            if(isinstance(sheet.cell(r, 28).value, numbers.Number)):
                TransmissionLoss6300 = sheet.cell(r, 28).value
            else:
                TransmissionLoss6300 = None

            if (isinstance(sheet.cell(r, 29).value, numbers.Number)):
                TransmissionLoss8000 = sheet.cell(r, 29).value
            else:
                TransmissionLoss8000 = None

            if (isinstance(sheet.cell(r, 30).value, numbers.Number)):
                TransmissionLoss10000 = sheet.cell(r, 30).value
            else:
                TransmissionLoss10000 = None

            # Assign values from each row
            values = (TestNumber, ProductName, GlassLite1, GlassAirSpace, GlassLite2, TransmissionLoss50,
                        TransmissionLoss63, TransmissionLoss80, TransmissionLoss100, TransmissionLoss125,
                        TransmissionLoss160, TransmissionLoss200, TransmissionLoss250,
                        TransmissionLoss315, TransmissionLoss400, TransmissionLoss500, TransmissionLoss630,
                        TransmissionLoss800, TransmissionLoss1000, TransmissionLoss1250, TransmissionLoss1600,
                        TransmissionLoss2000, TransmissionLoss2500, TransmissionLoss3150, TransmissionLoss4000,
                        TransmissionLoss5000, TransmissionLoss6300, TransmissionLoss8000, TransmissionLoss10000)

            print("ROW PROCESSED. INSERTING INTO DB \n")

            # Execute sql Query
            cursor.execute(query, values)

            print("INSERT COMPLETE \n")

    print("FINALIZING DB TRANSACTION")

    # Commit the transaction
    database.commit()

    print("CLOSING DATABASE CONNECTION \n")
    # Close the cursor
    cursor.close()

    # Close the database connection
    database.close()

    print("DATABASE CONNECTION CLOSED \n")

    # Print results
    print("IMPORT COMPLETE \n")
    columns = str(sheet.ncols)
    rows = str(sheet.nrows)
    print(columns + " COLUMNS AND " + rows + " ROWS PROCESSED")


def main():
    doTheThing()

if __name__ == "__main__":
    main()
else:
    print(__name__)
