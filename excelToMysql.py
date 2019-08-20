import MySQLdb
import xlrd
import numbers


def do_the_thing():
    print("STARTING \n")

    print("ESTABLISHING MYSQL CONNECTION \n")

    # Establish a MySQL connection
    database = MySQLdb.connect(
        host="localhost", user="root", passwd="SchucoUSA1234!", db="VCLDesignDB")

    print("MYSQL CONNECTION ESTABLISHED \n")

    # Get the cursor, which is used to traverse the database, line by line
    cursor = database.cursor()

    # Create the INSERT INTO sql query
    query = "INSERT INTO AcousticReport (FileName, ProductName, ProductCode, ProductType, OpeningType, GlassLite1, GlassAirSpaceOne, GlassLite2, " \
            "GlassAirSpaceTwo, GlassLite3, TransmissionLoss50, TransmissionLoss63, TransmissionLoss80, TransmissionLoss100, " \
            "TransmissionLoss125, TransmissionLoss160, TransmissionLoss200, TransmissionLoss250, " \
            "TransmissionLoss315, TransmissionLoss400, TransmissionLoss500, TransmissionLoss630, " \
            "TransmissionLoss800, TransmissionLoss1000, TransmissionLoss1250, TransmissionLoss1600, " \
            "TransmissionLoss2000, TransmissionLoss2500, TransmissionLoss3150, TransmissionLoss4000, " \
            "TransmissionLoss5000, TransmissionLoss6300, TransmissionLoss8000, TransmissionLoss10000) " \
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,%s, %s, %s, %s, %s, %s, " \
            "%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"

    print("GETTING EXCEL SHEET \n")

    # Open the workbook and define the worksheet
    book = xlrd.open_workbook("data.xlsx")
    sheet = book.sheet_by_name("Facade")

    print("PROCESSING EXCEL SHEET \n")

    # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
    for r in range(1, sheet.nrows):

        print("PROCESSING ROW " + str(r) + " \n")

        file_name = sheet.cell(r, 0).value
        product_name = sheet.cell(r, 1).value
        product_code = sheet.cell(r, 2).value
        product_type = sheet.cell(r, 3).value

        if sheet.cell(r, 4).value == "-" or sheet.cell(r, 4).value == "":
            opening_type = None
        else:
            opening_type = sheet.cell(r, 4).value

        glass_lite1 = sheet.cell(r, 6).value
        glass_air_space_one = sheet.cell(r, 7).value
        glass_lite2 = sheet.cell(r, 8).value

        if sheet.cell(r, 9).value == '-' or sheet.cell(r, 9).value == '':
            glass_air_space_two = None
        else:
            glass_air_space_two = sheet.cell(r, 9).value

        if sheet.cell(r, 10).value == '-' or sheet.cell(r, 10).value == '':
            glass_lite3 = None
        else:
            glass_lite3 = sheet.cell(r, 10).value

        if isinstance(sheet.cell(r, 11).value, numbers.Number):
            transmission_loss50 = sheet.cell(r, 11).value
        else:
            transmission_loss50 = None

        if isinstance(sheet.cell(r, 12).value, numbers.Number):
            transmission_loss63 = sheet.cell(r, 12).value
        else:
            transmission_loss63 = None

        if isinstance(sheet.cell(r, 13).value, numbers.Number):
            transmission_loss80 = sheet.cell(r, 13).value
        else:
            transmission_loss80 = None

        transmission_loss100 = sheet.cell(r, 14).value
        transmission_loss125 = sheet.cell(r, 15).value
        transmission_loss160 = sheet.cell(r, 16).value
        transmission_loss200 = sheet.cell(r, 17).value
        transmission_loss250 = sheet.cell(r, 18).value
        transmission_loss315 = sheet.cell(r, 19).value
        transmission_loss400 = sheet.cell(r, 20).value
        transmission_loss500 = sheet.cell(r, 21).value
        transmission_loss630 = sheet.cell(r, 22).value
        transmission_loss800 = sheet.cell(r, 23).value
        transmission_loss1000 = sheet.cell(r, 24).value
        transmission_loss1250 = sheet.cell(r, 25).value
        transmission_loss1600 = sheet.cell(r, 26).value
        transmission_loss2000 = sheet.cell(r, 27).value
        transmission_loss2500 = sheet.cell(r, 28).value
        transmission_loss3150 = sheet.cell(r, 29).value
        transmission_loss4000 = sheet.cell(r, 30).value
        transmission_loss5000 = sheet.cell(r, 31).value

        if isinstance(sheet.cell(r, 32).value, numbers.Number):
            transmission_loss6300 = sheet.cell(r, 32).value
        else:
            transmission_loss6300 = None

        if isinstance(sheet.cell(r, 33).value, numbers.Number):
            transmission_loss8000 = sheet.cell(r, 33).value
        else:
            transmission_loss8000 = None

        if isinstance(sheet.cell(r, 34).value, numbers.Number):
            transmission_loss10000 = sheet.cell(r, 34).value
        else:
            transmission_loss10000 = None

        # Assign values from each row
        values = (file_name, product_name, product_code, product_type, opening_type, glass_lite1, glass_air_space_one, glass_lite2, transmission_loss50,
                  glass_air_space_two, glass_lite3, transmission_loss63, transmission_loss80, transmission_loss100, transmission_loss125,
                  transmission_loss160, transmission_loss200, transmission_loss250,
                  transmission_loss315, transmission_loss400, transmission_loss500, transmission_loss630,
                  transmission_loss800, transmission_loss1000, transmission_loss1250, transmission_loss1600,
                  transmission_loss2000, transmission_loss2500, transmission_loss3150, transmission_loss4000,
                  transmission_loss5000, transmission_loss6300, transmission_loss8000, transmission_loss10000)

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
    do_the_thing()


if __name__ == "__main__":
    main()
else:
    print(__name__)
