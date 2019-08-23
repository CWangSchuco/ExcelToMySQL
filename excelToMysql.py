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
    query = "INSERT INTO Article (ArticleId, ArticleGuid, Name, Unit, ArticleTypeId, CrossSectionUrl, Description, InsideDimension, " \
            "OutsideDimension, Dimension, RightSlideRebate, LeftSlideRebate" \
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s"

    print("GETTING EXCEL SHEET \n")

    # Open the workbook and define the worksheet
    book = xlrd.open_workbook("data.xlsx")
    sheet = book.sheet_by_name("InwardOpening")

    print("PROCESSING EXCEL SHEET \n")

    # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
    for r in range(1, sheet.nrows):

        print("PROCESSING ROW " + str(r) + " \n")

        profileName = sheet.cell(r,0).value
        system = sheet.cell(r,1).value
        type = sheet.cell(r,2).value

        if isinstance(sheet.cell(r, 3).value, numbers.Number):
            depth = sheet.cell(r,3).value
        else:
            depth = None

        if isinstance(sheet.cell(r, 4).value, numbers.Number):
            insideWidth = sheet.cell(r,4).value
        else:
            insideWidth = None

        if isinstance(sheet.cell(r, 5).value, numbers.Number):
            outsideWidth = sheet.cell(r,5).value
        else:
            outsideWidth = None

        if isinstance(sheet.cell(r, 6).value, numbers.Number):
            offsetReference = sheet.cell(r,6).value
        else:
            offsetReference = None

        if isinstance(sheet.cell(r, 7).value, numbers.Number):
            rightSideRebate = sheet.cell(r,7).value
        else:
            rightSideRebate = None

        if isinstance(sheet.cell(r, 8).value, numbers.Number):
            leftSideRebate = sheet.cell(r,8).value
        else:
            leftSideRebate = None

        ArticleId = r
        ArticleGuid = None
        Name = "article__" + profileName
        Unit = "mm"

        if (type == "Outer Frame"):
            ArticleTypeId = 1
        elif(type == "Vent Frame"):
            ArticleTypeId = 2
        elif(type == "Glazing Bead"):
            ArticleTypeId = 3
        elif (type == "Glazing Gasket"):
            ArticleTypeId = 4
        elif (type == "Frame Foam"):
            ArticleTypeId = 5
        elif (type == "Vent Frame Gasket"):
            ArticleTypeId = 6
        elif (type == "Glazing Rebate Gasket"):
            ArticleTypeId = 7
        elif (type == "Vent Foam"):
            ArticleTypeId = 8
        elif (type == "Center Gasket"):
            ArticleTypeId = 9
        elif (type == "Mullion"):
            ArticleTypeId = 10
        elif (type == "Transom"):
            ArticleTypeId = 11
        elif (type == "Intermediate"):
            ArticleTypeId = 12

        CrossSectionUrl = None
        Description = type + ' ' + profileName
        InsideDimension = insideWidth
        OutsideDimension = outsideWidth
        Dimension = None
        RightSlideRebate = rightSideRebate
        LeftSlideRebate = leftSideRebate


        # Assign values from each row
        values = (ArticleId, ArticleGuid, Name, Unit, str(ArticleTypeId), CrossSectionUrl, Description, InsideDimension, OutsideDimension, Dimension, RightSlideRebate, LeftSlideRebate)

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
