# import tabula
#
# # data = tabula.read_pdf("./pdf1.pdf" , pages="all")
# # print(data)
#
# tabula.convert_into("pdf1.pdf","pdf1.csv",page="all",output_format="csv")

# Import the required Module
from tabula import read_pdf, convert_into

# Read a PDF File
# df = read_pdf("Test 4D29 Step 1 generated.pdf", pages='all')
# # convert PDF into CSV
# convert_into("Test 4D29 Step 1 generated.pdf", "Test 4D29 Step 1 generated.csv", output_format="csv", pages='all')


df1 = read_pdf("Test 4D29 Step 1.pdf", pages='all')
# # convert PDF into CSV
convert_into("Test 4D29 Step 1.pdf", "Test 4D29 Step 1.csv", output_format="csv", pages='all')
# print(df1)