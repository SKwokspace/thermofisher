from docx2pdf import convert

# Converting docx present in the same folder
# as the python file
convert("Test 4D29 Step 1.docx")

# Converting docx specifying both the input
# and output paths
# convert("GeeksForGeeks\GFG_1.docx", "Other_Folder\Mine.pdf")