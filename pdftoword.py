# from pdf2docx import parse
# from typing import Tuple
#
# #
# # def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
# #     """Converts pdf to docx"""
# #     if pages:
# #         pages = [int(i) for i in list(pages) if i.isnumeric()]
# #
# #     result = parse(pdf_file=input_file, docx_with_path=output_file, pages=pages)
# #     summary = {
# #         "File": input_file, "Pages": str(pages), "Output File": output_file
# #     }
# #     # Printing Summary
# #     print("## Summary ########################################################")
# #     print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
# #     print("###################################################################")
# #     return result
#
#
# def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
#     """Converts pdf to docx"""
#     if pages:
#         pages = [int(i) for i in list(pages) if i.isnumeric()]
#     result = parse(pdf_file=r'C:\Users\MohanMadhuri\PycharmProjects\thermofisher\pdf1.pdf', op=output_file, pages=pages)


from pdf2docx import Converter,parse

pdf_file = "./pdf1.pdf"
word_file = "pdf1word1.docx"

#converter method
# cv = Converter(pdf_file)
#
# cv.convert(word_file,start=0,end=None)
# cv.close()

#parse method

parse(pdf_file,word_file,start=0,end=None)
