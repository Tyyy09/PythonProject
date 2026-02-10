#Georgian College
#Class COMP1112
#student: Ichty Te
#ID: 200626964
#LAB: 12

from PyPDF2 import PdfReader
import time

# Open the PDF file
pdfFile = open(r"D:\week 14\test.pdf", "rb")
reader = PdfReader(pdfFile)

charset = "abcdefghi"

startTime = time.time()

if reader.is_encrypted:
    print("Starting brute force")
    # 6 nested loops to generate all 6‑letter combos
    for c1 in charset:
        for c2 in charset:
            for c3 in charset:
                for c4 in charset:
                    for c5 in charset:
                        for c6 in charset:
                            for c7 in charset:
                                for c8 in charset:
                                    for c9 in charset:
                                        pwd = c1 + c2 + c3 + c4 + c5 + c6 +c7 + c8 + c9
                                        result = reader.decrypt(pwd)
                                        if result != 0:
                                            print(f"Working with {charset}")
                                            endTime = time.time()
                                            print(f"Time taken: {endTime - startTime:.4f} seconds")
                                            exit()

