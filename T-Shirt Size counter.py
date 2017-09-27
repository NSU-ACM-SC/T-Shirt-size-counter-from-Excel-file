import xlrd
file_location = input("Enter the location of the XLSX (excel book) file: ")
"""For Example: C:/Users/USER/Desktop/PoloShirt.xlsx"""
sizeCol = int(input("Enter the column index number for T-shirt sizes: "))

workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

m,l,xl,xxl,xl3 = 0,0,0,0,0

for row in range(sheet.nrows):
        if sheet.cell_value(row, sizeCol)=='M':
            m=m+1
        elif sheet.cell_value(row,sizeCol)=='L':
            l=l+1
        elif sheet.cell_value(row,sizeCol)=='XL':
            xl=xl+1
        elif sheet.cell_value(row,sizeCol)=='XXL':
            xxl=xxl+1
        elif sheet.cell_value(row,sizeCol)=='3XL':
            xl3=xl3+1
        else:
            if row!=0:
                print("Wrong kind of data in Size. Entry (" + str(row) + ","+ str(sizeCol) +")" )


print("Number of M size T shirts: " + str(m))
print("Number of L size T shirts: " + str(l))
print("Number of XL size T shirts: " + str(xl))
print("Number of XXL size T shirts: " + str(xxl))
print("Number of 3XL size T shirts: " + str(xl3))

print("Total: " + str(m+l+xl+xxl+xl3))

reM = int(input("Duplicates of Size M: "))
reL = int(input("Duplicates of Size L: "))
reXL = int(input("Duplicates of Size XL: "))
reXXL = int(input("Duplicates of Size XXL: "))
reXL3 = int(input("Duplicates of Size 3XL: "))

print("FINAL:")
print("Number of M size T shirts: " + str(m-reM))
print("Number of L size T shirts: " + str(l-reL))
print("Number of XL size T shirts: " + str(xl-reXL))
print("Number of XXL size T shirts: " + str(xxl-reXXL))
print("Number of 3XL size T shirts: " + str(xl3-reXL3))
print("Total: " + str(m+l+xl+xxl+xl3 -reM-reL-reXL-reXXL-reXL3))



                        


        
