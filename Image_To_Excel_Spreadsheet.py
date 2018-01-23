from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill


from PIL import Image, ImageDraw, ImageFont

wb = load_workbook('./test.xlsx')
print(wb.get_sheet_names())

#get sheet object.
sheet1 = wb.get_sheet_by_name('Sheet1')
#print(sheet3['A1'].value)
#cell=sheet3['A1']

#sheet3.max_row
#sheet3.max_column


#sheet3['B2'] = "YO MAMA"



im = Image.open("image.jpg")
width,height = im.size
rgb_im = im.convert('RGB')

resolution = 1
print(width)
print(int(width*resolution))
print()
print(height)
print(int(height*resolution))
for x in range(0, int(width*resolution)):
    print("column: "+str(x))
    column_indicies = [get_column_letter((x*3)+1),get_column_letter((x*3)+2),get_column_letter((x*3)+3)]
    sheet1.column_dimensions[column_indicies[0]].width = 10/9
    sheet1.column_dimensions[column_indicies[1]].width = 10/9
    sheet1.column_dimensions[column_indicies[2]].width = 10/9
    for row in range(1, int(height*resolution)):
        rgb_array = rgb_im.getpixel((int(x/resolution),int(row/resolution)))
        for i in range(3):
            colors =[0,0,0]
            colors[i] = rgb_array[i]
            col = get_column_letter((x * 3) + i+1)

            cell = sheet1[col + str(row)]

            cell.value = rgb_array[i]
            color_string = "".join([str(hex(i))[2:].upper().rjust(2, "0") for i in colors])
            cell.fill=PatternFill(fill_type="solid", start_color='FF' + color_string, end_color='FF' + color_string)




wb.save('./test.xlsx')




