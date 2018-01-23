from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
import openpyxl
from PIL import Image, ImageDraw, ImageFont


def Convert_Image_To_Excel_Spreadsheet(image_path, resolution=.3, output_file="results.xlsx"):
    wb = openpyxl.Workbook()
    sheet = wb.get_sheet_by_name("Sheet")
    im = Image.open(image_path)
    width, height = im.size
    rgb_im = im.convert('RGB')
    for x in range(0, int(width * resolution)):
        column_indicies = [get_column_letter((x * 3) + 1), get_column_letter((x * 3) + 2),
                           get_column_letter((x * 3) + 3)]
        sheet.column_dimensions[column_indicies[0]].width = 10 / 9
        sheet.column_dimensions[column_indicies[1]].width = 10 / 9
        sheet.column_dimensions[column_indicies[2]].width = 10 / 9
        for row in range(1, int(height * resolution)):
            rgb_array = rgb_im.getpixel((int(x / resolution), int(row / resolution)))
            for i in range(3):
                colors = [0, 0, 0]
                colors[i] = rgb_array[i]
                col = get_column_letter((x * 3) + i + 1)

                cell = sheet[col + str(row)]

                cell.value = rgb_array[i]
                color_string = "".join([str(hex(i))[2:].upper().rjust(2, "0") for i in colors])
                cell.fill = PatternFill(fill_type="solid", start_color='FF' + color_string,
                                        end_color='FF' + color_string)

    wb.save(output_file)




