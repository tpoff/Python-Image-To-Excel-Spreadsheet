# Python-Image-To-Excel-Spreadsheet

This is just a fun side project I made one afternoon.
The goal of the project was to create a python script that would take an image and convert it to an excel spreadsheet.

This is done by reading the rgb data for each pixel in the image and mapping them to cells in the spreadsheet. The 'pixels' in the excel spreadsheet are groups of 3 cells, each shaded to a different color of red green or blue, when viewed at a low enough zoom it creates the image. Much like how an actual monitor works.

I also coded a way to change the 'resolution' of the excel spreadsheet. So rather than being a 1 to 1 ratio from pixels in the image to "pixels" in the spreadsheet you can convert it so the excel spreadsheet uses fewer pixels. This can help speed up the processing on larger images, and it still looks cool.

The script uses openpyxl to create and edit excel spreadsheets and PIL to load the image and read the rgb data.

All in all it was rather simple to write thanks to these two libraries :-) 
It's a fun little script and I hope you enjoy using it!
