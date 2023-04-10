from PIL import Image,ImageFont,ImageDraw


# Import the xlrd module
import xlrd
# Open the Workbook

workbook = xlrd.open_workbook("/Users/sam576/Downloads/RunWithBala/RHWB_S8_list.xlsx")

# Open the worksheet
worksheet = workbook.sheet_by_index(0)
i = 0
# Iterate the rows and columns
# for i in range(1, 332):
for i in range(1, 771):

    # Print the cell values with tab space worksheet.cell_value(i, 3)
    my_image = Image.open("/Users/sam576/Downloads/RunWithBala/RHWB_8.jpg")

    Imgwidth = my_image.size[0]
    # print("Img Width:" + str(Imgwidth))
    image_editable = ImageDraw.Draw(my_image)

    number_font =  ImageFont.truetype("/Users/sam576/Downloads/RunWithBala/Fonts/NotoSansDisplay-ExtraBold.ttf", 320)

    name_font = ImageFont.truetype("/Users/sam576/Downloads/RunWithBala/Fonts/NotoSansDisplay-ExtraBold.ttf", 135)
    contact_font = ImageFont.truetype("/Users/sam576/Downloads/RunWithBala/Fonts/BebasNeue-Regular.ttf", 75)


    #number_text = "235"
    number_text =  str(int(worksheet.cell_value(i, 0))).zfill(3)
    fullname_text  = worksheet.cell_value(i, 1)
    print(fullname_text)
    contact_present =  False
    t = worksheet.cell_value(i, 3)
    print("contact length="+str(len(t)))
    if len(t)>0:
        contact_text  = "Emergency Contact: "+worksheet.cell_value(i, 3)
        # contact_text  = "CRO: "+worksheet.cell_value(i, 3)
        # print(contact_text)
        contact_present =  True
    else:
        contact_text=""

    #name_text = "ADVIKAA"
    name_text = worksheet.cell_value(i, 2)
    draw = ImageDraw.Draw(my_image)
    textwidth = draw.textsize(name_text, font=name_font)[0]
    numberwidth = draw.textsize(number_text, font=number_font)[0]

    contactwidth = draw.textsize(contact_text, font=contact_font)[0]


    print("textwidth Width:" + str(textwidth))
    #name_xpos = 575 for 6 letter names
    name_length = len(name_text)
    # name_xpos = 700 for 3 Letters
    # name_xpos = 650 for 4 Letters
    # name_xpos = 600 for 5 Letters
    # name_xpos = 550 for 6 Letters
    # name_xpos = 500 for 7 Letters
    name_xpos = (Imgwidth - textwidth)/2
    number_xpos = (Imgwidth - numberwidth)/2
    contact_xpos = (Imgwidth - contactwidth)/2

    # image_editable.text((number_xpos,300), number_text, (0, 0, 0), font=number_font)
    # image_editable.text((name_xpos,750), name_text, (0, 0, 0), font=name_font)
    # image_editable.text((contact_xpos,750), contact_text, (0, 0, 0), font=contact_font)
    image_editable.text((number_xpos,250), number_text, (0, 0, 0), font=number_font)
    image_editable.text((name_xpos,620), name_text, (0, 0, 0), font=name_font)
    print(" before if, worksheet.cell_value= "+worksheet.cell_value(i, 3))

    # if contact_present:
        # print("  contact present for "+name_text)
        # image_editable.text((contact_xpos,910), contact_text, (0, 0, 0), font=contact_font)
    # print("")
    my_image.save("/Users/sam576/Downloads/RunWithBala/Bibs_new7/" + fullname_text + ".jpg")

