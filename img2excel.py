from PIL import Image
import numpy as np
import xlsxwriter


def rgb2hex(r, g, b):
    return "#" + hex((0x01 << 24) + (r << 16) + (g << 8) + b)[3:]


def img2excel(img_path, save_path):
    img = np.array(Image.open(img_path))

    workbook = xlsxwriter.Workbook(save_path)
    sheet = workbook.add_worksheet("img")

    known_color = {}
    for row in range(img.shape[0]):
        for col in range(img.shape[1]):
            R, G, B = img[row, col][:3]

            if (R, G, B) not in known_color:
                new_format = workbook.add_format()
                new_format.set_fg_color(rgb2hex(R, G, B))
                known_color.update({(R, G, B): new_format})

            this_format = known_color[(R, G, B)]
            sheet.write(row, col, "", this_format)

    sheet.set_column(0, img.shape[1]-1, 0.1)

    for row_num in range(img.shape[0]):
        sheet.set_row(row_num, 1)

    workbook.close()


if __name__ == "__main__":
    img2excel("sample.jpg", "result.xlsx")
