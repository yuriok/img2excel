from PIL import Image
import numpy as np
import xlsxwriter


def rgb2hex(r, g, b):
    return "#" + hex((0x01 << 24) + (r << 16) + (g << 8) + b)[3:]


def img2excel(img_path, save_path, sheet_name="img", pixel_size=1):
    img = np.array(Image.open(img_path))

    workbook = xlsxwriter.Workbook(save_path)
    sheet = workbook.add_worksheet(sheet_name)

    # the dict to cache the format
    known_color = {}
    for row in range(img.shape[0]):
        for col in range(img.shape[1]):
            R, G, B = img[row, col][:3]

            if (R, G, B) not in known_color:
                new_format = workbook.add_format()
                new_format.set_fg_color(rgb2hex(R, G, B))
                known_color.update({(R, G, B): new_format})

            this_format = known_color[(R, G, B)]
            # Set the foreground color of cell using format
            sheet.write_blank(row, col, "", this_format)

    # The relationships between pixel and column width is complex.
    # see https://support.microsoft.com/en-us/help/214123/description-of-how-column-widths-are-determined-in-excel
    # Although I have find the rule to convert pixel to width,
    # but it's still some problems,
    # the pixel value is not accurate enough as expected.
    row_height = pixel_size * 0.75
    if pixel_size >= 13:
        col_width = (pixel_size - 5) / 8.0
    else:
        col_width = pixel_size * 0.077
    col_width = float(format(col_width, ".2f"))

    # Set the column width and row height to make cells become square.
    sheet.set_column(0, img.shape[1]-1, col_width)

    for row_num in range(img.shape[0]):
        sheet.set_row(row_num, row_height)

    workbook.close()


if __name__ == "__main__":
    img2excel("sample.jpg", "result.xlsx")
