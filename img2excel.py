from PIL import Image
import numpy as np
import xlsxwriter


def rgb2hex(r: int, g: int, b: int) -> str:
    return "#" + hex((0x01 << 24) + (r << 16) + (g << 8) + b)[3:]


def pixel_size_to_pixel_count(string: str) -> float:
    if string == "s":
        return 1
    elif string == "l":
        return 10
    else:
        return 1


def img2excel(img_path: str, save_path: str, sheet_name: str = "img", pixel_size: str = "s"):
    img = np.array(Image.open(img_path).convert("RGB"))

    workbook = xlsxwriter.Workbook(save_path)
    sheet = workbook.add_worksheet(sheet_name)

    # the dict to cache the format
    known_color = {}
    for row in range(img.shape[0]):
        for col in range(img.shape[1]):
            R, G, B = img[row, col]

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
    # but there are still some problems,
    # the pixel value is not accurate enough as expected.
    piexel_count = pixel_size_to_pixel_count(pixel_size)
    row_height = piexel_count * 0.75
    if piexel_count >= 13:
        col_width = (piexel_count - 5) / 8.0
    else:
        col_width = piexel_count * 0.077
    col_width = float(format(col_width, ".2f"))

    # Set the column width and row height to make cells become square.
    sheet.set_column(0, img.shape[1]-1, col_width)

    for row_num in range(img.shape[0]):
        sheet.set_row(row_num, row_height)

    workbook.close()


if __name__ == "__main__":
    img2excel("sample.jpg", "result.xlsx")
