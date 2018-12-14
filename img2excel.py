#!/usr/bin/python3

import argparse
import sys

import xlsxwriter
from PIL import Image


def rgb2hex(r: int, g: int, b: int) -> str:
    return "#" + hex((0x01 << 24) + (r << 16) + (g << 8) + b)[3:]


def pixel_size_to_pixel_count(string: str) -> float:
    if string == "s":
        return 1
    elif string == "l":
        return 10
    else:
        return 1


def img2excel(img_path: str, save_path: str, sheet_name: str = "Image",
              pixel_size: str = "s", use_palette: bool = True):
    # use palette to support high-resolution

    rgb_img = Image.open(img_path).convert("RGB")
    width, height = rgb_img.size
    workbook = xlsxwriter.Workbook(save_path)
    sheet = workbook.add_worksheet(sheet_name)

    if use_palette:
        palette_img = rgb_img.convert("P", palette=Image.ADAPTIVE, colors=256)
        mode, palette_data = palette_img.palette.getdata()
        palette_pixels = palette_img.load()
        # prepare formats of excel workbook
        format_dict = {}
        for i in range(256):
            r, g, b = palette_data[i*3:i*3+3]
            new_format = workbook.add_format()
            new_format.set_fg_color(rgb2hex(r, g, b))
            format_dict[i] = new_format

        for row in range(height):
            for col in range(width):
                palette_num = palette_pixels[col, row]
                this_format = format_dict[palette_num]
                # Set the foreground color of cell using format
                sheet.write_blank(row, col, "", this_format)

    else:
        rgb_pixels = rgb_img.load()
        # the dict to cache the format
        format_color = {}
        for row in range(height):
            for col in range(width):
                r, g, b = rgb_pixels[col, row]

                if (r, g, b) not in format_color:
                    new_format = workbook.add_format()
                    new_format.set_fg_color(rgb2hex(r, g, b))
                    format_color.update({(r, g, b): new_format})

                this_format = format_color[(r, g, b)]
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
    sheet.set_column(0, width-1, col_width)

    for row_num in range(height):
        sheet.set_row(row_num, row_height)

    workbook.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("image", type=str,
                        help="The path of the image you want to convert.")
    parser.add_argument("-p", action="store_true", default=False,
                        help="Use palette to support high resolution (e.g. 1920x1080).")
    parser.add_argument("-l", action="store_true", default=False,
                        help="Set the size of cells as 10 pixels, default is 1 pixel.")

    parser.add_argument("--savepath", type=str, default="result.xlsx",
                        help="The path to save the result.")
    parser.add_argument("--sheetname", type=str, default="Image",
                        help="The name of sheet in excel file.")

    args = parser.parse_args()

    if args.l:
        pixel_size = "l"
    else:
        pixel_size = "s"
    img2excel(args.image, args.savepath, sheet_name=args.sheetname,
              pixel_size=pixel_size, use_palette=args.p)


if __name__ == "__main__":
    sys.exit(main())
