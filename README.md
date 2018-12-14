# img2excel

A small python script to convert images to excel files.

## Requirements

* Pillow
* XlsxWriter

## Usage

### Quick run

`python img2excel.py [your image]`<br>

### High-resolution image

If you want to convert a high-resolution (e.g. 1920x1080) image, plese pass `-p` (Palette) to use the palette of 256 colors, otherwise the output excel file will be broked.<br>
`python img2excel.py -p [your image]`<br>

### Pixel art style

For default, the width of each cell is set as 1 pixel width. If you want to show a image with pixel art style, please pass `-l` (Large), this will make the width cells become 10 pixels.<br>
`python img2excel.py -l [your image]`<br>

```shell
usage: img2excel.py [-h] [-p] [-l] [--savepath SAVEPATH]
                    [--sheetname SHEETNAME]
                    image

positional arguments:
  image                 The path of the image you want to convert.

optional arguments:
  -h, --help            show this help message and exit
  -p                    Use palette to support high resolution (e.g.
                        1920x1080).
  -l                    Set the size of cells as 10 pixels, default is 1
                        pixel.
  --savepath SAVEPATH   The path to save the result.
  --sheetname SHEETNAME
                        The name of sheet in excel file.
```

## Preview

![Sample image](https://raw.githubusercontent.com/yuriok/img2excel/master/sample.jpg)<br>
Figure 1. The sample image<br>
![result screenshot](https://raw.githubusercontent.com/yuriok/img2excel/master/screenshot.jpg)<br>
Figure 2. The screenshot of result<br>
