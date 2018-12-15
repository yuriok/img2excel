# img2excel

A small python script to convert images to excel files.

## Requirements

* Pillow
* XlsxWriter

## Usages

### Quick run

`python img2excel.py [your image]`<br>

### High resolution image

If you want to convert a high resolution (e.g. 1920x1080 or 3840x2160) image, please pass `-p` (Palette) to use the palette of 256 colors, otherwise the output excel file will be broken. **Note that, if use the palette, the picture quality will get worse.**<br><br>

`python img2excel.py -p [your image]`<br><br>

<img width="384" height="208" alt="screenshot of sample1" src="https://raw.githubusercontent.com/yuriok/img2excel/master/screenshot_of_sample1.png"/><br>
**The result of `sample1` (resolution is 2048x2048)**<br>
The excel file can be downloaded at [here](https://raw.githubusercontent.com/yuriok/img2excel/master/result_of_sample1.xlsx).
The image `sample1` is from [Kerry Lambert](http://source.pixite.co/kerry-fin/kf-overlays?photo=59LbNx2oJubZTQPjn). Published under a [CC0 1.0 license](https://creativecommons.org/publicdomain/zero/1.0/).<br>

### Pixel art style

For default, the size of each cell is set as 1 pixel. If you want to show a image with the pixel art style, please pass `-l` (Large), this will make the size of cells become 10 pixels.<br>
`python img2excel.py -l [your image]`<br>

### Details of usage

<pre>
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
</pre>
