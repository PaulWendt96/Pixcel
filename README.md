## Pixcel
Enthrall your friends. Impress your boss. Enjoy art in the comfort of your own ```Excel.Application```. (Disclaimer - I did not take the picture in the workbook - someone much more talented than I took that shot!). If you're looking for art advice, best to look elsewhere.

<img src="/Sample/sample_soap.PNG" alt="Soap Bubble in Excel"/>

## How do I install Pixcel?
```bash
git clone https://www.github.com/PaulWendt96/Pixcel
```

### How do I used Pixcel?
```bash
python Pixcelize.py -picture_path -save_path --scale SCALE
```

```picture_path``` can be either a single image or a path to multiple images. ```save_path``` should point to a directory where excel workbooks with beautiful Pixcelized art can be saved ```scale``` should indicate how many pixels the larger dimension of each image should have. Pixels correspond with resolution, so passing a higher number to the --scale
argument will generally result in higher-quality images, but will generally raise runtime as well.

## Why is Excel not the greatest art-viewing tool in the world?
Excel is not the greatest art-viewing tool in the world because it imposes some limits on how many unique formats cells can have (65,490. I had to go through about ~2 pages of Microsoft VBA documentation to look that up). Excel also restricts zooming; it does not allow you to zoom any further than 10%. Have fun with the runtime too - excel is slow at adding cell formats to workbooks. I've tried speeding this up too, since writes to Excel gain a lot of speed when you assign an array to ```Range.Value```. I can't find the link now, but I swear I found something that says that Excel can only write cell formats cell-by-cell, and that adding cell formats one by one is the only way you can go. If you know this not to be true, please let me know!

## Contributing
Pull requests are welcome.
