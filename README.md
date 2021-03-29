## Pixcel
Turn this:

<img src="/Pictures/soap.jpg" alt="Soap Bubble"/>

Into this:

<img src="/Sample/sample_soap.PNG" alt="Soap Bubble in Excel"/>

## Installation
```bash
git clone https://www.github.com/PaulWendt96/Pixcel
```

### Usage
```bash
python Pixcelize.py -picture_path -save_path --scale SCALE
```

picture_path should be either the path to a single image OR a path to a directory storing multiple images
save_path should be a path to a directory in which excel workbooks can be saved
scale should indicate how many pixels the large dimension of each image should have. For instance,
for a 1080x1920 image, a --scale of 200 would indicate a height of 200 pixels and a width of 
200 * (1080/1920) pixels. Pixels correspond roughly with resolution, so passing a higher number to the --scale
argument will generally result in higher-quality images.

## Notes -- Excel Limitations
Note that Excel imposes some limits on the size and quality of the images we can produce with Pixcel.
Limitations that I ran into are below.

1. Excel can only zoom out up to 10x. For very large --scale values, you might not be able to see the entire picture even at maximum zoom. This usually does not become an issue for --scale arguments <= 800
2. Excel limits the number of unique cell formats to 65,490. At large enough --scale values (in which the number of cells being printed is greater than this limit), it is important to ensure that the number of individual colors to be used does not exceed 65,490. The script will raise an error if you try to produce an excel workbook with more than 65,490 different colors.
3. Excel is slow at adding cell formats to workbooks. While the script runs quickly for small --scale arguments, larger --scale arguments (>= 250) can lead to 5-10 minute runtimes. The primary reason for this is because excel needs to loop cell by cell, adding formats one at a time. I have not been able to figure out a way to apply formats faster than this. I believe this might be an inherent limitation of using Excel as an artistic tool, but if you know of a better approach please reach out to me!

## Contributing
Pull requests are welcome.

