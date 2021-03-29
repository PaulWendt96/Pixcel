import cv2
from win32com.client import Dispatch
from contextlib import contextmanager
import time
import os
import logging
import argparse

MAX_WIDTH = 312
MAX_HEIGHT = 386


def get_image(path, width=100, height=100):
    image = cv2.imread(path, cv2.IMREAD_COLOR)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    width = min(MAX_WIDTH, width)
    height = min(MAX_HEIGHT, height)
    new_image = cv2.resize(image, (width, height))
    return new_image

def get_image_scale(path, higher_dim=min(MAX_WIDTH, MAX_HEIGHT)):
    image = cv2.imread(path, cv2.IMREAD_COLOR)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    dimensions = image.shape
    current_height, current_width, _ = dimensions
    scale = current_height/current_width

    ratio = current_width/current_height
    if scale > 1: # height > width
        height = int(higher_dim)
        width = int(higher_dim * (current_height/current_width) ** -1)
    else:
        width = int(higher_dim)
        height = int(higher_dim * (current_height/current_width))
    new_image = cv2.resize(image, (width, height))
    return new_image


def image_toRGBdict(image):
    col, row, _ = image.shape
    pixels = {}
    for c in range(col):
        for r in range(row):
            pixel = (c, r)
            pixels[pixel] = rgb_to_hex(tuple(image[c][r])[::-1])
    return pixels

def rgb_to_hex(rgb):
    strvalue = '%02x%02x%02x' % rgb
    ivalue = int(strvalue, 16)
    return ivalue

def write_to_excel(rgbs: dict, save_as: str, worksheet: str='Sheet1') -> None:

    @contextmanager
    def open_excel(path: str):
        col_start = 1
        col_end = 1000
        app = Dispatch('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Add()
        for x in range(col_start, col_end):
            workbook.Worksheets['Sheet1'].Columns(x).ColumnWidth = 2
        yield workbook
        workbook.SaveAs(path)
        app.DisplayAlerts = True
        app.Quit()

    with open_excel(save_as) as wb:
        s = wb.Worksheets(worksheet)
        for cell_addr, cell_rgb in rgbs.items():
            row, col = cell_addr
            row += 1
            col += 1
            s.Cells(row, col).Interior.Color = cell_rgb
    return


def image_to_excel(image_path, save_as, worksheet='Sheet1', do_scaling=True, scale=250, height=100, width=100):
    MAX_FORMATS = 65490 # hardcoded limit; see support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
    if do_scaling:
        image = get_image_scale(image_path, higher_dim=scale)
    else:
        image = get_image(image_path, height=height, width=width)
    rgb_dict = image_toRGBdict(image)
    if len(set(rgb_dict.values())) > MAX_FORMATS:
        raise ValueError("Error -- more cell formats than excel will allow (max is {})".format(MAX_FORMATS))
    write_to_excel(rgb_dict, save_as)


def pic_dir_to_excel_dir(pic_dir, excel_dir, do_scaling=True, scale=250, width=100, height=100):
    assert os.path.isdir(pic_dir)
    assert os.path.isdir(excel_dir)
    for file in os.listdir(pic_dir):
        print('Processing {}...'.format(file))
        try:
            file_path = os.path.join(pic_dir, file)
            file_name, _ = file.split('.')
            save_as = os.path.join(excel_dir, '{}.xlsx'.format(file_name))
            image_to_excel(file_path, save_as, do_scaling=do_scaling, scale=scale, height=height, width=width)
        except BaseException as e:
            print('Error processing file {}. Proceeding to next file'.format(file))
            logging.error(e)


def multiprocess(pics, excel_dir=None, do_scaling=True, scale=250, width=None, height=None):
    import os
    import multiprocessing

    if not excel_dir:
        excel_dir = os.getcwd()

    if do_scaling:
        assert scale
        assert not width
        assert not height
    else:
        assert not scale
        assert width
        assert height

    jobs = []
    if os.path.isdir(pics):
        for file in os.listdir(pics):
            print('Starting {}...'.format(file))
            file_path = os.path.join(pics, file)
            file_name, _ = file.split('.')
            save_as = os.path.join(excel_dir, '{}.xlsx'.format(file_name))
            p = multiprocessing.Process(target=image_to_excel, args=(file_path, save_as, None, do_scaling, scale, height, width))
            p.start()
            jobs.append(p)

        for job in jobs:
            job.join()

    elif os.path.isfile(pics):
        file_name, _ = pics.split('.')
        save_as = os.path.join(excel_dir, '{}.xlsx'.format(file_name))
        p = multiprocessing.Process(target=image_to_excel, args=(pics, save_as, None, do_scaling, scale, height, width))
        p.start()
        p.join()

    else:
        raise ValueError("Error -- 'pics' argument (currently '{}') must be either file or directory".format(pics))


if __name__ == '__main__':
    pics = r"C:\Users\paul_\PycharmProjects\Picxel\Pictures"
    save_dir = r"C:\Users\paul_\PycharmProjects\Picxel\Workbooks"
    test = r"C:\Users\paul_\PycharmProjects\Picxel\Pictures\david.jpg"

    parser = argparse.ArgumentParser(description='Pixelate image(s) to Excel')
    parser.add_argument("--scale", default=200, type=int)
    parser.add_argument('picture', help='picture file OR directory in which pictures are stored')
    parser.add_argument('save', help="directory in which completed excel workbooks should be saved")
    args = parser.parse_args()

    assert os.path.isdir(args.save)

    picture = os.path.abspath(args.picture)
    save = os.path.abspath(args.save)
    scale = args.scale

    import time
    start_time = time.time()
    multiprocess(picture, save, scale=scale)
    end_time = time.time()
    runtime = end_time - start_time
    print('{} seconds elapsed'.format(runtime))