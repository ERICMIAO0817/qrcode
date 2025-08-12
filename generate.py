# import modules
import os

import qrcode
from PIL import Image as pil
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter


class Generator:
    # def run(self, input_path, out_path):
    #     # df = pd.read_csv('output/妇幼家具修改1.csv')
    #     df = pd.read_excel(input_path)
    #     df.to_csv('temp.csv', index=False)
    #     df = pd.read_csv('temp.csv')
    #     titles = list(df.head(0))[:-1]
    #     # print(titles)
    #     for index, row in df.iterrows():
    #         strs = ''
    #         for i in tqdm.tqdm(range(0, len(titles))):
    #             strs += titles[i] + ':' + str(row[i]) + '\n'
    #         self.gen_code(strs, index, out_path)

    def gen_code(self, contains, name, sub_dir):
        QRcode = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H
        )
        QRcode.add_data(contains)
        QRcode.make()
        QRcolor = 'black'
        QRimg = QRcode.make_image(
            fill_color=QRcolor, back_color="white").convert('RGB')
        QRimg.save(f'{sub_dir}/{name}.png')
        # print(f'二维码{name}已生成!')
