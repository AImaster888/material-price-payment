# -*- coding: utf-8 -*-
"""
清掉 xlsx 檔案屬性裡的「上次修改者」。

Excel 存檔時會自動把本機使用者名稱寫進 docProps/core.xml 的 <cp:lastModifiedBy>，
範例檔要公開，不能帶出個人姓名。不能用 openpyxl 開啟再存（會丟掉公式計算結果），
所以直接改 zip 裡的那一個 XML，其他內容原封不動。

用法：python strip_metadata.py 檔案1.xlsx [檔案2.xlsx ...]
"""
import re
import shutil
import sys
import tempfile
import zipfile

CORE = 'docProps/core.xml'


def strip(path):
    tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    tmp.close()
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp.name, 'w') as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == CORE:
                text = data.decode('utf-8')
                text = re.sub(r'<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>',
                              '<cp:lastModifiedBy></cp:lastModifiedBy>', text)
                data = text.encode('utf-8')
            dst.writestr(info, data)
    shutil.move(tmp.name, path)


if __name__ == '__main__':
    for p in sys.argv[1:]:
        strip(p)
        print('cleared lastModifiedBy: {0}'.format(p))
