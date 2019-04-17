# -*- coding: utf-8 -*-
# Created with PyCharm
# @Time    : 2019-04-16 01:55
# @Author  : Tongjie Wang

import argparse
import os
import pathlib
import re
import shutil
import sys
import zipfile

TMP_DIR_PATH = '.xlsx_crack_tmp'
WORK_BOOK_PATH = 'xl/workbook.xml'
WORK_SHEETS_DIR_PATH = 'xl/worksheets'


def _make_copy(path: str) -> str:
    return shutil.copy2(path, path.replace('.xls', '.cracked.xls'))  # make it also work with .xlsm file


def _print_error_msg(e: Exception) -> None:
    print(f'{sys.argv[0]}: error: type: {str(type(e))[8:-2]}; message: {str(e)}',
          file=sys.stderr)
    exit(1)


def _parse_args() -> str:
    parser = argparse.ArgumentParser(description='Remove .xlsx file password')
    parser.add_argument(dest='file_name', help='path to source file', type=str)
    args = parser.parse_args()
    return args.file_name


def remove_password(path: str) -> None:
    try:
        # extract files to a directory
        copy_path = pathlib.Path(_make_copy(path))
        temp_dir = copy_path.parent / pathlib.Path(TMP_DIR_PATH)
        temp_dir.mkdir(exist_ok=True)
        file_copy = zipfile.ZipFile(copy_path)
        file_copy.extractall(temp_dir)
        file_copy.close()
        os.remove(str(copy_path))

        # remove password inside workbook
        with open(f'{temp_dir}/{WORK_BOOK_PATH}', 'r') as workbook:
            data = workbook.read()
            data = "".join(re.split(r'<workbookProtection[^>]*/>', data))

        with open(f'{temp_dir}/{WORK_BOOK_PATH}', 'w') as workbook:
            workbook.write(data)

        # remove password inside worksheet(s)
        worksheet_root = pathlib.Path(f'{temp_dir}/{WORK_SHEETS_DIR_PATH}')
        for item in worksheet_root.iterdir():
            result = re.match(r'^sheet(0|[1-9][0-9]*)\.xml$', str(item).rsplit('/', 1)[1])
            if item.is_file() and result is not None:
                with open(str(item), 'r') as worksheet:
                    data = worksheet.read()
                    data = "".join(re.split(r'<sheetProtection[^>]*/>', data))

                with open(str(item), 'w') as worksheet:
                    worksheet.write(data)

        # pack files into a zip file
        shutil.make_archive(str(copy_path), 'zip', temp_dir)
        shutil.rmtree(temp_dir)
        shutil.move(str(f'{copy_path}.zip'), copy_path)

        print(f'{sys.argv[0]}: notice: cracked file is {copy_path}')

    except zipfile.BadZipFile as e:
        _print_error_msg(e)
    except zipfile.LargeZipFile as e:
        _print_error_msg(e)
    except KeyError as e:
        _print_error_msg(e)
    except ValueError as e:
        _print_error_msg(e)
    except OSError as e:
        _print_error_msg(e)
    except Exception:
        print(f'{sys.argv[0]}: error: unknown error, see Traceback below\n', file=sys.stderr)
        raise


if __name__ == '__main__':
    remove_password(_parse_args())
