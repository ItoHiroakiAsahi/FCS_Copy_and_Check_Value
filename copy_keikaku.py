"""
プロジェクト計画書（計画変更届）の
シミュレーションに依存しない項目の書き写しを行う関数を定義する。
"""
import os
import argparse
import numpy as np
import pandas as pd
import win32com.client
import compare
import settings
import utils

def _copy_col_width_row_height(target_ws, referred_ws) -> None:
    """概要
    参照シートの列の幅、行の高さを対象のシートに反映する。

    Parameters
    ----------
    target_ws
        列の幅、行の高さを反映する対象のシート。

    referred_ws
        列の幅、行の高さを参照するシート。

    Returns
    ----------
    None
    """
    referred_wk = referred_ws.UsedRange
    target_wk = target_ws.Range(referred_wk.Address)
    target_value = np.array(target_wk.Value)
    col_list, row_list = utils.from_range_address_to_each_col_row_list(referred_wk.Address)
    col_list.reverse()
    row_list.reverse()
    is_all_col_nan = True
    is_all_row_nan = True
    for col in col_list:
        if is_all_col_nan and pd.isna(target_value[:, col_list.index(col)]).all():
            continue
        else:
            is_all_col_nan = False
        if target_ws.Columns(col).ColumnWidth != referred_ws.Columns(col).ColumnWidth:
            target_ws.Columns(col).ColumnsWidth = referred_ws.Columns(col).ColumnWidth
    for row in row_list:
        if is_all_row_nan and pd.isna(target_value[row_list.index(row)]).all():
            continue
        else:
            is_all_row_nan = False
        if target_ws.Rows(row).RowHeight != referred_ws.Rows(row).RowHeight:
            target_ws.Rows(row).RowHeight = referred_ws.Rows(row).RowHeight
    return

def copy_keikaku_value(target_keikaku_path: str, referred_keikaku_path: str,
                       save_path: str = '', ver: str = '1.3.0') -> None:
    """概要
    プロジェクト登録書に記載された内容のうち、シミュレーションに依存しない項目を
    別のプロジェクト登録書に対して書き写す。

    Parameters
    ----------
    target_keikaku_path: str
        書き写す対象のエクセルファイルのパスを示すstr型。
    
    referred_keikaku_path: str
        値を参照するエクセルファイルのパスを示すstr型。

    save_path: str
        書き写したエクセルファイルを保存するファイルパスを示すstr型。

    ver: str
        プロジェクト計画書のフォーマットを示すstr型。1.3.0のみを許容。

    Returns
    ----------
    None
    """
    if ver != '1.3.0':
        raise ValueError('現在プロジェクト登録書のフォーマットは1.3.0のみしか対応していません。')
    print('excelファイルを開いています。')
    app = win32com.client.Dispatch('Excel.Application')
    app.Visible = True
    target_wb = app.Workbooks.Open(os.getcwd() + '/' + target_keikaku_path)
    referred_wb = app.Workbooks.Open(os.getcwd() + '/' + referred_keikaku_path)

    for sheet_name in settings.COPY_CELL_ADDRESS_DICT.keys():
        target_ws = target_wb.Sheets(sheet_name.value)
        referred_ws = referred_wb.Sheets(sheet_name.value)

        # セルの内容をコピー
        compare.perform(sheet_name, target_ws, referred_ws, 
                        settings.COPY_CELL_ADDRESS_DICT[sheet_name], how='copy')

        # 行の高さ、列の幅を反映
        _copy_col_width_row_height(target_ws, referred_ws)

    if save_path == '':
        L = len('.xlsx')
        save_path = target_keikaku_path[:-L] + 'のコピー' + target_keikaku_path[-L:]
    app.DisplayAlerts = False
    target_wb.SaveAs(os.getcwd() + '/' + save_path)
    target_wb.Close()
    referred_wb.Close()
    app.Quit()
    app.DisplayAlerts = True
    return

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('keikaku_path', type = str, help = 'FilePath')
    args = parser.parse_args()
    copy_keikaku_value(args.keikaku_path)