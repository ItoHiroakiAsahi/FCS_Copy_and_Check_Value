"""
プロジェクト計画書（計画変更届）の差分を赤字にする関数を定義する。
"""
import argparse
import datetime
import os
import re
import win32com.client
import check_calc_sheets
import check_info_sheets
import check_rsh_sheets
import compare
from constants import KeikakuSheet, Color
import settings

def make_diff_red(target_file_path: str, referred_file_path: str, overwrite: bool = False,
                  save_path: str = '') -> None:
    """"概要
    2つのプロジェクト計画書（計画変更届）を受け取り、差分を赤字で表示する。

    Parameters
    ----------
    target_file_path: str
        差分を赤字にするエクセルファイルのパスを示すstr型。

    referred_file_path: str
        差分を確認する際に参照するエクセルファイルのパスを示すstr型。

    overwrite: bool, False
        ファイルの上書きを行うか否かを示すbool型。デフォルトはFalse。

    save_path: str, ''
        ファイルの上書きを行わない場合に、差分を赤字にしたファイルの保存先を示すstr型。
        ''が指定されている場合は、元のファイル名に時刻を加えて保存する。デフォルトは''。

    Returns
    ----------
    None
    """
    app = win32com.client.Dispatch('Excel.Application')
    app.Visible = True
    target_wb = app.Workbooks.Open(os.getcwd() + '/' + target_file_path)
    referred_wb = app.Workbooks.Open(os.getcwd() + '/' + referred_file_path)

    # シミュレーションに依存しない記入項目の差分を確認
    for sheet_name in list(set(list(settings.COMPARE_CELL_ADDRESS_DICT.keys()) 
                               + list(settings.COMPARE_AND_CHANGE_OTHER_CELL_VALUE_DICT.keys()))):
        target_sheet = target_wb.Sheets(sheet_name.value)
        referred_sheet = referred_wb.Sheets(sheet_name.value)
        # 差分がある場合に、該当するセルを赤字に変更
        if sheet_name in settings.COMPARE_CELL_ADDRESS_DICT.keys():
            compare.perform(sheet_name, target_sheet, referred_sheet, 
                            settings.COMPARE_CELL_ADDRESS_DICT[sheet_name], 'check')
        # 差分の有無に応じて、変更の有無のセルの値を変更
        if sheet_name in settings.COMPARE_AND_CHANGE_OTHER_CELL_VALUE_DICT.keys():
            compare.compare_and_change_other_cell_value(
                target_sheet, referred_sheet,settings.COMPARE_AND_CHANGE_OTHER_CELL_VALUE_DICT[sheet_name])

    # 情報記入シートの差分を確認
    for sheet_name in settings.INFO_SHEET_LIST:
        target_ws = target_wb.Sheets(sheet_name.value)
        referred_ws = referred_wb.Sheets(sheet_name.value)
        if sheet_name == KeikakuSheet.IKUSEI_INFO:
            compare.perform(sheet_name, target_sheet, referred_sheet, 
                            settings.IKUSEI_INFO_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_info_sheets.check_cell_address_list_ikusei_info(target_ws, referred_ws)
        elif sheet_name == KeikakuSheet.TENNEN_INFO:
            compare.perform(sheet_name, target_sheet, referred_sheet,
                            settings.TENNEN_INFO_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_info_sheets.check_cell_address_list_tennen_info(target_ws, referred_ws)
        elif sheet_name == KeikakuSheet.IN_PJ_EMISSION_INFO:
            compare.perform(sheet_name, target_sheet, referred_sheet,
                            settings.IN_PJ_EMISSION_INFO_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_info_sheets.check_cell_address_list_in_pj_emission_info(target_ws, referred_ws)
        elif sheet_name == KeikakuSheet.OUT_PJ_INFO:
            compare.perform(sheet_name, target_sheet, referred_sheet,
                            settings.OUT_PJ_INFO_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_info_sheets.check_cell_address_list_out_pj_info(target_ws, referred_ws)
        for address in l:
            target_ws.Range(address).Font.Color = Color.RED.value

    # 幹材積量算定シートの差分を確認
    for sheet_name in settings.RSH_SHEET_LIST:
        target_ws = target_wb.Sheets(sheet_name.value)
        referred_ws = referred_wb.Sheets(sheet_name.value)
        if sheet_name == KeikakuSheet.IKUSEI_RSH:
            compare.perform(sheet_name, target_ws, referred_ws,
                            settings.IKUSEI_RSH_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_rsh_sheets.check_cell_address_list_ikusei_rsh(target_ws, referred_ws)
        elif sheet_name == KeikakuSheet.TENNEN_RSH:
            compare.perform(sheet_name, target_ws, referred_ws,
                            settings.TENNEN_RSH_PARAMS.OUT_OF_PATTERN_CELL_LIST, 'check')
            l = check_rsh_sheets.check_cell_address_list_tennen_rsh(target_ws, referred_ws)
        for address in l:
            target_ws.Range(address).Font.Color = Color.RED.value

    # 吸収量算定シートの差分を確認
    for sheet_name in settings.CALC_SHEET_LIST:
        target_ws = target_wb.Sheets(sheet_name.value)
        referred_ws = referred_wb.Sheets(sheet_name.value)
        if sheet_name == KeikakuSheet.IKUSEI_CALCULATION:
            l = check_calc_sheets.check_cell_address_list_ikusei_calc(target_ws, referred_ws)
        elif sheet_name == KeikakuSheet.TENNEN_CALCULATION:
            l = check_calc_sheets.check_cell_address_list_tennen_calc(target_ws, referred_ws)
        for address in l:
            target_ws.Range(address).Font.Color = Color.RED.value

    app.DisplayAlerts = False
    if overwrite:
        target_wb.Save()
    else:
        if save_path == '':
            L = len('.xlsx')
            last_ref_path = re.split('/|"\\"', referred_file_path)[-1]
            dt = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            save_path = target_file_path[:-L] + '_赤字変更(参照ファイル：{})_{}'\
                .format(last_ref_path[:-L], dt) + target_file_path[-L:]
        target_wb.SaveAs(os.getcwd() + '/' + save_path)
    target_wb.Close()
    referred_wb.Close()
    app.Quit()
    app.DisplayAlerts = True
    return

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('target_keikaku_path', type = str, help = 'TargetFilePath')
    parser.add_argument('referred_keikaku_path', type = str, help = 'ReferredFilePath')
    args = parser.parse_args()
    make_diff_red(args.target_keikaku_path, args.referred_keikaku_path)