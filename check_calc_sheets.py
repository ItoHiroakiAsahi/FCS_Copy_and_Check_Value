"""
プロジェクト計画書（計画変更届）に含まれるシートのうち、
吸収量算定シートの情報を比較する関数を定義する。
"""
from typing import List
from check_info_sheets import _get_max_row_from_ws
import settings

def _check_address_list(target_ws, referred_ws, 
                        check_col_list: List[str], row_offset: int) -> List[str]:
    """概要
    吸収量算定シートの情報を比較し、差分のあるセルのアドレスをリストに格納する。

    Parameters
    ----------
    target_ws
        差分を赤字にする幹材積量算定シート。

    referred_ws
        差分を参照する幹材積量算定シート。
    
    check_col_list: List[str]
        差分を確認する列の列名を格納するList[str]型。

    row_offset: int
        行方向のオフセット数を示すint型。

    Returns
    ----------
    check_cell_address_list: List[str]
        差分のあるセルの番地を示すstr型を格納するList[str]型。
    """
    max_row = _get_max_row_from_ws(target_ws, referred_ws)
    check_address_list = []
    for col in check_col_list:
        compare_address_range = '{c}{r1}:{c}{r2}'.format(
            c = col, r1 = row_offset + 1, r2 = max_row + 1)
        # 数式に対する変更を確認
        target_value = target_ws.Range(compare_address_range).Formula
        referred_value = referred_ws.Range(compare_address_range).Formula
        for r_num in range(len(target_value)):
            if target_value[r_num][0] != referred_value[r_num][0]:
                check_address_list.append('{}{}'.format(col, row_offset + r_num + 1))
    return check_address_list

def check_cell_address_list_ikusei_calc(target_ws, referred_ws) -> List[str]:
    """概要
    2つの（自動計算）吸収量（育成林）算定シート（001、003共通）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する（自動計算）吸収量（育成林）算定シート（001、003共通）。
    
    referred_ws:
        差分を確認するための（自動計算）吸収量（育成林）算定シート（001、003共通）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するList[str]型。
    """
    return _check_address_list(target_ws, referred_ws, 
                               settings.IKUSEI_CALCULATION_PARAMS.CHECK_COL_LIST, 
                               settings.IKUSEI_CALCULATION_PARAMS.ROW_OFFSET)

def check_cell_address_list_tennen_calc(target_ws, referred_ws) -> List[str]:
    """概要
    2つの（自動計算）吸収量（天然生林）算定シート（FO-001）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する（自動計算）吸収量（天然生林）算定シート（FO-001）。
    
    referred_ws:
        差分を確認するための（自動計算）吸収量（天然生林）算定シート（FO-001）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するList[str]型。
    """
    return _check_address_list(target_ws, referred_ws,
                               settings.TENNEN_CALCULATION_PARAMS.CHECK_COL_LIST,
                               settings.TENNEN_CALCULATION_PARAMS.ROW_OFFSET)