"""
2つのプロジェクト計画書（計画変更届）のうちの特定のシートに含まれる番地ごとの情報を確認し、
両者に差分があった場合に書き込みまたは赤字変更の処理を行う関数を定義する。
"""
from typing import List
import numpy as np
from constants import KeikakuSheet, Color
import settings
import utils

def perform(sheet_name: KeikakuSheet, target_ws, referred_ws, 
            compare_address_list: List[str], how: str) -> None:
    """概要
    与えられたセル番地に対して、2つのワークシートの値を比較し、両者が異なる場合は一方のワークシートに対して
    値の書き写しまたは赤字表示の処理を行う。
    
    Parameters
    ----------
    sheet_name
        ワークシートのシート名を示すKeikakuSheet型。

    target_ws
        書き写されるワークシート。

    referred_ws
        値を参照するワークシート。

    compare_address_list: List[str]
        値を比較する範囲のstr型を格納するlist型。

    how: str
        差分のあるセル範囲に対して、値を書き写すか、赤字表示にするかを指定するstr型。
        copyが与えられれば書き写し、checkが与えられれば赤字表示にする。それ以外の値はValueErrorを返す。

    Returns
    ----------
    None
    """
    target_wk = target_ws.UsedRange
    target_value = np.array(target_wk.Value)
    referred_wk = referred_ws.Range(target_wk.Address)
    referred_value = np.array(referred_wk.Value)
    referred_cell = utils.get_cell_address_from_range_address(target_wk.Address)
    referred_cell_loc = utils.from_cell_address_to_column_row_int(referred_cell)

    for address in compare_address_list:
        relative_address_loc = utils.relative_range_address_loc(address, referred_cell_loc)
        try:
            target_array = target_value[relative_address_loc[0][0]:relative_address_loc[1][0]+1,\
                                relative_address_loc[0][1]:relative_address_loc[1][1]+1]
            referred_array = referred_value[relative_address_loc[0][0]:relative_address_loc[1][0]+1,\
                                relative_address_loc[0][1]:relative_address_loc[1][1]+1]
            if not (referred_array == target_array).all():
                if how == 'copy':
                    _write(sheet_name, target_ws, address, referred_array)
                elif how == 'check':
                    _make_red(target_ws, address)
                else:
                    raise ValueError('howにはcopyまたはcheckを指定してください。')
        except Exception as e:
            print(e)
            print(sheet_name.value, address)
    return

def _write(sheet_name: KeikakuSheet, target_ws, address: str, referred_array: np.array) -> None:
    """概要
    ワークシートに対して、指定したアドレスに指定した値を書き込む。

    Parameters
    ----------
    sheet_name
        ワークシートのシート名を示すKeikakuSheet型。

    target_ws
        値を書き込むワークシート。

    address: str
        値を書き込む範囲を示すstr型。

    referred_array: np.array
        書き込む値を示すnp.array型。

    Returns
    ----------
    None
    """
    if sheet_name in settings.NUM_TO_STR_ADDRESS_DICT.keys():
        if address in settings.NUM_TO_STR_ADDRESS_DICT[sheet_name]:
            referred_array = np.vectorize(utils.from_str_num_to_text)(referred_array)
    target_ws.Range(address).Value = referred_array

def _make_red(target_ws, address: str) -> None:
    """概要
    ワークシートに対して、指定したアドレスの字を赤字にする。
    
    Parameters
    ----------
    target_ws
        字を赤字にするワークシート。

    address: str
        字を赤字にする範囲を示すstr型。
        
    Returns
    ----------
    None
    """
    target_ws.Range(address).Font.Color = Color.RED.value
    return