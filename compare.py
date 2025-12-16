"""
2つのプロジェクト計画書（計画変更届）のうちの特定のシートに含まれる番地ごとの情報を確認し、
両者に差分があった場合に書き込みまたは赤字変更の処理を行う関数を定義する。
"""
from typing import List, Dict, Tuple, Optional
import numpy as np
from constants import KeikakuSheet, Color, ChangeFlag
import settings
import utils

def _extract_array(array: np.array, relative_address_loc: Tuple[Tuple[int]]) -> np.array:
    """概要
    np.array型の2次元配列から、relative_address_locで指定した範囲を抽出する。

    Parameters
    ----------
    array: np.array
        抽出される2次元配列を格納するnp.array型。

    relative_address_loc: Tuple[Tuple[int]]
        抽出する範囲を示すint型を格納するTuple型。

    Returns
    ----------
    return_array: np.array
        抽出されたnp.array型。
    """
    return array[relative_address_loc[0][0]:relative_address_loc[1][0]+1,
                 relative_address_loc[0][1]:relative_address_loc[1][1]+1]

def _is_same(address: str, referred_cell_loc: str, target_value: np.array, 
             referred_value: np.array, return_value_address: Optional[str] = None)\
                -> Tuple[bool, Optional[np.array]]:
    """概要
    与えられたaddress, referred_cell_locから求まる相対位置の範囲に対して、
    2つのnp.array型を比較し、両者の値が同一であるか否かのbool型と、
    referred_valueから抽出した値のtuple型を返す。

    Parameters
    ----------
    address: str
        値を比較するセル範囲を示すstr型。

    referred_cell_loc: str
        target_value, referred_valueの値を格納する範囲のうち、左上のセルの座標を示すstr型。

    target_value: np.array
        値を比較するnp.array型。

    referred_value: np.array
        値を参照するnp.array型。
        与えられたreferred_valueのうち、return_value_addressに対応する範囲の値が返される。

    return_value_address: Optional[str] = None
        referred_valueから値を抽出する範囲を示すstr型。Noneの場合は、
        addressと同じ範囲を抽出する。デフォルトはNone。

    Return
    ----------
    check_tuple: Tuple[bool, Optional[np.array]]
        指定した範囲において、2つのnp.array型の持つ値が同一であるか否かを示すbool値と、
        referred_valueのうちreturn_value_addressで指定される範囲のnp.array型を格納するtuple型。
    """
    relative_address_loc = utils.relative_range_address_loc(address, referred_cell_loc)
    try:
        target_array = _extract_array(target_value, relative_address_loc)
        referred_array = _extract_array(referred_value, relative_address_loc)
        if return_value_address is None:
            return (referred_array == target_array).all(), referred_array
        else:
            return_address_loc = utils.relative_range_address_loc(return_value_address, 
                                                                  referred_cell_loc)
            return_array = _extract_array(referred_value, return_address_loc)
            return (referred_array == target_array).all(), return_array
    except Exception as e:
        print(e)
        print(address)
        return False, None

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
        check_tuple = _is_same(address, referred_cell_loc, target_value, referred_value)
        if not check_tuple[0]:
            if how == 'copy':
                _write(sheet_name, target_ws, address, check_tuple[1])
            elif how == 'check':
                _make_red(target_ws, address)
            else:
                raise ValueError('howにはcopyまたはcheckを指定してください。')
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
    return

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

def compare_and_change_other_cell_value(target_ws, referred_ws, 
                                        return_address_dict: Dict[str, str]) -> None:
    """概要
    2つのワークシートを比較し、ある範囲において値が異なるか否かに応じて、別の範囲の値を更新する。

    Parameters
    ----------
    target_ws
        値を更新するワークシート。

    referred_ws
        値を参照するワークシート。

    return_address_dict: Dict[str, str]
        差分を比較する番地と、比較した結果に応じて値を更新する範囲の対応を示すDict[str, str]型。

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

    for address in return_address_dict.keys():
        check_tuple = _is_same(address, referred_cell_loc, target_value, referred_value,
                               return_address_dict[address])
        # 書き込み処理をまとめることで少し処理時間を短縮できるが、ここでは手抜き
        if check_tuple[0] and check_tuple[1][0][0] != ChangeFlag.NOT_CHANGED.value:
            target_ws.Range(return_address_dict[address]).Value = ChangeFlag.NOT_CHANGED.value
        elif not check_tuple[0] and check_tuple[1][0][0] != ChangeFlag.CHANGED.value:
            target_ws.Range(return_address_dict[address]).Value = ChangeFlag.CHANGED.value
    return