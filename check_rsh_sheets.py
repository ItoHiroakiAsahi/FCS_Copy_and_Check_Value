"""
プロジェクト計画書（計画変更届）に含まれるシートのうち、
幹材積量算定シートの情報を比較する関数を定義する。
"""
from typing import List
from constants import KeikakuSheet
import settings
import utils

RSH_SHEET_LIST = [
    KeikakuSheet.IKUSEI_RSH,
    KeikakuSheet.TENNEN_RSH
]

def _species_rank_list(value: tuple, col_interval: int) -> List[str]:
    """概要
    幹材積量算定シートを受け取り、樹種＋地位名のstr型を格納するlist型を返す。
    
    Parameters
    ----------
    value: tuple
        シート全体の情報が記載されたtuple型。

    Returns
    ----------
    l: list
        樹種＋地位名のstr型を格納するlist型。
    """
    t = value[0]
    return [t[i] for i in range(len(t)) if i % col_interval == 0]

def _check_cell_address_list_rsh(target_ws, referred_ws, col_offset: int, row_offset: int,
                             col_interval: int, species_rank_ref_cell_address: str) -> List[str]:
    """概要
    幹材積量算定シートの情報を比較し、差分のあるセルのアドレスをリストに格納する。

    Parameters
    ----------
    target_ws
        差分を赤字にする幹材積量算定シート。

    referred_ws
        差分を参照する幹材積量算定シート。

    col_offset: int
        列方向のオフセット数を示すint型。

    row_offset: int
        行方向のオフセット数を示すint型。

    col_interval: int
        行方向のパターンの間隔を示すint型。

    species_rank_ref_cell_address: str
        樹種＋地位名を記入する範囲のうち左端のセルの番地を示すstr型。

    Returns
    ----------
    check_cell_address_list: List[str]
        差分のあるセルの番地を示すstr型を格納するlist型。
    """
    check_cell_address_list = []
    bottom_right_cell_address = utils.get_cell_address_from_range_address(
        utils.get_max_range(target_ws.UsedRange.Address, referred_ws.UsedRange.Address),
        loc = 'bottom_right'
    )
    range_address = '{}:{}'.format(species_rank_ref_cell_address, bottom_right_cell_address)
    max_age = utils.from_cell_address_to_column_row_int(bottom_right_cell_address)[1] - row_offset
    target_value = target_ws.Range(range_address).Value
    referred_value = referred_ws.Range(range_address).value
    target_species_rank_list = _species_rank_list(target_value, col_interval)
    referred_species_rank_list = _species_rank_list(referred_value, col_interval)
    for species_rank in target_species_rank_list:
        t_index = target_species_rank_list.index(species_rank)
        t_col = t_index * col_interval
        if species_rank in referred_species_rank_list:        
            r_index = referred_species_rank_list.index(species_rank)
            r_col = r_index * col_interval
            for age in range(1, max_age + 1):
                if target_value[age + 2][t_col] != referred_value[age + 2][r_col]:
                    check_cell_address_list.append('{}{}'.format(
                        utils.toAlpha3(t_col + col_offset + 1),
                        age + row_offset + 1))
        else:
            check_cell_address_list.append('{}{}'.format(
                utils.toAlpha3(t_col + col_offset + 1),
                row_offset - 2
            ))
            for age in range(1, max_age + 1):
                check_cell_address_list.append('{}{}'.format(
                        utils.toAlpha3(t_col + col_offset + 1),
                        age + row_offset))
    return check_cell_address_list

def check_cell_address_list_ikusei_rsh(target_ws, referred_ws) -> List[str]:
    """概要
    幹材積量算定シート_育成林および主伐用（001、003共通）同士を比較し、差分のあるセルを赤字に更新する。

    Parameters
    ----------
    target_ws
        差分を赤字にする幹材積量算定シート。

    referred_ws
        差分を参照する幹材積量算定シート。

    Returns
    ----------
    check_cell_address_list: List[str]
        差分のあるセルの番地を示すstr型を格納するlist型。
    """
    return _check_cell_address_list_rsh(target_ws, referred_ws, settings.IKUSEI_RSH_PARAMS.COL_OFFSET,
                                        settings.IKUSEI_RSH_PARAMS.ROW_OFFSET,
                                        settings.IKUSEI_RSH_PARAMS.COL_INTERVAL,
                                        settings.IKUSEI_RSH_PARAMS.SPECIES_RANK_REF_CELL_ADDRESS)

def check_cell_address_list_tennen_rsh(target_ws, referred_ws) -> List[str]:
    """概要
    幹材積量算定シート_天然生林（FO-001）同士を比較し、差分のあるセルを赤字に更新する。

    Parameters
    ----------
    target_ws
        差分を赤字にする幹材積量算定シート。

    referred_ws
        差分を参照する幹材積量算定シート。

    Returns
    ----------
    check_cell_address_list: List[str]
        差分のあるセルの番地を示すstr型を格納するlist型。
    """
    return _check_cell_address_list_rsh(target_ws, referred_ws, settings.TENNEN_RSH_PARAMS.COL_OFFSET,
                                        settings.TENNEN_RSH_PARAMS.ROW_OFFSET,
                                        settings.TENNEN_RSH_PARAMS.COL_INTERVAL,
                                        settings.TENNEN_RSH_PARAMS.SPECIES_RANK_REF_CELL_ADDRESS)