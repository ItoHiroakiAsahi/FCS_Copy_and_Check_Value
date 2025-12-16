"""
プロジェクト計画書（計画変更届）に含まれるシートのうち、
情報記入シートの情報を比較する関数を定義する。
"""
from typing import Union, Dict, List, Tuple
import pandas as pd
from constants import KeikakuSheet
import settings
import utils

INFO_SHEET_LIST = [
    KeikakuSheet.IKUSEI_INFO,
    KeikakuSheet.TENNEN_INFO,
    KeikakuSheet.IN_PJ_EMISSION_INFO,
    KeikakuSheet.OUT_PJ_INFO,
]
_FOREST_NAME_COL_NAME = 'forest_name'

def _get_max_row_from_ws(ws1, ws2) -> int:
    """概要
    2つのワークシートを受け取り、記載されている中で最も下部の行番号を返す。

    Parameters
    ----------
    ws1, ws2
        参照するワークシート。

    Returns
    ----------
    row_num: int
        ws1, ws2に記載されている中で最も下部に存在するセルの行番号。
    """
    ws1_bottom_row = utils.from_cell_address_to_column_row_int(\
        ws1.UsedRange.Address.split(':')[1])[1]
    ws2_bottom_row = utils.from_cell_address_to_column_row_int(\
        ws2.UsedRange.Address.split(':')[1])[1]
    return max(ws1_bottom_row, ws2_bottom_row)
    
def _forest_name(forest: Union[tuple, pd.Series]) -> str:
    """概要
    特定の林地情報を格納するtuple型またはpd.Series型から、林地名を表すstr型を返す。

    Parameters
    ----------
    forest: Union[tuple, pd.Series]
        特定の林地情報を格納するtuple型またはpd.Series型。

    Returns
    ----------
    forest_name: str
        林地名を表すstr型。
    """
    forest_name = ''
    for i in range(0, len(forest)):
        if isinstance(forest,tuple):
            val = forest[i]
        elif isinstance(forest, pd.Series):
            val = forest.iat[i]
        else:
            raise TypeError('_forest_name()はtupleまたはpd.Seriesのみを引数に受け入れます。')
        if pd.isna(val):
            break
        else:
            if forest_name != '':
                forest_name += '-'
            if isinstance(val, float):
                val = int(val)
            forest_name += str(val)
    return forest_name

def _add_forest_name(df: pd.DataFrame, forest_name_col_list: List[str]) -> pd.DataFrame:
    """概要
    林地情報を格納するdataframe型に対して、林地名を格納した新たな列を追加する。

    Parameters
    ----------
    df: pd.DataFrame
        林地情報を格納するdataframe型。

    forest_name_col_list: List[str]
        林地名に関連する情報を持つ列名を格納したlist型。

    Returns
    ----------
    df: pd.DataFrame
        林地情報を格納するdataframe型に対して、林地名を格納した新たな列を追加したdataframe型。
    """
    col_num_list = [utils.from_alpha_to_num(c) - 1 for c in forest_name_col_list]
    df[_FOREST_NAME_COL_NAME] = df.iloc[:, col_num_list].apply(
        _forest_name, axis=1)
    return df

def _forest_info_diff_score(target_forest_info: pd.Series, referred_forest_info: pd.Series,
                         compare_col_num_list: List[int]) -> int:
    """概要
    2つの林地情報を受け取った上で、両者の値が異なる番地の数を返す。

    Parameters
    ----------
    target_forest_info, referred_forest_info: pd.Series
        乖離度を調べる林地情報。

    compare_col_num_list: List[int]
        乖離度を計算する際に使用する列番号のリスト。

    Returns
    ----------
    score: int
        2つの林地の乖離度を示すint型。compare_col_num_listで指定されたすべての列情報が
        同一ならば0、同一でない要素が1つ増えるごとに+1される。
    """
    target_forest_info = target_forest_info.fillna('nan')
    referred_forest_info = referred_forest_info.fillna('nan')
    s = 0
    for col_num in compare_col_num_list:
        try:
            if target_forest_info.iloc[col_num] != referred_forest_info.iloc[col_num]:
                s += 1
        except Exception as e:
            print(e, col_num, target_forest_info)
            raise e
    return s

def _forest_df_diff_dict(target_df: pd.DataFrame, referred_df: pd.DataFrame,
                         compare_col_num_list: List[int]) -> Dict[str, Dict[str, int]]:
    """概要
    林地情報を格納するDataFrame型を2つ受け取り、個々の林地情報同士をすべての組み合わせで比較し、
    差分がある情報の位置を格納したリスト型をvalueに持つdict型を返す。

    Parameters
    ----------
    target_df: pd.DataFrame
        赤字で記載するworksheetの情報を格納するnp.DataFrame型。

    referred_df: pd.DataFrame
        差分を検出するためのworksheetの情報を格納するpd.DataFrame型。

    Returns
    ----------
    d: Dict[str, Dict[str, List[int]]]
        すべての林地情報同士の組み合わせにおいて、差分が認められる位置を格納したリスト型を
        valueに持つdict型。
    """
    d = {}
    # 抽出したすべての林地の組み合わせにおいて、値の異なるセル番地を抽出し、辞書型に保存
    for i in range(0, len(target_df)):
        t_ser = target_df.iloc[i]
        dd = {}
        for j in range(0, len(referred_df)):
            r_ser = referred_df.iloc[j]
            dd[r_ser.name] = _forest_info_diff_score(t_ser, r_ser, compare_col_num_list)
        d[t_ser.name] = dd
    return d

def _min_val_info_from_dict(d: Dict[str, Dict[str, List[int]]]) -> Tuple[int, int]:
    """概要
    2つの林地情報の差分情報を格納した辞書型を受け取り、差分が最も少ない組み合わせのindexを返す。

    Parameters
    ----------
    d: Dict[str, Dict[str, List[int]]]
        すべての林地情報同士の組み合わせにおいて、差分が認められる位置を格納したリスト型を
        valueに持つdict型。
        
    Returns
    ----------
    t_df_index: int
        差分が最も少ない組み合わせのうち、差分を赤字にするworksheetに紐づく林地情報の位置。

    r_df_index: int
        差分が最も少ない組み合わせのうち、参照されるworksheetに紐づく林地情報の位置。
    """
    min_d = {key: min(d[key].values()) for key in d.keys()}
    min_s = min(min_d.values())
    t_df_index = -1
    for key in min_d.keys():
        if min_d[key] == min_s:
            t_df_index = key
            break
    if t_df_index == -1:
        raise ValueError('条件に合うt_df_indexが存在しません。\n' \
        'd: {}, min_s: {}, t_df_index: {}, r_df_index: {}'.format(
            d, min_s, t_df_index, r_df_index))
    r_df_index = -1
    for key in d[t_df_index].keys():
        if d[t_df_index][key] == min_s:
            r_df_index = key
            break
    if r_df_index == -1:
        raise ValueError('条件に合うt_df_indexが存在しません。\n' \
        'd: {}, min_s: {}, t_df_index: {}, r_df_index: {}'.format(
            d, min_s, t_df_index, r_df_index))
    return t_df_index, r_df_index

def _diff_col_num_list(target_forest_info: pd.Series, referred_forest_info: pd.Series, 
                       check_col_num_list: List[int]) -> List[int]:
    """概要
    2つの林地情報を受け取った上で、両者の値が異なる列番号のリストを返す。

    Parameters
    ----------
    target_forest_info, referred_forest_info: pd.Series
        差分を比較する林地情報

    check_col_num_list: List[int]
        差分を確認する際に使用する列番号のリスト。

    Returns
    ----------
    l: List[int]
        差分のある列番号を格納するlist型。
    """
    target_forest_info = target_forest_info.fillna('nan')
    referred_forest_info = referred_forest_info.fillna('nan')
    l = []
    for col_num in check_col_num_list:
        try:
            if target_forest_info.iat[col_num] != referred_forest_info.iat[col_num]:
                l.append(col_num)
        except Exception as e:
            print(e, col_num, target_forest_info)
            raise e
    return l

def _check_cell_address_list(target_ws, referred_ws, check_col_list: List[str], 
                             forest_name_col_list: List[str], compare_col_list: List[str], 
                             col_offset: int, row_offset: int) -> List[str]:
    """概要
    2つの情報記入シートを比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する情報記入シート。
    
    referred_ws:
        差分を確認するための情報記入シート。

    check_col_list: List[str]
        差分を確認する列の列名を格納するlist型。

    forest_name_col_list: List[str]
        林地名を格納する列の列名を格納するlist型。

    compare_col_list: List[str]
        2つの林地情報の乖離度を調べるにあたって参照する列の列名を格納するlist型。

    col_offset: int
        列方向のオフセット数を示すint型。

    row_offset: int
        行方向のオフセット数を示すint型。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するリスト型。
    """
    check_cell_address_list = []            
    range_address = '{}{}:{}{}'.format(\
        check_col_list[0], row_offset + 1, check_col_list[-1], 
        _get_max_row_from_ws(target_ws, referred_ws))
    # 処理が複雑なためDataFrame型に変換
    target_df = pd.DataFrame(target_ws.Range(range_address).Value).dropna(how='all')
    referred_df = pd.DataFrame(referred_ws.Range(range_address).Value).dropna(how='all')
    # dataframe型に林地名の列を追加
    target_df = _add_forest_name(target_df, forest_name_col_list)
    referred_df = _add_forest_name(referred_df, forest_name_col_list)
    compare_col_num_list = [utils.from_alpha_to_num(alpha) - col_offset - 1 
                            for alpha in compare_col_list]
    check_col_num_list = [utils.from_alpha_to_num(alpha) - col_offset - 1 
                          for alpha in check_col_list]
    # 林地ごとに林地名を抽出して処理（混交林に対応）
    for forest_name in target_df[_FOREST_NAME_COL_NAME].unique():
        t_df = target_df[target_df[_FOREST_NAME_COL_NAME] == forest_name]
        r_df = referred_df[referred_df[_FOREST_NAME_COL_NAME] == forest_name]
        # 抽出したすべての林地の組み合わせにおいて、値の異なるセル番地を抽出し、辞書型に保存
        d = _forest_df_diff_dict(t_df, r_df, compare_col_num_list)
        while(True):
            if len(d) == 0 or len(list(d.values())[0]) == 0:
                break
            # 乖離度が最も小さい組み合わせを抽出
            # （値の異なるセルが最小のもの、最小の組み合わせが複数ある場合は、先頭のもの同士を採用）
            t_df_index, r_df_index = _min_val_info_from_dict(d)
            
            # 乖離度の小さかった組み合わせにおいて、差分を赤字にするセルのリストとして格納
            diff_col_num_list = _diff_col_num_list(
                t_df.loc[t_df_index], r_df.loc[r_df_index], check_col_num_list)
            check_cell_address_list += ['{}{}'.format(utils.toAlpha3(col_num + col_offset + 1), 
                                                      t_df_index + row_offset + 1) 
                                        for col_num in diff_col_num_list]
            
            # 辞書型から抽出した組み合わせの情報を削除
            del d[t_df_index]
            if len(d) == 0:
                break
            for key in d.keys():
                del d[key][r_df_index]
            
         # 林地が追加されていた場合はすべての情報を赤字で表示
        if len(d) != 0 and len(list(d.values())[0]) == 0:
            for t_df_index in d.keys():
                check_cell_address_list.append('{c1}{r}:{c2}{r}'.format(
                    r = t_df_index + row_offset + 1, c1 = check_col_list[0], c2 = check_col_list[-1]))
    return check_cell_address_list

def check_cell_address_list_ikusei_info(target_ws, referred_ws) -> List[str]:
    """概要
    2つの【吸収量（育成林）算定用】情報記入シート（001、003共通）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する【吸収量（育成林）算定用】情報記入シート（001、003共通）。
    
    referred_ws:
        差分を確認するための【吸収量（育成林）算定用】情報記入シート（001、003共通）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するリスト型。
    """
    return _check_cell_address_list(target_ws, referred_ws,
                                    settings.IKUSEI_INFO_PARAMS.CHECK_COL_LIST,
                                    settings.IKUSEI_INFO_PARAMS.FOREST_NAME_COL_LIST,
                                    settings.IKUSEI_INFO_PARAMS.COMPARE_COL_LIST,
                                    settings.IKUSEI_INFO_PARAMS.COL_OFFSET,
                                    settings.IKUSEI_INFO_PARAMS.ROW_OFFSET)

def check_cell_address_list_tennen_info(target_ws, referred_ws) -> List[str]:
    """概要
    2つの【吸収量（天然生林）算定用】情報記入シート（FO-001）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する【吸収量（天然生林）算定用】情報記入シート（FO-001）。
    
    referred_ws:
        差分を確認するための【吸収量（天然生林）算定用】情報記入シート（FO-001）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するリスト型。
    """
    return _check_cell_address_list(target_ws, referred_ws,
                                    settings.TENNEN_INFO_PARAMS.CHECK_COL_LIST,
                                    settings.TENNEN_INFO_PARAMS.FOREST_NAME_COL_LIST,
                                    settings.TENNEN_INFO_PARAMS.COMPARE_COL_LIST,
                                    settings.TENNEN_INFO_PARAMS.COL_OFFSET,
                                    settings.TENNEN_INFO_PARAMS.ROW_OFFSET)

def check_cell_address_list_in_pj_emission_info(target_ws, referred_ws) -> List[str]:
    """概要
    2つの【排出量（PJ内）算定用】情報記入シート（001、003共通）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する【排出量（PJ内）算定用】情報記入シート（001、003共通）。
    
    referred_ws:
        差分を確認するための【排出量（PJ内）算定用】情報記入シート（001、003共通）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するリスト型。
    """
    return _check_cell_address_list(target_ws, referred_ws,
                                    settings.IN_PJ_EMISSION_INFO_PARAMS.CHECK_COL_LIST,
                                    settings.IN_PJ_EMISSION_INFO_PARAMS.FOREST_NAME_COL_LIST,
                                    settings.IN_PJ_EMISSION_INFO_PARAMS.COMPARE_COL_LIST,
                                    settings.IN_PJ_EMISSION_INFO_PARAMS.COL_OFFSET,
                                    settings.IN_PJ_EMISSION_INFO_PARAMS.ROW_OFFSET)

def check_cell_address_list_out_pj_info(target_ws, referred_ws) -> List[str]:
    """概要
    2つの【主伐再造林（PJ外）算定用】情報記入シート（FO-001）を比較し、林地名をもとに林地情報を紐づける。
    両者に差があった場合、target_wsのセル番地をリストに格納して返す。

    Parameters
    ----------
    target_ws:
        差分を赤字に変更する【主伐再造林（PJ外）算定用】情報記入シート（FO-001）。
    
    referred_ws:
        差分を確認するための【主伐再造林（PJ外）算定用】情報記入シート（FO-001）。

    Returns:
    ----------
    check_cell_address_list: List[str]
        差分のあるtarget_wsのセル番地を格納するリスト型。
    """
    return _check_cell_address_list(target_ws, referred_ws,
                                    settings.OUT_PJ_INFO_PARAMS.CHECK_COL_LIST,
                                    settings.OUT_PJ_INFO_PARAMS.FOREST_NAME_COL_LIST,
                                    settings.OUT_PJ_INFO_PARAMS.COMPARE_COL_LIST,
                                    settings.OUT_PJ_INFO_PARAMS.COL_OFFFSET,
                                    settings.OUT_PJ_INFO_PARAMS.ROW_OFFSET)