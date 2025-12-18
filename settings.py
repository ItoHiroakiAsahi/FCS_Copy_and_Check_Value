"""
コピーや差分比較時に参照するセル番地、範囲の情報を定義する。
"""
from typing import List, Dict
from constants import KeikakuSheet
import utils

class REGISTER_APPLICATION_PARAMS:
    CHANGE_OTHER_CELL_VALUE_DICT = {'E26': 'P26', 'E27': 'P27', 'E28': 'P28', 'E29': 'P29', 
                                    'E30': 'P30', 'E31': 'P31', 'E33': 'P33', 'E34': 'P34',
                                    'E35': 'P35', 'E36': 'P36', 'E37': 'P37'}
    COPY_ADDRESS_LIST = ['E6:E7', 'G10', 'E11:E14', 'E16:I16', 'E17:E21', 'H18:P18', 'I23:I24',
                          'E26:E31', 'P26:P31', 'E33:E37', 'P33:P37', 'E42:E43', 'G45', 'E47', 'E50',
                          'E55', 'F56', 'H56', 'J56', 'E60', 'F61', 'H61', 'J61']
    NUM_TO_STR_ADDRESS_LIST = ['G45']

class OTHER_PARTICIPANTS_PARAMS:
    ROW_INTERVAL = 21
    ROW_PATTERN_NUM = 10
    CHANGE_OTHER_CELL_VALUE_DICT_PATTERN = {'B9': 'M9', 'B10': 'M10', 'B11': 'M11', 'B12': 'M12',
                                            'B13': 'M13', 'B14': 'M14'}
    CHECK_CELLS_PATTERN = ['D4', 'B5', 'B6', 'B8', 'D8', 'F8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14',
                           'M9', 'M10', 'M11', 'M12', 'M13', 'M14', 'B15', 'B17', 'B18', 'B19', 'E16:M16',
                           'G21', 'G22']
    COPY_CELLS_PATTERN = ['D4', 'B5:B6', 'B8:F8', 'B9:B14', 'M9:M14', 'B15:B19', 'E16:M16', 'G21:G22']

class CONFIRMATION_OF_CONDITION_5_PARAMS:
    CHECK_ADDRESS_LIST = ['I4', 'K4', 'M4']
    COPY_ADDRESS_LIST = ['I4', 'K4', 'M4']

class CONFIRMATION_OF_CONDITION_6_PARAMS:
    CHECK_ADDRESS_LIST = ['I4', 'K4', 'M4']
    COPY_ADDRESS_LIST = ['I4', 'K4', 'M4']

class PLAN_CHANGE_NORTIFICATION_PARAMS:
    CHECK_ADDRESS_LIST = ['E5', 'F6', 'H6', 'J6', 'E7', 'E8', 'E9', 'E10', 'E21', 'E22',
                          'E29', 'I29', 'K29', 'M29', 'E30', 'E31', 'E32', 'E33']
    COPY_ADDRESS_LIST = ['E5', 'F6', 'H6', 'J6', 'E7:E10', 'E21:E22',
                         'E29:E33', 'I29', 'K29', 'M29']
    CHECK_TEXT_CELL_LIST = ['E10', 'E33']

class OVERVIEW_PARAMS:
    CHECK_ADDRESS_LIST = ['C8', 'E17', 'C18', 'C19', 'C21', 'E21', 'G21', 'A22', 'C26', 'G26']
    COPY_ADDRESS_LIST = ['C8', 'E17', 'C18', 'C19', 'C21', 'E21', 'G21', 'A22']

class METHODOLOGY_FO001_PARAMS:
    CHECK_ADDRESS_LIST = ['I3', 'D5', 'D7', 'D10', 'D11', 'D12', 'G13', 'G14', 'G15', 'G16', 'G17', 'G18',
                          'D19', 'D20', 'D21', 'I21','A24', 'A29', 'A33', 'A39', 'A43', 'A47', 'A48', 
                          'A52', 'A53', 'A54', 'B56', 'B57', 'I57', 'B82', 'B83', 'B84', 'B85',
                          'D86', 'D88', 'G87', 'B89']
    COPY_ADDRESS_LIST = ['I3', 'D5', 'D7', 'D10', 'A24', 'A29', 'A33', 'A39', 'A43', 'A47:A48', 
                         'I57', 'B82:B85', 'D86:D88', 'G87', 'B89']
    CHECK_TEXT_CELL_LIST = ['D5', 'D10', 'A24', 'A29', 'A33', 'A39', 'A43', 'B82', 'B83', 'B84', 'B85', 'B89']

class SKK_CHANGES_PARAMS:
    ROW_INTERVAL = 11
    COL_INTERVAL = 3
    ROW_PATTERN_NUM = 10
    COL_PATTERN_NUM = 10
    OUT_OF_PATTERN_CELL_LIST = ['C3']
    CHECK_CELLS_PATTERN = ['C8', 'C9', 'C10', 'C11', 'C12', 'C13', 'C14']
    COPY_CELLS_PATTERN = ['C8:C14']

class MULTIPLE_SKK_INFO_PARAMS:
    COL_INTERVAL = 14
    COL_PATTERN_NUM = 20
    CHECK_CELLS_PATTERN = ['A10', 'A14', 'A19']
    COPY_CELLS_PATTERN = ['A10', 'A14', 'A19']
    CHECK_TEXT_CELL_PATTERN = ['A10', 'A14', 'A19']

class DATA_MANAGEMENT_PARAMS:
    CHECK_ADDRESS_LIST = ['E4', 'E5', 'E9', 'M13', 'M14']
    COPY_ADDRESS_LIST = ['E4:E5', 'E9', 'M13:M14']
    CHECK_TEXT_CELL_LIST = ['E9']

class SPECIAL_NOTES_PARAMS:
    CHECK_ADDRESS_LIST = ['F3', 'F4', 'A6', 'F10', 'F11', 'D13', 'D14', 'I14', 
                          'D17', 'D18', 'E21', 'E23']
    COPY_ADDRESS_LIST = ['F3:F4', 'A6', 'F10:F11', 'D13', 'D14', 'I14', 'D17:D18', 'E21', 'E23']
    CHECK_TEXT_CELL_LIST = ['A6', 'E23']

class MONITORING_PLAN_FO001_PARAMS:
    CHECK_ADDRESS_LIST = utils.from_range_address_list_to_each_cell_adress_list(['K4:AQ53'])
    COPY_ADDRESS_LIST = ['K4:AQ53']
    CHECK_TEXT_CELL_LIST = utils.from_range_address_list_to_each_cell_adress_list(['O4:AQ53'])

class IKUSEI_RSH_PARAMS:
    COL_INTERVAL = 4
    ROW_OFFSET = 14
    COL_OFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['F8', 'F9']
    SPECIES_RANK_REF_CELL_ADDRESS = 'B12'
    VALUE_REF_CELL_ADDRESS = 'B15'

class IKUSEI_INFO_PARAMS:
    ROW_OFFSET = 9
    COL_OFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['N5', 'Q5']
    CHECK_COL_LIST = [utils.toAlpha3(i) for i in range(2, 32)]
    FOREST_NAME_COL_LIST = [utils.toAlpha3(i) for i in range(4, 14)]
    COMPARE_COL_LIST = ['O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'AD']

class IKUSEI_CALCULATION_PARAMS:
    ROW_OFFSET = 45
    CHECK_COL_LIST = ['AJ']

class TENNEN_RSH_PARAMS:
    COL_INTERVAL = 4
    ROW_OFFSET = 13
    COL_OFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['F7', 'F8']
    SPECIES_RANK_REF_CELL_ADDRESS = 'B11'
    VALUE_REF_CELL_ADDRESS = 'B14'

class TENNEN_INFO_PARAMS:
    ROW_OFFSET = 7
    COL_OFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['L3', 'O3']
    CHECK_COL_LIST = [utils.toAlpha3(i) for i in range(2, 21)]
    FOREST_NAME_COL_LIST = [utils.toAlpha3(i) for i in range(2, 12)]
    COMPARE_COL_LIST = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']

class TENNEN_CALCULATION_PARAMS:
    ROW_OFFSET = 45
    CHECK_COL_LIST = ['AA']

class IN_PJ_EMISSION_INFO_PARAMS:
    ROW_OFFSET = 10
    COL_OFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['O6', 'Q6']
    CHECK_COL_LIST = [utils.toAlpha3(i) for i in range(2, 23)]
    FOREST_NAME_COL_LIST = [utils.toAlpha3(i) for i in range(4, 14)]
    COMPARE_COL_LIST = ['O', 'P', 'Q', 'R', 'S', 'T', 'U']

class IN_PJ_HWP_PARAMS:
    FIRST_ROW_NUM = 48
    LAST_ROW_NUM = 161
    ROW_INTERVAL = 3
    SPECIES_CHECK_COLS = ['C', 'D', 'E', 'H', 'J', 'M', 'O']
    FIRST_CELL_CHECK_COLS = ['I', 'K', 'L', 'N', 'P', 'Q', 'W', 'X', 'Y', 'Z',
                             'AB', 'AC', 'AD', 'AE']
    YEAR_CHECK_COLS = ['R', 'S', 'V', 'AA', 'AF', 'AG', 'AH', 'AK', 'AL', 'AO', 
                       'AQ', 'AP', 'AS', 'AT', 'AV', 'AW', 'AX', 'AZ', 'BA',
                       'BC', 'BE', 'BF', 'BG', 'BH', 'BJ', 'BK', 'BL']
    
class OUT_PJ_INFO_PARAMS:
    ROW_OFFSET = 10
    COL_OFFFSET = 1
    OUT_OF_PATTERN_CELL_LIST = ['P7', 'R7']
    CHECK_COL_LIST = [utils.toAlpha3(i) for i in range(2, 40)]
    FOREST_NAME_COL_LIST = [utils.toAlpha3(i) for i in range(4, 14)]
    COMPARE_COL_LIST = ['O', 'P', 'Q', 'R', 'S', 'U', 'AH', 'AI', 'AJ', 'AK', 'AM']

def _OTHER_PARTICIPANTS_CHANGE_OTHER_CELL_VALUE_DICT() -> Dict[str, str]:
    """概要
    代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の変更の有無を確認するセルと
    値の更新をするセルの組み合わせを指定するdict型を作成。

    Parameters
    ----------
    None

    Returns
    ----------
    d: Dict[str, str]
        代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の変更の有無を確認するセルと
        値の更新をするセルの組み合わせを指定するdict型
    """
    d = {}
    for i in range(0, OTHER_PARTICIPANTS_PARAMS.ROW_PATTERN_NUM):
        for range_address in OTHER_PARTICIPANTS_PARAMS.CHANGE_OTHER_CELL_VALUE_DICT_PATTERN.keys():
            pre_address = utils.move_range_address(range_address, \
                                                   down = i * OTHER_PARTICIPANTS_PARAMS.ROW_INTERVAL)
            post_address = utils.move_range_address(OTHER_PARTICIPANTS_PARAMS.\
                                                    CHANGE_OTHER_CELL_VALUE_DICT_PATTERN[range_address],
                                                    down = i * OTHER_PARTICIPANTS_PARAMS.ROW_INTERVAL)
            d[pre_address] = post_address
    return d
    
def _OTHER_PARTICIPANTS_CHECK_LIST() -> List[str]:
    """概要
    代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の値の比較をするセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の値の比較をするセルを指定するlist型。
    """
    l = []
    for i in range(0, OTHER_PARTICIPANTS_PARAMS.ROW_PATTERN_NUM):
        for range_address in OTHER_PARTICIPANTS_PARAMS.CHECK_CELLS_PATTERN:
            l.append(utils.move_range_address(range_address, \
                down = i * OTHER_PARTICIPANTS_PARAMS.ROW_INTERVAL))
    return l

def _OTHER_PARTICIPANTS_COPY_LIST() -> List[str]:
    """概要
    代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の書き込みセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        代表以外のプロジェクト実施者もしくはプログラム型運営・管理者の書き込みセルを指定するlist型。
    """
    l = []
    for i in range(0, OTHER_PARTICIPANTS_PARAMS.ROW_PATTERN_NUM):
        for range_address in OTHER_PARTICIPANTS_PARAMS.COPY_CELLS_PATTERN:
            l.append(utils.move_range_address(range_address, \
                down = i * OTHER_PARTICIPANTS_PARAMS.ROW_INTERVAL))
    return l

def _SKK_CHANGES_CHECK_LIST() -> List[str]:
    """概要
    森林経営計画の適用条件１への適用と計画の変遷（FO-001）の値の比較をするセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        森林経営計画の適用条件１への適用と計画の変遷（FO-001）の値の比較をするセルを指定するlist型。
    """
    l = SKK_CHANGES_PARAMS.OUT_OF_PATTERN_CELL_LIST
    for r in range(0, SKK_CHANGES_PARAMS.ROW_PATTERN_NUM):
        for c in range(0, SKK_CHANGES_PARAMS.COL_PATTERN_NUM):
            right = c * SKK_CHANGES_PARAMS.COL_INTERVAL
            down = r * SKK_CHANGES_PARAMS.ROW_INTERVAL
            # パターンの列、行間隔が変わるので、個別対応
            if c >= 5:
                right += 1
            if r >= 1:
                down -= 1
            for cell in SKK_CHANGES_PARAMS.CHECK_CELLS_PATTERN:
                l.append(utils.move_range_address(cell, right=right, down=down))
    return l

def _SKK_CHANGES_COPY_LIST() -> List[str]:
    """概要
    森林経営計画の適用条件１への適用と計画の変遷（FO-001）の書き込みセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        森林経営計画の適用条件１への適用と計画の変遷（FO-001）の書き込みセルを指定するlist型。
    """
    l = SKK_CHANGES_PARAMS.OUT_OF_PATTERN_CELL_LIST
    for r in range(0, SKK_CHANGES_PARAMS.ROW_PATTERN_NUM):
        for c in range(0, SKK_CHANGES_PARAMS.COL_PATTERN_NUM):
            right = c * SKK_CHANGES_PARAMS.COL_INTERVAL
            down = r * SKK_CHANGES_PARAMS.ROW_INTERVAL
            # パターンの列、行間隔が変わるので、個別対応
            if c >= 5:
                right += 1
            if r >= 1:
                down -= 1
            for cell in SKK_CHANGES_PARAMS.COPY_CELLS_PATTERN:
                l.append(utils.move_range_address(cell, right=right, down=down))
    return l

def _MULTIPLE_SKK_INFO_CHECK_LIST() -> List[str]:
    """概要
    2.2 複数森林経営計画用（FO-001）の値の比較をするセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        2.2 複数森林経営計画用（FO-001）の値の比較をするセルを指定するlist型。
    """
    l = []
    for i in range(0, MULTIPLE_SKK_INFO_PARAMS.COL_PATTERN_NUM):
        for cell in MULTIPLE_SKK_INFO_PARAMS.CHECK_CELLS_PATTERN:
            l.append(utils.move_range_address(cell, \
                right = i * MULTIPLE_SKK_INFO_PARAMS.COL_INTERVAL))
    return l

def _MULTIPLE_SKK_INFO_COPY_LIST() -> List[str]:
    """概要
    2.2 複数森林経営計画用（FO-001）の書き込みセルを指定するlist型を作成。
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        2.2 複数森林経営計画用（FO-001）の書き込みセルを指定するlist型。
    """
    l = []
    for i in range(0, MULTIPLE_SKK_INFO_PARAMS.COL_PATTERN_NUM):
        for cell in MULTIPLE_SKK_INFO_PARAMS.COPY_CELLS_PATTERN:
            l.append(utils.move_range_address(cell, \
                right = i * MULTIPLE_SKK_INFO_PARAMS.COL_INTERVAL))
    return l

def _MULTIPLE_SKK_INFO_CHECK_TEXT_LIST() -> List[str]:
    """概要
    
    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        2.2 複数森林経営計画用（FO-001）の文字列の比較をするセルを指定するlist型。
    """
    l = []
    for i in range(0, MULTIPLE_SKK_INFO_PARAMS.COL_PATTERN_NUM):
        for cell in MULTIPLE_SKK_INFO_PARAMS.CHECK_TEXT_CELL_PATTERN:
            l.append(utils.move_range_address(cell, \
                right = i * MULTIPLE_SKK_INFO_PARAMS.COL_INTERVAL))
    return l

def _IN_PJ_HWP_CHECK_LIST() -> List[str]:
    """概要
    【吸収量（PJ内HWP）】情報記入・算定シート（FO-001）の値の比較をするセルを指定するlist型を作成。

    Parameters
    ----------
    None

    Returns
    ----------
    l: List[str]
        【吸収量（PJ内HWP）】情報記入・算定シート（FO-001）の値の比較をするセルを指定するlist型。
    """
    l = []
    for col in IN_PJ_HWP_PARAMS.SPECIES_CHECK_COLS:
        for row_num in range(IN_PJ_HWP_PARAMS.FIRST_ROW_NUM, \
                             IN_PJ_HWP_PARAMS.LAST_ROW_NUM + 1):
            l.append(col + str(row_num))
    for col in IN_PJ_HWP_PARAMS.FIRST_CELL_CHECK_COLS:
        l.append(col + str(IN_PJ_HWP_PARAMS.FIRST_ROW_NUM))
    for col in IN_PJ_HWP_PARAMS.YEAR_CHECK_COLS:
        for row_num in range(IN_PJ_HWP_PARAMS.FIRST_ROW_NUM, \
                             IN_PJ_HWP_PARAMS.LAST_ROW_NUM):
            if row_num % IN_PJ_HWP_PARAMS.ROW_INTERVAL == 0:
                l.append(col + str(row_num))
    return l

COPY_CELL_ADDRESS_DICT = {
    KeikakuSheet.REGISTER_APPLICATION: REGISTER_APPLICATION_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.OTHER_PARTICIPANTS: _OTHER_PARTICIPANTS_COPY_LIST(),
    KeikakuSheet.CONFIRMATION_OF_CONDITION_5: CONFIRMATION_OF_CONDITION_5_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.CONFIRMATION_OF_CONDITION_6: CONFIRMATION_OF_CONDITION_6_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.PLAN_CHANGE_NORTIFICATION: PLAN_CHANGE_NORTIFICATION_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.OVERVIEW: OVERVIEW_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.METHODOLOGY_FO001: METHODOLOGY_FO001_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.SKK_CHANGES: _SKK_CHANGES_COPY_LIST(),
    KeikakuSheet.MULTIPLE_SKK_INFO: _MULTIPLE_SKK_INFO_COPY_LIST(),
    KeikakuSheet.DATA_MANAGEMENT: DATA_MANAGEMENT_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.SPECIAL_NOTES: SPECIAL_NOTES_PARAMS.COPY_ADDRESS_LIST,
    KeikakuSheet.MONITORING_PLAN_FO001: MONITORING_PLAN_FO001_PARAMS.COPY_ADDRESS_LIST
}

# 文字列であることを明記して書き写す必要のあるセル
NUM_TO_STR_ADDRESS_DICT = {
    KeikakuSheet.REGISTER_APPLICATION: REGISTER_APPLICATION_PARAMS.NUM_TO_STR_ADDRESS_LIST
}

# 行の高さ、列の幅を変更できるシート
COPY_WIDTH_AND_HEIGHT_SHEET_LIST = [
    KeikakuSheet.PLAN_CHANGE_NORTIFICATION,
    KeikakuSheet.METHODOLOGY_FO001,
    KeikakuSheet.MULTIPLE_SKK_INFO,
    KeikakuSheet.DATA_MANAGEMENT,
    KeikakuSheet.MONITORING_PLAN_FO001
]

# 差分を確認して変更箇所を赤字にするセル番地の辞書形式
COMPARE_CELL_ADDRESS_DICT = {
    KeikakuSheet.OTHER_PARTICIPANTS: _OTHER_PARTICIPANTS_CHECK_LIST(),                                          
    KeikakuSheet.CONFIRMATION_OF_CONDITION_5: CONFIRMATION_OF_CONDITION_5_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.CONFIRMATION_OF_CONDITION_6: CONFIRMATION_OF_CONDITION_6_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.PLAN_CHANGE_NORTIFICATION: PLAN_CHANGE_NORTIFICATION_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.OVERVIEW: OVERVIEW_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.METHODOLOGY_FO001: METHODOLOGY_FO001_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.SKK_CHANGES: _SKK_CHANGES_CHECK_LIST(),
    KeikakuSheet.MULTIPLE_SKK_INFO: _MULTIPLE_SKK_INFO_CHECK_LIST(),
    KeikakuSheet.DATA_MANAGEMENT: DATA_MANAGEMENT_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.SPECIAL_NOTES: SPECIAL_NOTES_PARAMS.CHECK_ADDRESS_LIST,
    KeikakuSheet.MONITORING_PLAN_FO001: utils.from_range_address_list_to_each_cell_adress_list(['K4:AQ53']),
    KeikakuSheet.IN_PJ_HWP: _IN_PJ_HWP_CHECK_LIST()
}

# 差分がある場合、別のセルの値を変更するセル番地の辞書形式
COMPARE_AND_CHANGE_OTHER_CELL_VALUE_DICT = {
    KeikakuSheet.REGISTER_APPLICATION: REGISTER_APPLICATION_PARAMS.CHANGE_OTHER_CELL_VALUE_DICT,
    KeikakuSheet.OTHER_PARTICIPANTS: _OTHER_PARTICIPANTS_CHANGE_OTHER_CELL_VALUE_DICT()
}

# 差分を比較する情報記入シート
INFO_SHEET_LIST = [
    KeikakuSheet.IKUSEI_INFO,
    KeikakuSheet.TENNEN_INFO,
    KeikakuSheet.IN_PJ_EMISSION_INFO,
    KeikakuSheet.OUT_PJ_INFO,
]

# 差分を比較する幹材積量算定シート
RSH_SHEET_LIST = [
    KeikakuSheet.IKUSEI_RSH,
    KeikakuSheet.TENNEN_RSH
]

# 差分を比較する吸収量算定シート
CALC_SHEET_LIST = [
    KeikakuSheet.IKUSEI_CALCULATION,
    KeikakuSheet.TENNEN_CALCULATION
]

# 差分がある場合、文字列の値を検証するセル番地の辞書形式
CHECK_TEXT_CELL_DICT = {
    KeikakuSheet.PLAN_CHANGE_NORTIFICATION: PLAN_CHANGE_NORTIFICATION_PARAMS.CHECK_TEXT_CELL_LIST,
    KeikakuSheet.METHODOLOGY_FO001: METHODOLOGY_FO001_PARAMS.CHECK_TEXT_CELL_LIST,
    KeikakuSheet.MULTIPLE_SKK_INFO: _MULTIPLE_SKK_INFO_CHECK_TEXT_LIST(),
    KeikakuSheet.DATA_MANAGEMENT: DATA_MANAGEMENT_PARAMS.CHECK_TEXT_CELL_LIST,
    KeikakuSheet.SPECIAL_NOTES: SPECIAL_NOTES_PARAMS.CHECK_TEXT_CELL_LIST,
    KeikakuSheet.MONITORING_PLAN_FO001: MONITORING_PLAN_FO001_PARAMS.CHECK_TEXT_CELL_LIST
}