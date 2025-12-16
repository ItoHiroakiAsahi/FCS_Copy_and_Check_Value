"""
win32comを用いたエクセル操作に用いる便利ツールを定義する。
"""

from typing import List, Tuple
import warnings

ALPHABET = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 
            'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

def toAlpha3(num: int) -> str:
        """概要
        列番号をアルファベットの文字列に変換する。
        openpyxlのライブラリを使うのが一般的だが、そのためだけにimportするのが面倒なので自作。

        Parameters
        ----------
        num: int
            変換する列番号を表すint型。
        
        Returns
        ----------
        Alpha: str
            変換された列名を表すstr型。
        """
        h = int((num - 1 - 26) / (26 * 26))
        i = int((num - 1 - (h * 26 * 26)) / 26)
        j = num - (i * 26) - (h * 26 * 26)
        Alpha = ''
        for k in h, i, j:
            if k != 0:
                Alpha += chr(k + 64)
        return Alpha

def from_alpha_to_num(alpha: str) -> int:
    """概要
    列名を列番号に変換する。

    Parameters
    -----------
    alpha: str
        列名を表すstr型。

    Returns
    ----------
    num: int
        列番号を表すint型。
    """
    n = 0
    length = 0
    for s in reversed(alpha):
        if s in ALPHABET:
            n += (ord(s) - 64) * (26 ** length)
            length += 1
        else:
            raise ValueError('次の文字を列番号に変更できません。:{}'.format(alpha))
    return n

def from_cell_address_to_column_row_letter(address: str) -> Tuple[str, str]:
    """概要
    セル番号を表すstr型を列名、行名のstr型を格納したtuple型に変換する。
    例：A1→(A, 1)

    Parameters
    ----------
    address: str
        セル番号を表すstr型

    Returns
    ----------
    t: tuple
        ({列名}, {行名})からなるtuple型。
    """
    address = address.replace('$', '')
    column_letter = ''
    for s in address:
        if s in ALPHABET:
            column_letter += s
    row_letter = address.replace(column_letter, '')
    try:
        int(row_letter)
    except ValueError:
        raise ValueError('次のアドレスが不適です。:{}'.format(address))
    return (column_letter, row_letter)

def from_cell_address_to_column_row_int(address: str) -> Tuple[int, int]:
    """概要
    セル番号を表すstr型を列番号、行番号のint型を格納したtuple型に変換する。
    例：A1→(1, 1)

    Parameters
    ----------
    address: str
        セル番号を表すstr型

    Returns
    ----------
    t: tuple
        ({列番号}, {行番号})からなるtuple型。
    """
    address = address.replace('$', '')
    address_tuple = from_cell_address_to_column_row_letter(address)
    return (from_alpha_to_num(address_tuple[0]), int(address_tuple[1]))

def from_range_address_to_column_row_int(address: str) -> Tuple[Tuple[int]]:
    """概要
    セル範囲を表すstr型を左上のセル、右下のセルそれぞれの列番号、行番号のint型を格納したtuple型に変換する。
    例：A1:B2→((1, 1), (2, 2))

    Parameters
    ----------
    address: str
        セル範囲を表すstr型

    Returns
    ----------
    t: tuple
        (({左上のセルの列番号}, {左上のセルの行番号}), ({右下のセルの列番号}, {右下のセルの行番号}))
        からなるtuple型。
    """
    address = address.replace('$', '')
    if ':' in address:
        address_list = address.split(':')
        if len(address_list) != 2:
            raise ValueError('次のセル範囲に:が2つ以上含まれます。{}'.format(address))
        first_cell = address_list[0]
        last_cell = address_list[1]
    else:
        first_cell = address
        last_cell = address
    return (from_cell_address_to_column_row_int(first_cell), 
            from_cell_address_to_column_row_int(last_cell))

def from_column_row_int_to_cell_address(col_num: int, row_num: int) -> str:
    """"概要
    列番号、行番号を受け取り、セル番地を表すstr型を返す。

    Parameters
    ----------
    col_num: int
        セルの列番号を表すint型。

    row_num: int
        セルの行番号を表すint型。

    Returns
    -----------
    address: str
        セル番地を表すstr型。
    """
    return toAlpha3(col_num) + str(row_num)

def move_cell_address(address: str, right: int = 0, down: int = 0) -> str:
    address_tuple = from_cell_address_to_column_row_letter(address)
    col_letter = address_tuple[0]
    row_letter = address_tuple[1]
    if right != 0:
        col_num = from_alpha_to_num(col_letter) + right
        if col_num < 1:
            warnings.warn('A列より左の列が存在しないので、A列に変換しています。', stacklevel=2)
            col_num = 1
        col_letter = toAlpha3(col_num)
    if down != 0:
        row_num = int(row_letter) + down
        if row_num < 1:
            warnings.warn('1行より上の行が存在しないので、1行に変換しています。', stacklevel=2)
            row_num = 1
        row_letter = str(row_num)
    return col_letter + row_letter

def move_range_address(range: str, right: int = 0, down: int = 0) -> str:
    """概要
    範囲を指定したstr型を受け取り、それを移動した範囲を表すstr型を返す。

    Parameters
    ----------
    range: str
        範囲を指定するstr型。

    right: int
        右方向に移動するセル数。

    down: int
        下方向に移動するセル数。

    Returns
    ----------
    return_range: str
        rangeを右方向にright、下方向にdown移動させた範囲を示すstr型。
    """
    if ':' in range:
        range_list = range.split(':')
        if len(range_list) != 2:
            raise ValueError('次のセル番地が不適です。{}'.format(range))
        return '{}:{}'.format(move_cell_address(range_list[0], right, down),
                              move_cell_address(range_list[1], right, down))
    else:
        return move_cell_address(range, right, down)
    
def relative_cell_address_loc(target_cell_loc: Tuple[int], referred_cell_loc: Tuple[int]) -> Tuple[int]:
    """概要
    参照セルを起点とするターゲットセルの相対座標をtuple型で返す。

    Parameters
    ----------
    target_cell_loc: tuple:
        ターゲットセルの座標を示すtuple型。(列番号、行番号)のint型を格納する。

    referred_cell_loc: tuple:
        参照セルの座標を示すtuple型。(列番号、行番号)のint型を格納する。

    Returns
    ----------
    return_tuple: tuple:
        参照セルを起点とするターゲットセルの座標を示すtuple型。(列番号、行番号)のint型を格納する。
    """
    relative_cell_loc = (target_cell_loc[1] - referred_cell_loc[1], \
                         target_cell_loc[0] - referred_cell_loc[0])
    if relative_cell_loc[0] <0 or relative_cell_loc[1] < 0:
        raise ValueError('セルの指定範囲が不適です。{}'.format(target_cell_loc, referred_cell_loc))
    return relative_cell_loc

def relative_range_address_loc(target_address: str, referred_cell_loc: Tuple[int]) -> Tuple[Tuple[int]]:
    """概要
    参照セルを起点とするターゲット範囲の相対座標をtuple型で返す。

    Parameters
    ----------
    target_add: tuple:
        ターゲットセルの座標を示すtuple型。(列番号、行番号)のint型を格納する。

    referred_cell_loc: tuple:
        参照セルの座標を示すtuple型。(列番号、行番号)のint型を格納する。

    Returns
    ----------
    return_tuple: tuple
        範囲の左上のセル、右下のセルそれぞれの参照セルに対する相対座標を格納するtuple型。
        範囲がセルを指定している場合は、そのセルの相対座標が重複して格納される。
    """
    if ':' in target_address:
        target_cell_list = target_address.split(':')
        first_cell_loc = from_cell_address_to_column_row_int(target_cell_list[0])
        first_relative_cell_loc = relative_cell_address_loc(first_cell_loc, referred_cell_loc)
        last_cell_loc = from_cell_address_to_column_row_int(target_cell_list[1])
        last_relative_cell_loc = relative_cell_address_loc(last_cell_loc, referred_cell_loc)
        return (first_relative_cell_loc, last_relative_cell_loc)
    else:
        target_cell_loc = from_cell_address_to_column_row_int(target_address)
        relative_cell_loc = relative_cell_address_loc(target_cell_loc, referred_cell_loc)
        return (relative_cell_loc, relative_cell_loc)
    
def from_range_address_list_to_each_cell_adress_list(range_list: list) -> list:
    """概要
    range_listに格納されている範囲を個々のセル番地を表すstr型に書き起こし、
    list型に格納する。

    Parameters
    ----------
    range_list: list
        範囲を表すstr型が格納されたlist型。

    cell_address_list: list
        個々のセル番地を表すstr型が格納されたlist型。
    """
    cell_address_list = []
    for range_val in range_list:
        if ':' in range_val:
            range_loc = range_val.split(':')
            if len(range_loc) != 2:
                raise ValueError('次の範囲に:が2つ以上含まれます。:{}'.format(range_val))
            first_cell_loc = from_cell_address_to_column_row_int(range_loc[0])
            last_cell_loc = from_cell_address_to_column_row_int(range_loc[1])
            for col_num in range(first_cell_loc[0], last_cell_loc[0] + 1):
                for row_num in range(first_cell_loc[1], last_cell_loc[1] + 1):
                    cell_address_list.append(toAlpha3(col_num) + str(row_num))
        else:
            cell_address_list.append(range_val)
    return cell_address_list

# def _duplicate_range_address_pattern(range_list: List[str], col_interval: int = 1, col_pattern_num: int = 1, 
#                             row_interval: int = 1, row_pattern_num: int = 1) -> List[str]:
#     """概要
#     範囲を指定するstr型を格納したlist型を受け取り、列方向及び行方向にコピーした範囲を表すstr型を格納する
#     list型を返す。

#     Parameters
#     -----------
#     range_list: list
#         範囲を指定するstr型を格納したlist型。

#     col_inteval: int, default 1
#         列方向にコピーする間隔を示すint型。

#     col_pattern_num: int, default 1
#         列方向にコピーする数を示すint型。

#     row_interval: in, default 1
#         行方向にコピーする間隔を示すint型。

#     row_pattern_num: int, default 1
#         行方向にコピーする数を示すint型。
#     """
#     duplicated_range_list = []
#     for range in range_list:
#         for i in range(0, col_pattern_num):
#             for j in range(0, row_pattern_num):
#                 duplicated_range_list.append(
#                     move_range_address(range, i*col_interval, j*row_interval))
#     return  duplicated_range_list

def get_max_range(range1: str, range2: str) -> str:
    """概要
    セルは範囲を表す2つの文字列から、両方を含む最小の範囲を示すstr型を返す。

    Parameters
    ----------
    range1, range2: str
        範囲を示すstr型。

    Returns
    ----------
    range: str
        ws1, ws2を包含する範囲のうち、最小の範囲を示すstr型。
    """
    address1_tuple = from_range_address_to_column_row_int(range1)
    address2_tuple = from_range_address_to_column_row_int(range2)
    left_col_num = min(address1_tuple[0][0], address2_tuple[0][0])
    top_row_num = min(address1_tuple[0][1], address2_tuple[0][1])
    right_col_num = max(address1_tuple[1][0], address2_tuple[1][0])
    bottom_row_num = max(address1_tuple[1][1], address2_tuple[1][1])
    return from_column_row_int_to_cell_address(left_col_num, top_row_num) + ':' \
        + from_column_row_int_to_cell_address(right_col_num, bottom_row_num)

def get_cell_address_from_range_address(range_address: str, loc: str = 'top_left') -> str:
    """概要
    範囲を示すstr型を受け取り、四隅のうち指定した位置のセルのセル番地を返す。

    Parameters
    ----------
    range_sddress: str
        範囲を示すstr型。

    loc: str
        範囲のうちの位置を示すstr型。

    Returns
    ----------
    cell_address: str
        範囲のうち指定した位置のセルのセル番地を示すstr型。
    """
    range_address = range_address.replace('$', '')
    if ':' in range_address:
        range_address_list = range_address.split(':')
        if len(range_address_list) != 2:
            raise ValueError('次のセル範囲に:が2つ以上含まれます。{}'.format(range_address))
        if loc == 'top_left':
            return range_address_list[0]
        elif loc == 'bottom_left':
            return from_cell_address_to_column_row_letter(range_address_list[0])[0] + \
                from_cell_address_to_column_row_letter(range_address_list[1])[1]
        elif loc == 'top_right':
            return from_cell_address_to_column_row_letter(range_address_list[1])[0] + \
                from_cell_address_to_column_row_letter(range_address_list[0])[1]
        elif loc == 'bottom_right':
            return range_address_list[1]
        else:
            raise ValueError('locにはtop_left, bottom_left, top_right, bottom_rightのいずれかを\
                             指定してください。')
    else:
        return range_address
    
def from_range_address_to_each_col_row_list(range_address: str) -> Tuple[List[str]]:
    col_row_loc = from_range_address_to_column_row_int(range_address)
    col_list = [toAlpha3(c_num) for c_num in range(col_row_loc[0][0], col_row_loc[1][0] + 1)]
    row_list = [str(r_num) for r_num in range(col_row_loc[0][1], col_row_loc[1][1] + 1)]
    return col_list, row_list

    
def from_str_num_to_text(value: str) -> str:
    """概要
    エクセルシートに文字列を記入できるようにする。
    例:FO001の「001」を記入できるようにする。

    Parameters
    ----------
    value: str
        記入したい文字列。

    Returns
    ----------
    text: str
        エクセルにテキストとして記入できる文字列。
    """
    return "'{}".format(value)
    
def from_range_address_list_to_range_address(range_address_list: List[str]) -> str:
    """概要
    リストに含まれる範囲のをすべて含む範囲のうち最小のものを返す。

    Parameters
    ----------
    range_address_list: List[str]
        範囲を示すstr型が格納されたlist型。

    Returns
    ----------
    range_address: str
        range_address_listに含まれる範囲のうち、すべてを含む範囲で最小のもの。
    """
    if len(range_address_list) == 0:
        return ''
    else:
        c_r_int_0 = from_range_address_to_column_row_int(range_address_list[0])
        c_min = c_r_int_0[0][0]
        c_max = c_r_int_0[1][0]
        r_min = c_r_int_0[0][1]
        r_max = c_r_int_0[1][1]
        if len(range_address_list) > 1:
            for i in range(1, len(range_address_list)):
                c_r_int = from_range_address_to_column_row_int(range_address_list[i])
                c_min = c_r_int[0][0] if c_min > c_r_int[0][0] else c_min
                c_max = c_r_int[1][0] if c_max < c_r_int[1][0] else c_max
                r_min = c_r_int[0][1] if r_min > c_r_int[0][1] else r_min
                r_max = c_r_int[1][1] if r_max < c_r_int[1][1] else r_max
        if c_min == r_min and c_max == r_max:
            return from_column_row_int_to_cell_address(c_min, r_min)
        else:
            return from_column_row_int_to_cell_address(c_min, r_min) + ':' \
                + from_column_row_int_to_cell_address(c_max, r_max)