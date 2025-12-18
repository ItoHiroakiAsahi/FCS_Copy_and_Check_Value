"""
2つのテキスト分を比較し、差分の情報を返す。
"""
import difflib
from typing import List, Tuple
import MeCab

def _wakati_list(text: str) -> List[str]:
    """概要
    Mecabによる形態素解析を行い、テキスト分を単語ごとのリストにして返す。

    Parameters
    ----------
    text: str
        形態素を行う文章を示すstr型。

    Returns
    ----------
    words: List[str]
        文章に含まれている単語ごとのList[str]型。
    """
    # 単語をスペース区切りで出力する
    tagger = MeCab.Tagger("-Owakati")
    words = tagger.parse(text).strip().split()
    return words

def _words_to_char_loc_and_len(words: List[str], start_int: int, end_int: int
                               ) -> Tuple[int]:
    """概要
    単語が格納されたList[str]型を受け取り、もともとの文章における指定した単語の文字の位置、
    並びに指定した単語から指定した単語までの間の文字数をTuple[int]型で返す。
    
    Parameters
    -----------
    words: List[str]
        単語が格納されたList[str]型。

    start_int: int
        wordsにおける文字数のカウントを開始する単語の位置を示すint型。
    
    end_int: int
        wordsにおける文字数のカウントを終了する単語の位置を示すint型。

    Returns
    ----------
    t: Tuple[int]
        文字数のカウントを開始する単語の先頭の文字の位置、および
        文字数のカウントを開始してから終了するまでの文字数を格納するTuple[int]型。
    """
    s_loc = 0
    for i in range(start_int):
        s_loc += len(words[i])
    s_len = 0
    for j in range(start_int, end_int):
        s_len += len(words[j])
    return (s_loc, s_len)

def find_text_diff(target_text: str, referred_text: str) -> List[Tuple[int]]:
    """概要
    2つの文章を受け取り、差分があった場合に差分の開始する文字位置と差分のある文字数の長さを
    Tuple[int]型に格納したList[Tuple[int]]型を返す。
    
    Parameters
    ----------
    target_text: str
        差分を比較する文章を示すstr型。返り値のおける差分の文字位置、文字数の長さは
        target_textに対するものである。

    referred_text: str
        差分を比較する文章を示すstr型。

    Returns
    ----------
    diff_char_tuple_list: List[Tuple[int]]
        2つの文章の差分のうち、target_textに対して追加または変更されたものについて、
        差分の開始する文字位置と各差分の文字数の長さをペアにしたTuple[int]型を作成し、
        各差分をList[Tuple[int]]型に格納したもの。
    """
    target_words = _wakati_list(target_text)
    referred_words = _wakati_list(referred_text)
    sm = difflib.SequenceMatcher(None, referred_words, target_words)
    diff_char_tuple_list = []
    for opcode, _, _, j1, j2 in sm.get_opcodes():
        if opcode == 'equal' or opcode == 'delete':
            pass
        elif opcode == 'insert' or opcode == 'replace':
            diff_char_tuple_list.append(_words_to_char_loc_and_len(target_words, j1, j2))       
    return diff_char_tuple_list