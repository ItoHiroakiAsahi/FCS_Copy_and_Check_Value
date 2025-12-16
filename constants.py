"""
プロジェクト計画書（計画変更届）の書き込み、赤字変更処理の際に使用する定数情報を定義する。
"""
from enum import Enum

class KeikakuSheet(Enum):
    """概要
    プロジェクト計画書のシート名を格納する。
    """
    INTRODUCTION = 'はじめに（本資料の説明）'
    EXPLANATION_FO001 = '【FO-001】対象地および実施者の義務に係る手続きの説明'
    EXPLANATION_FO003 = '【FO-003】対象地および実施者の義務に係る手続きの説明'
    REGISTER_APPLICATION = '登録申請書'
    OTHER_PARTICIPANTS = '代表以外のプロジェクト実施者もしくはプログラム型運営・管理者'
    CONFIRMATION_OF_CONDITION_5 = '（参考様式）FO-001適用条件5に係る確認書'
    CONFIRMATION_OF_CONDITION_6 = '（参考様式）FO-001適用条件6に係る誓約書'
    CONFIRMATION_OF_CONDITION_7 = '（参考様式）FO-003適用条件7に係る誓約書'
    CONFIRMATION_OF_FO003_OBLIGATION = '（参考様式）FO-003のプロジェクト実施者の義務に係る確認書'
    PLAN_CHANGE_NORTIFICATION = '【HP公開】計画変更届'
    PLAN_COVER = '【HP公開】プロジェクト計画書表紙'
    OVERVIEW = '【HP公開】1.プロジェクト概要'
    METHODOLOGY_FO001 = '【HP公開】2.適用する方法論（FO-001）'
    SKK_CHANGES = '森林経営計画の適用条件１への適用と計画の変遷（FO-001）'
    MULTIPLE_SKK_INFO = '2.2 複数森林経営計画用（FO-001）'
    METHODOLOGY_NORMAL_FO003 = '【HP公開】2.適用する方法論（FO-003通常型）'
    METHODOLOGY_PROGRAM_FO003 = '【HP公開】2.適用する方法論（FO-003プログラム型）'
    OPERATION_MANAGEMENT_SYSTEM_FO003 = '【HP公開】2.6FO-003プログラム型の運営・管理体制'
    FOREST_DRAWING = 'プロジェクト実施地の図面（FO-001,003共通）'
    DATA_MANAGEMENT = '【HP公開】3.データ管理'
    SPECIAL_NOTES = '【HP公開】4.特記事項'
    ADDITIONALITY = '5.追加性に関する情報'
    ABSORPTION_AMOUNT = '【HP公開】6.吸収量の算定方法'
    PROGRAM_FO003_PLAN = '【HP公開】6-6.FO-003プログラム型の活動計画'
    MONITORING_PLAN_FO001 = '【HP公開】7.モニタリング計画（FO-001）'
    MONITORING_PLAN_FO003 = '【HP公開】7.モニタリング計画（FO-003）'
    IKUSEI_RSH = '幹材積量算定シート_育成林および主伐用（001、003共通）'
    IKUSEI_INFO = '【吸収量（育成林）算定用】情報記入シート（001、003共通）'
    IKUSEI_CALCULATION = '（自動計算）吸収量（育成林）算定シート（001、003共通）'
    TENNEN_RSH = '幹材積量算定シート_天然生林（FO-001）'
    TENNEN_INFO = '【吸収量（天然生林）算定用】情報記入シート（FO-001）'
    TENNEN_CALCULATION = '（自動計算）吸収量（天然生林）算定シート（FO-001）'
    IN_PJ_EMISSION_INFO = '【排出量（PJ内）算定用】情報記入シート（001、003共通）'
    IN_PJ_EMISSION_CALCULAITON = '（自動計算）排出量（PJ内）算定シート（001、003共通)'
    IN_PJ_HWP = '【吸収量（PJ内HWP）】情報記入・算定シート（FO-001）'
    OUT_PJ_INFO = '【主伐再造林（PJ外）算定用】情報記入シート（FO-001）'
    OUT_PJ_CALCULATION = '（自動計算）主伐再造林（PJ外）算定シート（FO-001）'
    RESERVATION_RECORD = '（参考様式）「森林の保護」実施記録例（FO-001）'
    IKUSEI_RSH_TABLE = '（記入不要）幹材積量シート_育成林'
    TENNEN_RSH_TABLE = '（記入不要）幹材積量シート_天然生林'
    CUTTING_RSH_TABLE = '（記入不要）幹材積量シート_主伐用'
    FISCAL_YEAR_TABLE = '（記入不要）年度計算シート'

class Color(Enum):
    """概要
    エクセルの書き込みに使用する色の値を格納する。
    """
    RED = 0xFF

class ChangeFlag(Enum):
    """概要
    書類に書き込む変更の有無のフラグ
    """
    CHANGED = "有"
    NOT_CHANGED = "無"