# プロジェクト計画書、計画変更届のコピー、差分の赤字表示について

## はじめに
森林J-クレジット創出支援システム（以下FCS）はシステムに入力した情報に基づいてプロジェクト計画書、計画変更届、モニタリング報告書を出力する機能があるが、2025年度までに開発・改修したシステムにおいては、計画変更届において変更箇所を赤字で記載する機能が実装されていない。そこで、2つのプロジェクト計画書/計画変更届を受け取り、差分を赤字にして出力するスクリプトを実装した。また、プロジェクト計画書のうちシミュレーションに依存しない項目について、入力項目を別のファイルに書き写すスクリプトを実装した。

## コピー
```
from copy_keikaku import copy_keikaku_value

target_keikaku_path = '計画変更届.xlsx'
referred_keiakaku_path = 'プロジェクト登録書_変更前.xlsx'

copy_keikaku_value(
    target_keikaku_path　= target_keikaku_path,
    referred_keikaku_path = referred_keikaku_path,
    save_path = '',
    ver = '1.3.0',
    overwrite = False
)
```
copy_keikaku_valueに引数として指定できる変数は以下の通り。

#### target_keikaku_path
書き込みを行うプロジェクト計画書、計画変更届を指定するファイルパスを示すstr型。

#### referred_keikaku_path
シミュレーションに依存しない項目の値を参照するプロジェクト計画書、計画変更届のファイルパスを示すstr型。

#### save_path
`overwrite`に`False`が指定されている場合に、書き込みを行ったエクセルファイルの保存先を示すstr型。`''`が指定されている場合、`target_keikaku_path`、`_コピー(参照ファイル：referred_keikaku_path)_日付.xlsx`を保存先に指定する。デフォルトは`''`。

#### ver
書き込みを行うプロジェクト計画書のバージョンを指定するstr型。`1.3.0`は対応していないためエラーを返す。デフォルトは`1.3.0`。

#### overwrite
ファイルの上書きを行うか否かを示すbool型。`True`が指定されている場合、`save_path`の値によらずに`target_keikaku_path`に上書きされる。`False`が指定されている場合、`save_path`に対して与えられたパスに保存する。デフォルトは`False`。

## 差分の赤字変更
```
from check_henko import make_diff_red

target_keikaku_path = '計画変更届.xlsx'
referred_keiakaku_path = 'プロジェクト登録書_変更前.xlsx'


make_diff_red(
    target_file_path = target_file_path,
    referred_file_path = referred_file_path,
    overwrite = False,
    save_path = ''
)

target_file_path: str, referred_file_path: str, overwrite: bool = False,
                  save_path: str = '
```

make_diff_redに対して指定できる変数は以下の通り。

#### target_file_path
差分の赤字表示を行うプロジェクト計画書、計画変更届を指定するファイルパスを示すstr型。

#### referred_file_path
差分を参照するプロジェクト計画書、計画変更届のファイルパスを示すstr型。

#### overwrite
ファイルの上書きを行うか否かを示すbool型。`True`が指定されている場合、`save_path`の値によらずに`target_file_path`に上書きされる。`False`が指定されている場合、`save_path`に対して与えられたパスに保存する。デフォルトは`False`。

#### save_path
`overwrite`に`False`が指定されている場合に、書き込みを行ったエクセルファイルの保存先を示すstr型。`''`が指定されている場合、`target_keikaku_path`、`_赤字変更(参照ファイル：referred_keikaku_path)_日付.xlsx`を保存先に指定する。デフォルトは`''`。