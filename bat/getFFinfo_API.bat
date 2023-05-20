@echo off

rem VBAマクロより、Excelがおいてあるカレントディレクトリに移動させます。
cd %1

rem 移動したディレクトリ内にあるFF情報を取得するPythonを起動させます。
python ProgramPy\getFFinfo_API.py %2

rem 正しく処理できたか確認します。VBAのフラグでON―OFFを判断します。
if %3 equ 1 (
	pause
)