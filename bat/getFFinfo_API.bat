@echo off

rem VBA�}�N�����AExcel�������Ă���J�����g�f�B���N�g���Ɉړ������܂��B
cd %1

rem �ړ������f�B���N�g�����ɂ���FF�����擾����Python���N�������܂��B
python ProgramPy\getFFinfo_API.py %2

rem �����������ł������m�F���܂��BVBA�̃t���O��ON�\OFF�𔻒f���܂��B
if %3 equ 1 (
	pause
)