@echo off
echo ******************* > who.txt
echo  hostsファイル検索結果 >> who.txt
echo ******************* >> who.txt

find "living" %windir%\System32\drivers\etc\hosts >> who.txt
notepad.exe who.txt