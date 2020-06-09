## 要件

iniファイルを読み込むためのスクリプト



基本構造



### パラメータ

name=value



### セクション

パラメータのグループ分けに使われる。

[section]



### コメント

コメント開始はセミコロン（;）

; comment



名前の重複：重複した場合、エラーを発生させる。

大文字と小文字の区別：区別なし

空行：許可

名前/値の区切り文字：=（イコール）のみ

階層構造：無し

行頭にあるスペース：無視する

セクションが宣言されていないパラメータは無視する



関数名





引数

iniSection：セクション名

iniKey：キー名

fileName：INIファイル名



Dictionary（連想配列）を使う

配列の要素の中に配列突っ込んでいるなコレ





## 参考

[初期化ファイル(INIファイル)の読み書き](http://home.a00.itscom.net/hatada/windows/file/profile01.html)

[INIファイル - Wikipedia](https://ja.wikipedia.org/wiki/INI%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB)

[[VBS]VBscriptからINIの内容を取得する](https://kuroparu.com/iniget/)

[VBScript : ini ファイルの値を取得する](https://logicalerror.seesaa.net/article/129131803.html)





