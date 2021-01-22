## 目的

iniファイルの値を取得するスクリプト



## iniファイルの構造

### セクション

パラメータのグループ分けに使われる。セクションは必ず付ける必要がある。

[section]



### パラメータ

name=value



### コメント

コメント開始はセミコロン（;）

; comment



## 特徴

大文字と小文字の区別：区別あり

空行：許可

名前/値の区切り文字：=（イコール）のみ

階層構造：無し

行頭にあるスペース：無視する

セクションが宣言されていないパラメータは無視する



### 関数名

readINI：iniファイルの情報を連想配列に入れる。getINIに呼び出されている内部的な関数

getINI：連想配列からセクションとキーに該当する値を取り出す



### 引数

sectionName：セクション名

keyName：キー名

fileName：INIファイル名



## 参考

[初期化ファイル(INIファイル)の読み書き](http://home.a00.itscom.net/hatada/windows/file/profile01.html)

[INIファイル - Wikipedia](https://ja.wikipedia.org/wiki/INI%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB)

[[VBS]VBscriptからINIの内容を取得する](https://kuroparu.com/iniget/)

[VBScript : ini ファイルの値を取得する](https://logicalerror.seesaa.net/article/129131803.html)
