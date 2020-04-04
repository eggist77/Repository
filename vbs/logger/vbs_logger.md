
vbs_logger

format(csv)
"date","loglevel","message"

loglevel=error,warning,info

property
(logger).setName -> ログファイルの名前。指定した名前＋yyyymmdd.logで作成される。
(logger).setLogpath -> ログの出力先。指定しなければvbsファイルのフォルダにlogフォルダを作成。
                       その中に出力

method
(logger).error "message"
(logger).warnig "message"
(logger).info "message"

EOF
