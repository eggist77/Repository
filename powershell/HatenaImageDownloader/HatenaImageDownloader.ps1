
# @description Hatena Photo image Downloader
# @auther T.N.
# @version 1.0
# @since 2021-01-04
# @update 2021-01-04

$f = (Get-Content url.txt) -as [string[]]

$i=1
foreach ($url in $f) {
    $a = $url.Split("/")
    $fname = $a[-2] + "_" + $a[-1]
    Invoke-WebRequest -Uri $url -OutFile $fname
    $i++
}

