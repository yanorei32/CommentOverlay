HOST_PORT=`ipconfig | \
	iconv -f cp932 -t utf-8 | \
	grep Wi-Fi -A 20 | \
	grep IPv4 | \
	grep -o '\([0-9]\+\.\)\{3\}[0-9]\+' | \
	head -n 1`\:6928

echo http://$HOST_PORT> /dev/clipboard

./comment.exe http://$HOST_PORT ${@:1} \
	| iconv -f cp932 -t utf-8

