#!/bin/bash
if [[ ! "$USER" = "codepiano" ]]; then
    node_bin=$(which node)
    node_home=${node_bin%%/bin/node}
    echo "guess nodejs home is:${node_home}"
    gitbook_home=${node_home}/lib/node_modules/gitbook
    echo "guess gitbook module path is:$gitbook_home"
    if [[ -d $gitbook_home ]]; then
        echo "try to replace favicon"
        cp favicon.ico $gitbook_home/theme/assets/images/favicon.ico
    fi
fi

function build() {
    echo 使用的相对路径，会清空同级目录下的manual目录
    mkdir manual
    rm -rf manual/*
    cp index.html manual/
    echo ---------------- mac --------------------
    gitbook build Mac -f "${format}" -o "manual/mac" "$verbose"
    echo ---------------- windows ----------------
    gitbook build PC -f "${format}" -o "manual/windows" "$verbose"
    echo ---------------- web --------------------
    gitbook build Web -f "${format}" -o "manual/web" "$verbose"
    echo ---------------- android ----------------
    gitbook build android -f "${format}" -o "manual/android" "$verbose"
    echo ---------------- ipad -------------------
    gitbook build iPad -f "${format}" -o "manual/ipad" "$verbose"
    echo ---------------- iphone -----------------
    gitbook build iPhone -f "${format}" -o "manual/iphone" "$verbose"
    echo ---------------- enterprise -----------------
    gitbook build enterprise -f "${format}" -o "manual/enterprise" "$verbose"
    echo ---------------- plugin -----------------
    gitbook build plugin -f "${format}" -o "manual/plugin" "$verbose"
    echo ---------------- wizbox -----------------
    gitbook build wizbox -f "${format}" -o "manual/wizbox" "$verbose"
    echo ----------------success--------------------
}

# default build format
format='site'

if [[ ! $# -eq 0 ]]; then
	while getopts "f:" optname
	do
		case "$optname" in
			"f")
                echo "$OPTARG"
                if [ ! "$OPTARG" = '' ];then
                    format="$OPTARG"
                fi
				;;
        esac
	done
fi
build
# only check pdf
if [[ "$format" = 'ebook' ]];then
    mv manual/mac/index.pdf ./mac.pdf
    mv manual/windows/index.pdf ./windows.pdf
    mv manual/web/index.pdf ./web.pdf
    mv manual/android/index.pdf ./android.pdf
    mv manual/ipad/index.pdf ./ipad.pdf
    mv manual/iphone/index.pdf ./iphone.pdf
    mv manual/enterprise/index.pdf ./enterprise.pdf
    mv manual/wizbox/index.pdf ./wizbox.pdf
    mv manual/plugin/index.pdf ./plugin.pdf
fi
