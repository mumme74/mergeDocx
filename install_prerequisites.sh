#!/usr/bin/env bash

declare -a CMDS=(python3 python py)

for cmd in "${CMDS[@]}"; do
  PY=`command -v $cmd --version`
  if [ -n "${PY}" ] ; then
    break
  fi
  echo $PY
done

if [ -z "$PY" ] ; then
  echo "Pyhton är inte installerat"
  exit 1
fi


echo "Använder $PY som commando för python"
set -x # echo on
${PY} -m pip install -r requirements.txt

