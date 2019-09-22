#!/bin/sh

pushd $PWD > /dev/null
cd $(dirname $0)/lib
echo $PWD
echo Please enter target database information...
read -p "SID [$ORACLE_SID]: " sid
sid=${sid:-$ORACLE_SID}
read -p "User: " user
loadjava -u $user@$sid -r -v -jarsasdbobjects -fileout ../install_jdk5.log stax-api-1.0-2.jar sjsxp-1.0.2.jar exceldbtools-1.5.jar
popd > /dev/null
