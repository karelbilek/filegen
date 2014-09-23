#!/usr/bin/env bash

#nevyznam se v silenym ANT buildu, co dela netbeans, tak do nej nechci moc sahat

ant

#netusim, proc tam ant neprida knihovny, a ROZHODNE se mi tu silenost nechce zkoumat
#tak radsi delam uplne cely rucne

cd dist 
jar umf ../MANIFEST.MF  FileGen.jar
cp -r ../libs lib
