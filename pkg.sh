#!/bin/bash

echo 使用的相对路径，会清空同级目录下的manual目录
mkdir manual 
rm -rf manual/*
cp index.html manual/
echo ----------------生成mac--------------------
gitbook build Mac
mv Mac/_book manual/mac
echo ----------------生成windows----------------
gitbook build PC
mv PC/_book manual/windows
echo ----------------生成web--------------------
gitbook build web
mv Web/_book manual/web
echo ----------------生成android----------------
gitbook build android
mv android/_book manual/android
echo ----------------生成ipad-------------------
gitbook build iPad
mv iPad/_book manual/ipad
echo ----------------生成iphone-----------------
gitbook build iPhone
mv iPhone/_book manual/iphone
echo ----------------success--------------------
