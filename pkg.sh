#!/bin/bash

echo 使用的相对路径，会清空同级目录下的preview目录
mkdir preview 
rm -rf preview/*
cp index.html preview/
echo ----------------生成mac--------------------
gitbook build Mac
mv Mac/_book preview/mac
echo ----------------生成windows----------------
gitbook build PC
mv PC/_book preview/windows
echo ----------------生成web--------------------
gitbook build web
mv Web/_book preview/web
echo ----------------生成android----------------
gitbook build android
mv android/_book preview/android
echo ----------------生成ipad-------------------
gitbook build iPad
mv iPad/_book preview/ipad
echo ----------------生成iphone-----------------
gitbook build iPhone
mv iPhone/_book preview/iPhone
echo ----------------success--------------------
