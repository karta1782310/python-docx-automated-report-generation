#! /bin/bash

locale-gen zh_TW.UTF-8
echo 'export LANGUAGE="zh_TW.UTF-8"' >> /root/.bashrc
echo 'export LANG="zh_TW.UTF-8"' >> /root/.bashrc
echo 'export LC_ALL="zh_TW.UTF-8"' >> /root/.bashrc
update-locale LANG=zh_TW.UTF-8