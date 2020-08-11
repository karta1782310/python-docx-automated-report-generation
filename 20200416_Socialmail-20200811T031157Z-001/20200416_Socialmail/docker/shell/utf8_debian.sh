#! /bin/bash

# not work very well: only display, not type

locale-gen zh_TW.UTF-8
update-locale LANG=zh_TW.UTF-8
echo 'export LANGUAGE="zh_TW.UTF-8"' >> /root/.bashrc
echo 'export LANG="zh_TW.UTF-8"' >> /root/.bashrc
echo 'export LC_ALL="zh_TW.UTF-8"' >> /root/.bashrc
echo 'export LC_LANG="zh_TW.UTF-8"' >> /root/.bashrc
dpkg-reconfigure locales
