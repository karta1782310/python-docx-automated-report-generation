#! /bin/bash

sed -i "s/my_host_name/${APACHE_SEVER_NAME}/g" /etc/apache2/sites-available/socialmails.conf

service apache2 start
service atd start
python3 /var/www/socialmails/schedule_server/server.py &
/bin/bash