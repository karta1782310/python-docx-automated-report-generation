#! /bin/bash

service postfix start
service dovecot start
service mysql start
service rsyslog start
service apache2 start
service atd start
python3 /var/www/socialmails/schedule_server/server.py &
 