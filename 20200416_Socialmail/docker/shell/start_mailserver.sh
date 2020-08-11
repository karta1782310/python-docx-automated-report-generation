#! /bin/bash

sed -i "s/my_host_name/${SMTP_HOST_NAME}/g" /etc/postfix/main.cf
sed -i "s/db_host_name/${DB_HOST_NAME}/g" \
    /etc/postfix/mysql-virtual-alias-maps.cf \
    /etc/postfix/mysql-virtual-mailbox-domains.cf \
    /etc/postfix/mysql-virtual-mailbox-maps.cf

sed -i "s/my_host_name/${SMTP_HOST_NAME}/g" /etc/dovecot/dovecot.conf
sed -i "s/db_host_name/${DB_HOST_NAME}/g" /etc/dovecot/dovecot-sql.conf.ext

service postfix start
service dovecot start
service rsyslog start
/bin/bash