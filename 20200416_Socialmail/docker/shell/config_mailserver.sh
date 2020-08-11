#! /bin/bash

sed -i "s/my_host_name/${1}/g" /etc/postfix/main.cf
sed -i "s/db_host_name/${2}/g" \
    /etc/postfix/mysql-virtual-alias-maps.cf \
    /etc/postfix/postfix/mysql-virtual-mailbox-domains.cf \
    /etc/postfix/mysql-virtual-mailbox-maps.cf

sed -i "s/my_host_name/${1}/g" /etc/dovecot/dovecot.conf
sed -i "s/db_host_name/${2}/g" /etc/dovecot/dovecot-sql.conf.ext