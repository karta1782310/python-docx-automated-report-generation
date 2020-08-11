#! /bin/bash

service mysql start

mysql -u mailserver --password=mailserver123 mailserver << EOFMYSQL
insert into virtual_domains(id,name) values(1,'${SMTP_HOST_NAME}');

insert into virtual_users(id,domain_id,password,email)  
values (1,1,ENCRYPT('root1234', CONCAT('\$6$', SUBSTRING(SHA(RAND()), -16))),'root@$SMTP_HOST_NAME');
insert into virtual_users(id,domain_id,password,email)  
values (2,1,ENCRYPT('yucc1234', CONCAT('\$6$', SUBSTRING(SHA(RAND()), -16))),'yucc@$SMTP_HOST_NAME');

insert into virtual_aliases(id,domain_id,source,destination)  
values (1,1,'all@$SMTP_HOST_NAME','root@$SMTP_HOST_NAME');
insert into virtual_aliases(id,domain_id,source,destination) 
values (2,1,'all@$SMTP_HOST_NAME','yucc@$SMTP_HOST_NAME');
EOFMYSQL


# root_pw=${1}
# domain_name=${2}

# mysql -u root --password=${root_pw} << EOFMYSQL
# create database mailserver character set utf8;
# create user mailserver@'%' identified by 'mailserver123';
# grant all on mailserver.* to mailserver@'%' identified by 'mailserver123';
# EOFMYSQL

# mysql -u mailserver --password=mailserver123 mailserver << EOFMYSQL
# CREATE TABLE virtual_domains (  
#   id int(11) NOT NULL auto_increment,  
#   name varchar(50) NOT NULL,  
#   PRIMARY KEY (id)
# ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

# CREATE TABLE virtual_users (  
#   id int(11) NOT NULL auto_increment,  
#   domain_id int(11) NOT NULL,  
#   password varchar(106) NOT NULL,  
#   email varchar(100) NOT NULL,  
#   PRIMARY KEY (id),  
#   UNIQUE KEY email (email),  
#   FOREIGN KEY (domain_id) REFERENCES virtual_domains(id) ON DELETE CASCADE
# ) ENGINE=InnoDB DEFAULT CHARSET=utf8;
    
# CREATE TABLE virtual_aliases (  
#   id int(11) NOT NULL auto_increment,  
#   domain_id int(11) NOT NULL,  
#   source varchar(100) NOT NULL,  
#   destination varchar(100) NOT NULL,  
#   PRIMARY KEY (id),  
#   FOREIGN KEY (domain_id) REFERENCES virtual_domains(id) ON DELETE CASCADE
# ) ENGINE=InnoDB DEFAULT CHARSET=utf8;
# EOFMYSQL
