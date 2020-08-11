#! /bin/bash

service mysql start

root_pw=${1}

mysql -u root --password=${root_pw} << EOFMYSQL
create database socialmails character set utf8;
create user socialmails@'%' identified by 'socialmails123';
grant all on socialmails.* to socialmails@'%' identified by 'socialmails123';
EOFMYSQL

mysql -u socialmails --password=socialmails123 socialmails << EOFMYSQL
CREATE TABLE eml_files (
  id varchar(10) NOT NULL,
  type varchar(100) NOT NULL,
  subject varchar(100) NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY id_type_sub_key (type,subject)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE mail_recipients (
  id varchar(10) NOT NULL,
  user_group varchar(100) NOT NULL,
  user_email varchar(100) NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY id_group_email_key (user_group,user_email)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE user_logs (
  id int(11) NOT NULL AUTO_INCREMENT,
  user_group varchar(100) NOT NULL,
  user_email varchar(100) NOT NULL,
  eml_type varchar(100) NOT NULL,
  eml_subject varchar(100) NOT NULL,
  date timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  action varchar(10) NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY uniq_user_logs (user_group,user_email,eml_type,eml_subject,action)
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8;
EOFMYSQL
