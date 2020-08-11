create database mailserver character set utf8;
create user mailserver@'%' identified by 'mailserver123';
grant all on mailserver.* to mailserver@'%' identified by 'mailserver123';

use mailserver;
CREATE TABLE virtual_domains (  
  id int(11) NOT NULL auto_increment,  
  name varchar(50) NOT NULL,  
  PRIMARY KEY (id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE virtual_users (  
  id int(11) NOT NULL auto_increment,  
  domain_id int(11) NOT NULL,  
  password varchar(106) NOT NULL,  
  email varchar(100) NOT NULL,  
  PRIMARY KEY (id),  
  UNIQUE KEY email (email),  
  FOREIGN KEY (domain_id) REFERENCES virtual_domains(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
    
CREATE TABLE virtual_aliases (  
  id int(11) NOT NULL auto_increment,  
  domain_id int(11) NOT NULL,  
  source varchar(100) NOT NULL,  
  destination varchar(100) NOT NULL,  
  PRIMARY KEY (id),  
  FOREIGN KEY (domain_id) REFERENCES virtual_domains(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- insert into virtual_domains(id,name) values(1,"${SMTP_HOST_NAME}");

-- insert into virtual_users(id,domain_id,password,email)  
-- values (1,1,ENCRYPT('root1234', CONCAT('\$6$', SUBSTRING(SHA(RAND()), -16))),"root@${SMTP_HOST_NAME}");
-- insert into virtual_users(id,domain_id,password,email)  
-- values (2,1,ENCRYPT('yucc1234', CONCAT('\$6$', SUBSTRING(SHA(RAND()), -16))),"yucc@${SMTP_HOST_NAME}");

-- insert into virtual_aliases(id,domain_id,source,destination)  
-- values (1,1,"all@${SMTP_HOST_NAME}","root@${SMTP_HOST_NAME}");
-- insert into virtual_aliases(id,domain_id,source,destination) 
-- values (2,1,"all@${SMTP_HOST_NAME}","yucc@${SMTP_HOST_NAME}");

create database socialmails character set utf8;
create user socialmails@'%' identified by 'socialmails123';
grant all on socialmails.* to socialmails@'%' identified by 'socialmails123';

use socialmails;
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
  eml_id varchar(10) NOT NULL,
  recipient_id varchar(10) NOT NULL,
  date timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  action varchar(10) NOT NULL,
  PRIMARY KEY (id),
  FOREIGN KEY (eml_id) REFERENCES eml_files(id) ON DELETE NO ACTION,
  FOREIGN KEY (recipient_id) REFERENCES mail_recipients(id) ON DELETE NO ACTION,
  UNIQUE KEY uniq_user_logs (eml_id,recipient_id,action)
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8;

-- use socialmails;
-- CREATE TABLE eml_files (
--   id varchar(10) NOT NULL,
--   type varchar(100) NOT NULL,
--   subject varchar(100) NOT NULL,
--   PRIMARY KEY (id),
--   UNIQUE KEY id_type_sub_key (type,subject)
-- ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- CREATE TABLE mail_recipients (
--   id varchar(10) NOT NULL,
--   user_group varchar(100) NOT NULL,
--   user_email varchar(100) NOT NULL,
--   PRIMARY KEY (id),
--   UNIQUE KEY id_group_email_key (user_group,user_email)
-- ) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- CREATE TABLE user_logs (
--   id int(11) NOT NULL AUTO_INCREMENT,
--   user_group varchar(100) NOT NULL,
--   user_email varchar(100) NOT NULL,
--   eml_type varchar(100) NOT NULL,
--   eml_subject varchar(100) NOT NULL,
--   date timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
--   action varchar(10) NOT NULL,
--   PRIMARY KEY (id),
--   UNIQUE KEY uniq_user_logs (user_group,user_email,eml_type,eml_subject,action)
-- ) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8;

