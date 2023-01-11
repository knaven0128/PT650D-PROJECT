# SQL Manager 2005 Lite for MySQL 3.7.7.1
# ---------------------------------------
# Host     : localhost
# Port     : 3306
# Database : weighingscaledb


SET FOREIGN_KEY_CHECKS=0;

DROP DATABASE IF EXISTS `weighingscaledb`;

CREATE DATABASE `weighingscaledb`
    CHARACTER SET 'latin1'
    COLLATE 'latin1_swedish_ci';

USE `weighingscaledb`;

#
# Structure for the `tblcomm` table : 
#

CREATE TABLE `tblcomm` (
  `commid` int(11) NOT NULL auto_increment,
  `portnum` int(11) default NULL,
  `commset` varchar(30) default NULL,
  `comm_len` int(11) default NULL,
  `comm_str` int(11) default NULL,
  `comm_symbol` varchar(30) default NULL,
  `comm_power` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`commid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblcomm1` table : 
#

CREATE TABLE `tblcomm1` (
  `commid` int(11) NOT NULL auto_increment,
  `portnum` int(11) default NULL,
  `commset` varchar(50) default NULL,
  `comm_len` int(20) default NULL,
  `comm_str` int(20) default NULL,
  `comm_symbol` varchar(30) default NULL,
  `comm_power` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`commid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblcount` table : 
#

CREATE TABLE `tblcount` (
  `countnumber` varchar(20) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblcustomer` table : 
#

CREATE TABLE `tblcustomer` (
  `customerid` int(11) NOT NULL auto_increment,
  `customer_name` varchar(29) default NULL,
  `address` varchar(30) default NULL,
  `contact_number` varchar(20) default NULL,
  `date_added` date default NULL,
  PRIMARY KEY  (`customerid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tbldestination` table : 
#

CREATE TABLE `tbldestination` (
  `id` int(11) NOT NULL auto_increment,
  `destination` varchar(20) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tbllogs` table : 
#

CREATE TABLE `tbllogs` (
  `logsId` int(11) NOT NULL auto_increment,
  `userName` varchar(100) default NULL,
  `actionLog` varchar(1000) default NULL,
  `dateLogs` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`logsId`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblproduct` table : 
#

CREATE TABLE `tblproduct` (
  `productid` int(11) NOT NULL auto_increment,
  `product_name` varchar(30) default NULL,
  `product_price` double(15,2) default NULL,
  `details` varchar(30) default NULL,
  `date_added` date default NULL,
  PRIMARY KEY  (`productid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblsetup` table : 
#

CREATE TABLE `tblsetup` (
  `soldidcnt` varchar(20) default NULL,
  `countNum` varchar(20) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblunitmeasure` table : 
#

CREATE TABLE `tblunitmeasure` (
  `unitid` int(11) NOT NULL auto_increment,
  `unit_name` varchar(30) default NULL,
  `unit_symbol` varchar(20) default NULL,
  PRIMARY KEY  (`unitid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tblweighing` table : 
#

CREATE TABLE `tblweighing` (
  `weighid` int(11) NOT NULL auto_increment,
  `consec_no` varchar(20) default NULL,
  `plate_number` varchar(29) default NULL,
  `weigher` varchar(30) default NULL,
  `transaction_date` date default NULL,
  `weigh_in` int(20) default NULL,
  `weigh_out` int(20) default NULL,
  `net_weight` int(20) default NULL,
  `qty` int(20) default NULL,
  `Unit` varchar(30) default NULL,
  `price` double(15,2) default NULL,
  `totalprice` double(15,2) default NULL,
  `datetime_weighin` datetime default NULL,
  `datetime_weighout` datetime default NULL,
  `customer_name` varchar(30) default NULL,
  `product_name` varchar(30) default NULL,
  `average` varchar(30) default '0.00',
  `scale_price` double(15,2) default NULL,
  `status` varchar(20) default NULL,
  `remarks` varchar(300) default NULL,
  `countnum` varchar(20) default NULL,
  `delStatus` tinyint(1) NOT NULL default '0',
  `destination` varchar(20) default NULL,
  PRIMARY KEY  (`weighid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for the `tblcomm` table  (LIMIT 0,500000)
#

INSERT INTO `tblcomm` (`commid`, `portnum`, `commset`, `comm_len`, `comm_str`, `comm_symbol`, `comm_power`) VALUES 
  (1,3,'9600,N,8,1',6,3,')',1);

COMMIT;

#
# Data for the `tblcomm1` table  (LIMIT 0,500000)
#

INSERT INTO `tblcomm1` (`commid`, `portnum`, `commset`, `comm_len`, `comm_str`, `comm_symbol`, `comm_power`) VALUES 
  (1,7,'9600,N,8,1',20,5,'$',1);

COMMIT;

#
# Data for the `tblcount` table  (LIMIT 0,500000)
#

INSERT INTO `tblcount` (`countnumber`) VALUES 
  ('0000001');

COMMIT;

#
# Data for the `tblcustomer` table  (LIMIT 0,500000)
#

INSERT INTO `tblcustomer` (`customerid`, `customer_name`, `address`, `contact_number`, `date_added`) VALUES 
  (1,'Customer','iloilo','123','2021-03-21'),
  (2,'4WARD','1','1','2022-04-21'),
  (3,'Oliver','1','1','2022-08-06'),
  (4,'Jeep','1','1','2022-10-06');

COMMIT;

#
# Data for the `tbldestination` table  (LIMIT 0,500000)
#

INSERT INTO `tbldestination` (`id`, `destination`) VALUES 
  (1,'test1');

COMMIT;

#
# Data for the `tbllogs` table  (LIMIT 0,500000)
#

INSERT INTO `tbllogs` (`logsId`, `userName`, `actionLog`, `dateLogs`) VALUES 
  (1,'knaven','Clear all Logs','2022-10-26 01:10:38'),
  (2,'knaven','User LogOut','2022-10-26 01:12:26'),
  (3,'w','User Login','2022-10-26 01:12:30'),
  (4,'w','User LogOut','2022-10-26 01:12:40'),
  (5,'s','User Login','2022-10-26 01:12:44'),
  (6,'knaven','User Login','2022-10-26 01:14:55'),
  (7,'knaven','User LogOut','2022-10-26 01:15:02'),
  (8,'s','User Login','2022-10-26 01:15:03'),
  (9,'knaven','User Login','2022-10-26 01:15:54'),
  (10,'knaven','User LogOut','2022-10-26 01:15:58'),
  (11,'s','User Login','2022-10-26 01:16:01'),
  (12,'knaven','User Login','2022-11-07 11:03:03'),
  (13,'knaven','User Login','2022-11-07 11:03:46'),
  (14,'s','User Login','2022-11-07 11:05:02'),
  (15,'s','User Login','2022-11-07 11:05:39'),
  (16,'s','User Login','2022-11-07 11:07:11'),
  (17,'s','User Login','2022-11-07 11:17:05'),
  (18,'s','User Login','2022-11-07 11:19:03'),
  (19,'s','User Login','2022-11-07 11:19:24'),
  (20,'s','User Login','2022-11-07 11:20:49'),
  (21,'s','User Login','2022-11-07 11:22:01'),
  (22,'s','User LogOut','2022-11-07 11:22:12'),
  (23,'encoder','User Login','2022-11-07 11:22:20'),
  (24,'encoder','Scale Offline','2022-11-07 11:22:27'),
  (25,'encoder','Scale Offline - granted by: a','2022-11-07 11:22:40'),
  (26,'s','User Login','2022-11-07 11:23:01'),
  (27,'s','Scale Offline','2022-11-07 11:23:11'),
  (28,'s','User Login','2022-11-07 11:24:19'),
  (29,'s','User Login','2022-11-07 11:25:48'),
  (30,'s','User Login','2022-11-07 11:26:56'),
  (31,'s','User Login','2022-11-07 11:28:16'),
  (32,'s','User Login','2022-11-07 11:29:41'),
  (33,'s','User Login','2022-11-07 11:31:26'),
  (34,'s','User Login','2022-11-07 11:32:24'),
  (35,'s','Scale Offline','2022-11-07 11:32:55'),
  (36,'s','User Login','2022-11-07 11:34:26'),
  (37,'s','User Login','2022-11-07 11:35:19'),
  (38,'s','User Login','2022-11-07 11:39:02'),
  (39,'s','User Login','2022-11-07 11:41:57'),
  (40,'s','User Login','2022-11-07 11:42:18'),
  (41,'s','User Login','2022-11-07 11:42:50'),
  (42,'s','User Login','2022-11-07 17:45:12'),
  (43,'s','User Login','2022-11-07 17:48:45'),
  (44,'s','User Login','2022-11-07 17:51:37'),
  (45,'s','User Login','2022-11-07 17:52:28'),
  (46,'s','Weigh OUT - Transaction No.: 00014- Plate Number: 123','2022-11-07 17:54:11'),
  (47,'s','Weigh IN - Transaction No.: 00017 - Plate Number: 2323','2022-11-07 17:56:31'),
  (48,'s','User Login','2022-11-07 17:57:33'),
  (49,'s','User Login','2022-11-07 17:59:24'),
  (50,'s','Weigh OUT - Transaction No.: 00017- Plate Number: 2323','2022-11-07 18:00:17'),
  (51,'s','User Login','2022-11-07 18:07:16'),
  (52,'s','Weigh OUT - Transaction No.: 00016- Plate Number: 12323','2022-11-07 18:07:50'),
  (53,'s','User Login','2022-11-07 18:20:39'),
  (54,'s','User Login','2022-11-07 18:22:02'),
  (55,'s','User Login','2022-11-07 18:23:16'),
  (56,'s','User Login','2022-11-07 18:25:14'),
  (57,'s','User Login','2022-11-07 18:26:38'),
  (58,'s','Add Destination - Destination name: test','2022-11-07 18:27:08'),
  (59,'s','Update Destination - Destination id: 1 - Destination name: test','2022-11-07 18:27:30'),
  (60,'s','User Login','2022-11-07 18:29:50'),
  (61,'s','User Login','2022-11-07 18:44:37'),
  (62,'s','User Login','2022-11-07 18:45:36'),
  (63,'s','User Login','2022-11-07 18:46:36'),
  (64,'s','User Login','2022-11-07 18:48:44'),
  (65,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:48:59'),
  (66,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:49:08'),
  (67,'s','User Login','2022-11-07 18:49:55'),
  (68,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:50:04'),
  (69,'s','User Login','2022-11-07 18:52:25'),
  (70,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:52:34'),
  (71,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:52:41'),
  (72,'s','User Login','2022-11-07 18:53:21'),
  (73,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:53:33'),
  (74,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:53:41'),
  (75,'s','User Login','2022-11-07 18:54:41'),
  (76,'s','Reprint OUT - Transaction No.: 00017','2022-11-07 18:54:51'),
  (77,'s','User Login','2022-11-07 18:56:00'),
  (78,'s','User Login','2022-11-07 18:57:48'),
  (79,'s','User Login','2022-11-07 18:59:01'),
  (80,'s','User Login','2022-11-07 18:59:49'),
  (81,'s','Weigh IN - Transaction No.: 00018 - Plate Number: 233','2022-11-07 19:01:31'),
  (82,'s','Weigh OUT - Transaction No.: 00018- Plate Number: 233','2022-11-07 19:02:02'),
  (83,'s','User Login','2022-11-07 19:05:01'),
  (84,'s','User Login','2022-11-07 19:07:08'),
  (85,'s','Delete All Weighing','2022-11-07 19:08:07'),
  (86,'s','Delete All Weighing','2022-11-07 19:08:24'),
  (87,'s','User Login','2022-11-07 19:11:23'),
  (88,'s','Weigh IN - Transaction No.: 00001 - Plate Number: 123','2022-11-07 19:11:57'),
  (89,'s','Scale Offline','2022-11-07 19:12:34'),
  (90,'s','Weigh OUT - Transaction No.: 00001- Plate Number: 123','2022-11-07 19:13:00'),
  (91,'s','User Login','2022-11-07 19:20:40'),
  (92,'s','Export to Excel Custom Report','2022-11-07 19:20:54'),
  (93,'s','Export to Excel Custom Report','2022-11-07 19:21:04'),
  (94,'s','User Login','2022-11-07 19:21:23'),
  (95,'s','Print Custom Report','2022-11-07 19:21:35'),
  (96,'s','Export to Excel Custom Report','2022-11-07 19:21:42'),
  (97,'s','User Login','2022-11-07 19:22:35'),
  (98,'s','Print Custom Report','2022-11-07 19:22:43'),
  (99,'s','Export to Excel Custom Report','2022-11-07 19:22:50'),
  (100,'s','Export to Excel Custom Report','2022-11-07 19:22:59');

COMMIT;

#
# Data for the `tblproduct` table  (LIMIT 0,500000)
#

INSERT INTO `tblproduct` (`productid`, `product_name`, `product_price`, `details`, `date_added`) VALUES 
  (1,'Palay',123,'Palay','2021-03-21'),
  (2,'Corn',123,'Corn','2021-03-21'),
  (3,'TEst',20,'2','2022-10-26');

COMMIT;

#
# Data for the `tblsetup` table  (LIMIT 0,500000)
#

INSERT INTO `tblsetup` (`soldidcnt`, `countNum`) VALUES 
  ('00001','0000000');

COMMIT;

#
# Data for the `tblunitmeasure` table  (LIMIT 0,500000)
#

INSERT INTO `tblunitmeasure` (`unitid`, `unit_name`, `unit_symbol`) VALUES 
  (1,'Kilo','kg');

COMMIT;

#
# Data for the `tblweighing` table  (LIMIT 0,500000)
#

INSERT INTO `tblweighing` (`weighid`, `consec_no`, `plate_number`, `weigher`, `transaction_date`, `weigh_in`, `weigh_out`, `net_weight`, `qty`, `Unit`, `price`, `totalprice`, `datetime_weighin`, `datetime_weighout`, `customer_name`, `product_name`, `average`, `scale_price`, `status`, `remarks`, `countnum`, `delStatus`, `destination`) VALUES 
  (1,'00001','123','s','2022-11-07',320,12,308,0,'Kilo',0,0,'2022-11-07 19:11:44','2022-11-07 19:13:00','Customer','Corn','0',0,'OUT','','0000001',0,'test1');

COMMIT;

