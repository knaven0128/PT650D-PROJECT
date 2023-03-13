-- MySQL dump 10.10
--
-- Host: localhost    Database: weighingscaledb
-- ------------------------------------------------------
-- Server version	5.0.24a-community-nt

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `tblcomm`
--

DROP TABLE IF EXISTS `tblcomm`;
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

--
-- Dumping data for table `tblcomm`
--


/*!40000 ALTER TABLE `tblcomm` DISABLE KEYS */;
LOCK TABLES `tblcomm` WRITE;
INSERT INTO `tblcomm` VALUES (1,3,'9600,N,8,1',30,1,'M',1);
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblcomm` ENABLE KEYS */;

--
-- Table structure for table `tblcomm1`
--

DROP TABLE IF EXISTS `tblcomm1`;
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

--
-- Dumping data for table `tblcomm1`
--


/*!40000 ALTER TABLE `tblcomm1` DISABLE KEYS */;
LOCK TABLES `tblcomm1` WRITE;
INSERT INTO `tblcomm1` VALUES (1,7,'9600,N,8,1',20,5,'$',1);
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblcomm1` ENABLE KEYS */;

--
-- Table structure for table `tblcount`
--

DROP TABLE IF EXISTS `tblcount`;
CREATE TABLE `tblcount` (
  `countnumber` varchar(20) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblcount`
--


/*!40000 ALTER TABLE `tblcount` DISABLE KEYS */;
LOCK TABLES `tblcount` WRITE;
INSERT INTO `tblcount` VALUES ('0000001');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblcount` ENABLE KEYS */;

--
-- Table structure for table `tblcustomer`
--

DROP TABLE IF EXISTS `tblcustomer`;
CREATE TABLE `tblcustomer` (
  `customerid` int(11) NOT NULL auto_increment,
  `customer_name` varchar(29) default NULL,
  `address` varchar(30) default NULL,
  `contact_number` varchar(20) default NULL,
  `date_added` date default NULL,
  PRIMARY KEY  (`customerid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblcustomer`
--


/*!40000 ALTER TABLE `tblcustomer` DISABLE KEYS */;
LOCK TABLES `tblcustomer` WRITE;
INSERT INTO `tblcustomer` VALUES (1,'Customer','iloilo','123','2021-03-21');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblcustomer` ENABLE KEYS */;

--
-- Table structure for table `tbllogs`
--

DROP TABLE IF EXISTS `tbllogs`;
CREATE TABLE `tbllogs` (
  `logsId` int(11) NOT NULL auto_increment,
  `userName` varchar(100) default NULL,
  `actionLog` varchar(1000) default NULL,
  `dateLogs` timestamp NOT NULL default CURRENT_TIMESTAMP on update CURRENT_TIMESTAMP,
  PRIMARY KEY  (`logsId`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tbllogs`
--


/*!40000 ALTER TABLE `tbllogs` DISABLE KEYS */;
LOCK TABLES `tbllogs` WRITE;
INSERT INTO `tbllogs` VALUES (1,'a','Clear all Logs','2021-03-25 14:36:28'),(2,'a','User Login','2021-03-25 14:37:07'),(3,'knaven','User Login','2021-03-25 14:37:23'),(4,'knaven','User Login','2021-03-25 14:41:42'),(5,'knaven','Back up All Records','2021-03-25 14:41:55');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tbllogs` ENABLE KEYS */;

--
-- Table structure for table `tblproduct`
--

DROP TABLE IF EXISTS `tblproduct`;
CREATE TABLE `tblproduct` (
  `productid` int(11) NOT NULL auto_increment,
  `product_name` varchar(30) default NULL,
  `product_price` double(15,2) default NULL,
  `details` varchar(30) default NULL,
  `date_added` date default NULL,
  PRIMARY KEY  (`productid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblproduct`
--


/*!40000 ALTER TABLE `tblproduct` DISABLE KEYS */;
LOCK TABLES `tblproduct` WRITE;
INSERT INTO `tblproduct` VALUES (1,'Palay',123.00,'Palay','2021-03-21'),(2,'Corn',123.00,'Corn','2021-03-21');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblproduct` ENABLE KEYS */;

--
-- Table structure for table `tblsetup`
--

DROP TABLE IF EXISTS `tblsetup`;
CREATE TABLE `tblsetup` (
  `soldidcnt` varchar(20) default NULL,
  `countNum` varchar(20) default NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblsetup`
--


/*!40000 ALTER TABLE `tblsetup` DISABLE KEYS */;
LOCK TABLES `tblsetup` WRITE;
INSERT INTO `tblsetup` VALUES ('00001','0000000');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblsetup` ENABLE KEYS */;

--
-- Table structure for table `tblunitmeasure`
--

DROP TABLE IF EXISTS `tblunitmeasure`;
CREATE TABLE `tblunitmeasure` (
  `unitid` int(11) NOT NULL auto_increment,
  `unit_name` varchar(30) default NULL,
  `unit_symbol` varchar(20) default NULL,
  PRIMARY KEY  (`unitid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblunitmeasure`
--


/*!40000 ALTER TABLE `tblunitmeasure` DISABLE KEYS */;
LOCK TABLES `tblunitmeasure` WRITE;
INSERT INTO `tblunitmeasure` VALUES (1,'Kilo','kg');
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblunitmeasure` ENABLE KEYS */;

--
-- Table structure for table `tblweighing`
--

DROP TABLE IF EXISTS `tblweighing`;
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
  PRIMARY KEY  (`weighid`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tblweighing`
--


/*!40000 ALTER TABLE `tblweighing` DISABLE KEYS */;
LOCK TABLES `tblweighing` WRITE;
INSERT INTO `tblweighing` VALUES (1,'00001','2112','Knaven Rey Sarroza','2021-03-21',213123,0,0,0,'NA',0.00,0.00,'2021-03-21 23:01:57',NULL,'Customer','Palay','0',0.00,'IN','','0000001',0);
UNLOCK TABLES;
/*!40000 ALTER TABLE `tblweighing` ENABLE KEYS */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

