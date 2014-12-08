/*
Navicat MySQL Data Transfer

Source Server         : localhost_3306
Source Server Version : 50173
Source Host           : localhost:3306
Source Database       : 20141110

Target Server Type    : MYSQL
Target Server Version : 50173
File Encoding         : 65001

Date: 2014-12-08 17:42:29
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for coord_newareaprojects
-- ----------------------------
DROP TABLE IF EXISTS `coord_newareaprojects`;
CREATE TABLE `coord_newareaprojects` (
  `ID` varchar(255) NOT NULL,
  `CityID` int(11) NOT NULL,
  `Name` varchar(255) DEFAULT NULL,
  `Result` bit(1) DEFAULT NULL,
  `County` varchar(255) DEFAULT NULL,
  `Note` varchar(255) DEFAULT NULL,
  `UpdateTime` date NOT NULL,
  `Visible` bit(1) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for coord_projects
-- ----------------------------
DROP TABLE IF EXISTS `coord_projects`;
CREATE TABLE `coord_projects` (
  `ID` varchar(255) NOT NULL,
  `CityID` int(11) NOT NULL,
  `Name` varchar(255) DEFAULT NULL,
  `Result` bit(1) DEFAULT NULL,
  `County` varchar(255) DEFAULT NULL,
  `Note` varchar(255) DEFAULT NULL,
  `UpdateTime` date NOT NULL,
  `Visible` bit(1) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for farmland
-- ----------------------------
DROP TABLE IF EXISTS `farmland`;
CREATE TABLE `farmland` (
  `ID` int(55) NOT NULL AUTO_INCREMENT,
  `ProjectID` varchar(255) NOT NULL,
  `Area` float(10,4) NOT NULL,
  `Type` int(11) NOT NULL,
  `Degree` int(11) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for projects
-- ----------------------------
DROP TABLE IF EXISTS `projects`;
CREATE TABLE `projects` (
  `ID` varchar(55) NOT NULL,
  `CityID` int(11) NOT NULL,
  `Name` varchar(127) DEFAULT NULL,
  `Result` tinyint(1) DEFAULT NULL,
  `Note` varchar(255) DEFAULT NULL,
  `County` varchar(55) DEFAULT NULL,
  `UpdateTime` datetime NOT NULL,
  `IsApplyDelete` bit(1) NOT NULL,
  `IsHasError` bit(1) NOT NULL,
  `IsShouldModify` bit(1) NOT NULL,
  `IsDecrease` bit(1) NOT NULL,
  `Area` float(10,4) DEFAULT NULL,
  `NewArea` float(10,4) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_PROJECT_ID` (`ID`),
  KEY `IX_PROJECT_CITY` (`CityID`),
  KEY `IX_PROJECT_UPDATETIME` (`UpdateTime`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for records
-- ----------------------------
DROP TABLE IF EXISTS `records`;
CREATE TABLE `records` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ProjectID` varchar(255) NOT NULL,
  `Type` int(11) NOT NULL,
  `CityID` int(11) NOT NULL,
  `IsError` bit(1) NOT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_RECORD_ID` (`ID`),
  KEY `IX_RECORD_PROJECTID` (`ProjectID`),
  KEY `IX_RECORD_CITYID` (`CityID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for reports
-- ----------------------------
DROP TABLE IF EXISTS `reports`;
CREATE TABLE `reports` (
  `ID` varchar(55) NOT NULL,
  `CityID` int(11) NOT NULL,
  `Type` int(11) NOT NULL,
  `Result` bit(1) DEFAULT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for seprojects
-- ----------------------------
DROP TABLE IF EXISTS `seprojects`;
CREATE TABLE `seprojects` (
  `ID` varchar(55) NOT NULL,
  `City` int(11) NOT NULL,
  `Name` varchar(127) DEFAULT NULL,
  `Result` tinyint(1) DEFAULT NULL,
  `County` varchar(55) DEFAULT NULL,
  `Note` varchar(255) DEFAULT NULL,
  `UpdateTime` datetime NOT NULL,
  `IsHasDoubt` bit(1) NOT NULL,
  `IsApplyDelete` bit(1) NOT NULL,
  `IsHasError` bit(1) NOT NULL,
  `IsPacket` bit(1) NOT NULL,
  `IsDescrease` bit(1) NOT NULL,
  `IsRelieve` bit(1) NOT NULL,
  `Area` float(10,4) DEFAULT NULL,
  `NewArea` float(10,4) DEFAULT NULL,
  `SurplusHookArea` float(10,4) DEFAULT NULL,
  `TrueHookArea` float(10,4) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for serecords
-- ----------------------------
DROP TABLE IF EXISTS `serecords`;
CREATE TABLE `serecords` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ProjectID` varchar(255) NOT NULL,
  `Type` int(11) NOT NULL,
  `City` int(11) NOT NULL,
  `IsError` bit(1) NOT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=18867 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for sereports
-- ----------------------------
DROP TABLE IF EXISTS `sereports`;
CREATE TABLE `sereports` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `City` int(11) NOT NULL,
  `Type` int(11) NOT NULL,
  `Result` bit(1) DEFAULT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for uploadfiles
-- ----------------------------
DROP TABLE IF EXISTS `uploadfiles`;
CREATE TABLE `uploadfiles` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `CityID` int(11) NOT NULL,
  `FileName` varchar(127) DEFAULT NULL,
  `CreateTime` datetime NOT NULL,
  `SavePath` varchar(127) DEFAULT NULL,
  `Type` int(11) NOT NULL DEFAULT '0',
  `State` int(1) NOT NULL DEFAULT '0',
  `ProcessMessage` varchar(1023) DEFAULT NULL,
  `Census` tinyint(1) NOT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_FILE_ID` (`ID`) USING BTREE,
  KEY `IX_FILE_CITY` (`CityID`)
) ENGINE=InnoDB AUTO_INCREMENT=75 DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for users
-- ----------------------------
DROP TABLE IF EXISTS `users`;
CREATE TABLE `users` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `Username` varchar(55) DEFAULT NULL,
  `Password` varchar(55) DEFAULT NULL,
  `Flag` tinyint(1) NOT NULL,
  `CityID` int(11) NOT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_USER_ID` (`ID`) USING BTREE,
  KEY `IX_USER_CITY` (`CityID`),
  KEY `IX_USERNAME` (`Username`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=utf8;
