/*
Navicat MySQL Data Transfer

Source Server         : localhost_3306
Source Server Version : 50173
Source Host           : localhost:3306
Source Database       : lc

Target Server Type    : MYSQL
Target Server Version : 50173
File Encoding         : 65001

Date: 2014-09-21 20:20:27
*/

SET FOREIGN_KEY_CHECKS=0;

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
  `Visible` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for projects
-- ----------------------------
DROP TABLE IF EXISTS `projects`;
CREATE TABLE `projects` (
  `ID` varchar(55) NOT NULL,
  `CityID` int(11) NOT NULL,
  `Name` varchar(255) DEFAULT NULL,
  `Result` tinyint(1) DEFAULT NULL,
  `County` varchar(255) DEFAULT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  `UpdateTime` datetime NOT NULL,
  `IsApplyDelete` tinyint(1) DEFAULT NULL,
  `IsHasError` tinyint(1) DEFAULT NULL,
  `IsShouldModify` tinyint(1) DEFAULT NULL,
  `IsDecrease` tinyint(1) DEFAULT NULL,
  `Area` double(255,0) DEFAULT NULL,
  `NewArea` double(255,0) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for records
-- ----------------------------
DROP TABLE IF EXISTS `records`;
CREATE TABLE `records` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ProjectID` varchar(55) NOT NULL,
  `Type` int(11) NOT NULL,
  `CityID` int(11) NOT NULL,
  `IsError` bit(1) DEFAULT NULL,
  `Note` varchar(1023) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=4287 DEFAULT CHARSET=utf8;

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
-- Table structure for uploadfiles
-- ----------------------------
DROP TABLE IF EXISTS `uploadfiles`;
CREATE TABLE `uploadfiles` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `CityID` int(11) NOT NULL,
  `FileName` varchar(55) DEFAULT NULL,
  `CreateTime` datetime NOT NULL,
  `SavePath` varchar(55) DEFAULT NULL,
  `Type` int(11) NOT NULL DEFAULT '0',
  `Proceeded` bit(1) DEFAULT b'0',
  `State` int(11) NOT NULL,
  `ProcessMessage` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_FILE_ID` (`ID`) USING BTREE,
  KEY `IX_FILE_CITY` (`CityID`)
) ENGINE=InnoDB AUTO_INCREMENT=85 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;
