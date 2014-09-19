/*
Navicat MySQL Data Transfer

Source Server         : local
Source Server Version : 50703
Source Host           : localhost:3306
Source Database       : lcchecker

Target Server Type    : MYSQL
Target Server Version : 50703
File Encoding         : 65001

Date: 2014-09-19 12:06:04
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
  `Visible` bit(1) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

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
  `UpdateTime` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `IsApplyDelete` bit(1) NOT NULL,
  `IsHasError` bit(1) NOT NULL,
  `IsShouldModify` bit(1) NOT NULL,
  `IsDecrease` bit(1) NOT NULL,
  `Area` decimal(10,0) DEFAULT NULL,
  `NewArea` decimal(10,0) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_PROJECT_ID` (`ID`),
  KEY `IX_PROJECT_CITY` (`CityID`),
  KEY `IX_PROJECT_UPDATETIME` (`UpdateTime`)
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
  `Proceeded` bit(1) NOT NULL DEFAULT b'0',
  PRIMARY KEY (`ID`),
  UNIQUE KEY `PK_FILE_ID` (`ID`) USING BTREE,
  KEY `IX_FILE_CITY` (`CityID`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;

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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;
