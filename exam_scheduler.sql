-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1:3306
-- Generation Time: May 17, 2025 at 02:17 PM
-- Server version: 9.1.0
-- PHP Version: 8.3.14

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `exam_scheduler`
--

-- --------------------------------------------------------

--
-- Table structure for table `rooms`
--

DROP TABLE IF EXISTS `rooms`;
CREATE TABLE IF NOT EXISTS `rooms` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(50) NOT NULL,
  `type` varchar(20) NOT NULL,
  `capacity` int NOT NULL,
  `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=11 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data for table `rooms`
--

INSERT INTO `rooms` (`id`, `name`, `type`, `capacity`, `created_at`, `updated_at`) VALUES
(1, '101', 'Classroom', 30, '2025-05-16 16:40:14', '2025-05-16 16:40:14'),
(2, '207', 'Classroom', 30, '2025-05-16 17:03:56', '2025-05-16 17:03:56'),
(3, '307', 'Classroom', 30, '2025-05-16 17:04:06', '2025-05-16 17:04:06'),
(4, '208', 'Classroom', 30, '2025-05-16 17:04:19', '2025-05-16 17:04:19'),
(5, '113', 'Lab', 30, '2025-05-16 17:04:29', '2025-05-16 17:04:29'),
(6, '205', 'Lab', 30, '2025-05-16 17:04:38', '2025-05-16 17:04:38'),
(7, '215', 'Lab', 30, '2025-05-16 17:04:51', '2025-05-16 17:04:51'),
(8, '6', 'Lab', 30, '2025-05-16 17:05:00', '2025-05-16 17:05:00'),
(9, '7', 'Lab', 30, '2025-05-16 17:05:09', '2025-05-16 17:05:09'),
(10, '8', 'Lab', 30, '2025-05-16 17:05:18', '2025-05-16 17:05:18');

-- --------------------------------------------------------

--
-- Table structure for table `schedules`
--

DROP TABLE IF EXISTS `schedules`;
CREATE TABLE IF NOT EXISTS `schedules` (
  `id` int NOT NULL AUTO_INCREMENT,
  `name` varchar(100) NOT NULL,
  `semester` varchar(20) NOT NULL,
  `exam_type` varchar(20) NOT NULL,
  `start_date` date NOT NULL,
  `config` text NOT NULL,
  `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data for table `schedules`
--

INSERT INTO `schedules` (`id`, `name`, `semester`, `exam_type`, `start_date`, `config`, `created_at`, `updated_at`) VALUES
(1, 'Midsem - Summer', '2', 'Regular', '2025-05-17', '{\"working_hours\": {\"start\": \"09:00\", \"end\": \"17:00\"}, \"exam_duration\": 180, \"break_duration\": 30}', '2025-05-17 12:46:56', '2025-05-17 12:46:56');

-- --------------------------------------------------------

--
-- Table structure for table `schedule_items`
--

DROP TABLE IF EXISTS `schedule_items`;
CREATE TABLE IF NOT EXISTS `schedule_items` (
  `id` int NOT NULL AUTO_INCREMENT,
  `schedule_id` int NOT NULL,
  `subject_id` int NOT NULL,
  `room_id` int NOT NULL,
  `exam_date` date NOT NULL,
  `start_time` varchar(20) NOT NULL,
  `end_time` varchar(20) NOT NULL,
  PRIMARY KEY (`id`),
  KEY `schedule_id` (`schedule_id`),
  KEY `subject_id` (`subject_id`),
  KEY `room_id` (`room_id`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data for table `schedule_items`
--

INSERT INTO `schedule_items` (`id`, `schedule_id`, `subject_id`, `room_id`, `exam_date`, `start_time`, `end_time`) VALUES
(1, 1, 3, 5, '2025-05-17', '02:00 PM', '05:00 PM'),
(2, 1, 4, 1, '2025-05-19', '09:00 AM', '12:00 PM'),
(3, 1, 5, 2, '2025-05-20', '09:00 AM', '12:00 PM');

-- --------------------------------------------------------

--
-- Table structure for table `subjects`
--

DROP TABLE IF EXISTS `subjects`;
CREATE TABLE IF NOT EXISTS `subjects` (
  `id` int NOT NULL AUTO_INCREMENT,
  `code` varchar(20) NOT NULL,
  `name` varchar(100) NOT NULL,
  `type` varchar(20) NOT NULL,
  `semester` varchar(20) NOT NULL,
  `difficulty` varchar(20) NOT NULL,
  `duration` int NOT NULL,
  `created_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` timestamp NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=8 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

--
-- Dumping data for table `subjects`
--

INSERT INTO `subjects` (`id`, `code`, `name`, `type`, `semester`, `difficulty`, `duration`, `created_at`, `updated_at`) VALUES
(1, '1', 'JAVA', 'Theory', '4', 'Medium', 120, '2025-05-16 16:07:57', '2025-05-16 16:07:57'),
(2, '2', 'SDWT', 'Theory', '2', 'Medium', 120, '2025-05-16 16:08:18', '2025-05-16 16:34:55'),
(3, '3', 'WDT-2', 'Practical', '2', 'Medium', 120, '2025-05-16 17:02:44', '2025-05-16 17:02:44'),
(4, '4', 'DBMS-1', 'Theory', '2', 'Medium', 120, '2025-05-16 17:03:23', '2025-05-16 17:03:23'),
(5, '5', 'CS-2', 'Theory', '2', 'Medium', 120, '2025-05-16 17:03:38', '2025-05-16 17:03:38'),
(6, '6', 'DBMS-2', 'Practical', '4', 'Medium', 120, '2025-05-16 18:21:17', '2025-05-16 18:21:17'),
(7, '7', 'Python', 'Regular', '1', 'Medium', 0, '2025-05-16 19:33:21', '2025-05-16 19:33:21');
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
