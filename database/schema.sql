-- Structure-only export from campus_maintenance; no row data is included.
-- Import only into a new/empty database. Do not run over an existing project database.
CREATE DATABASE IF NOT EXISTS `campus_maintenance` CHARACTER SET utf8mb4;
USE `campus_maintenance`;

CREATE TABLE `buildings` (
  `building_id` int NOT NULL AUTO_INCREMENT,
  `building_name` varchar(100) NOT NULL,
  `location` varchar(150) NOT NULL,
  PRIMARY KEY (`building_id`),
  UNIQUE KEY `building_name` (`building_name`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `categories` (
  `category_id` int NOT NULL AUTO_INCREMENT,
  `category_name` varchar(50) NOT NULL,
  `description` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`category_id`),
  UNIQUE KEY `category_name` (`category_name`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `materials` (
  `material_id` int NOT NULL AUTO_INCREMENT,
  `material_name` varchar(100) NOT NULL,
  `unit` varchar(30) NOT NULL,
  `unit_cost` decimal(10,2) NOT NULL,
  PRIMARY KEY (`material_id`),
  UNIQUE KEY `material_name` (`material_name`),
  CONSTRAINT `materials_chk_1` CHECK ((`unit_cost` >= 0))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `priorities` (
  `priority_id` int NOT NULL AUTO_INCREMENT,
  `priority_name` varchar(30) NOT NULL,
  `response_hours` int NOT NULL,
  PRIMARY KEY (`priority_id`),
  UNIQUE KEY `priority_name` (`priority_name`),
  CONSTRAINT `priorities_chk_1` CHECK ((`response_hours` > 0))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `technicians` (
  `technician_id` int NOT NULL AUTO_INCREMENT,
  `technician_name` varchar(100) NOT NULL,
  `specialization` varchar(50) NOT NULL,
  `phone` varchar(15) DEFAULT NULL,
  `status` varchar(20) DEFAULT 'Available',
  PRIMARY KEY (`technician_id`),
  UNIQUE KEY `phone` (`phone`),
  CONSTRAINT `technicians_chk_1` CHECK ((`status` in (_cp850'Available',_cp850'Busy',_cp850'Inactive')))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `users` (
  `user_id` int NOT NULL AUTO_INCREMENT,
  `full_name` varchar(100) NOT NULL,
  `email` varchar(120) NOT NULL,
  `phone` varchar(15) DEFAULT NULL,
  `role` varchar(30) NOT NULL DEFAULT 'Student',
  PRIMARY KEY (`user_id`),
  UNIQUE KEY `email` (`email`),
  CONSTRAINT `users_chk_1` CHECK ((`role` in (_cp850'Student',_cp850'Faculty',_cp850'Staff',_cp850'Admin')))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `rooms` (
  `room_id` int NOT NULL AUTO_INCREMENT,
  `building_id` int NOT NULL,
  `room_number` varchar(20) NOT NULL,
  `room_type` varchar(50) DEFAULT 'Classroom',
  PRIMARY KEY (`room_id`),
  UNIQUE KEY `building_id` (`building_id`,`room_number`),
  CONSTRAINT `rooms_ibfk_1` FOREIGN KEY (`building_id`) REFERENCES `buildings` (`building_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `maintenance_requests` (
  `request_id` int NOT NULL AUTO_INCREMENT,
  `request_number` varchar(30) NOT NULL,
  `user_id` int NOT NULL,
  `room_id` int NOT NULL,
  `category_id` int NOT NULL,
  `priority_id` int NOT NULL,
  `title` varchar(150) NOT NULL,
  `description` text NOT NULL,
  `request_date` datetime DEFAULT CURRENT_TIMESTAMP,
  `status` varchar(30) NOT NULL DEFAULT 'Pending',
  `assigned_date` datetime DEFAULT NULL,
  `closed_date` datetime DEFAULT NULL,
  `closure_details` varchar(500) DEFAULT NULL,
  PRIMARY KEY (`request_id`),
  UNIQUE KEY `request_number` (`request_number`),
  KEY `user_id` (`user_id`),
  KEY `room_id` (`room_id`),
  KEY `category_id` (`category_id`),
  KEY `priority_id` (`priority_id`),
  CONSTRAINT `maintenance_requests_ibfk_1` FOREIGN KEY (`user_id`) REFERENCES `users` (`user_id`),
  CONSTRAINT `maintenance_requests_ibfk_2` FOREIGN KEY (`room_id`) REFERENCES `rooms` (`room_id`),
  CONSTRAINT `maintenance_requests_ibfk_3` FOREIGN KEY (`category_id`) REFERENCES `categories` (`category_id`),
  CONSTRAINT `maintenance_requests_ibfk_4` FOREIGN KEY (`priority_id`) REFERENCES `priorities` (`priority_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `assignments` (
  `assignment_id` int NOT NULL AUTO_INCREMENT,
  `request_id` int NOT NULL,
  `technician_id` int NOT NULL,
  `assigned_at` datetime DEFAULT CURRENT_TIMESTAMP,
  `assignment_status` varchar(30) DEFAULT 'Assigned',
  PRIMARY KEY (`assignment_id`),
  UNIQUE KEY `request_id` (`request_id`),
  KEY `technician_id` (`technician_id`),
  CONSTRAINT `assignments_ibfk_1` FOREIGN KEY (`request_id`) REFERENCES `maintenance_requests` (`request_id`),
  CONSTRAINT `assignments_ibfk_2` FOREIGN KEY (`technician_id`) REFERENCES `technicians` (`technician_id`),
  CONSTRAINT `assignments_chk_1` CHECK ((`assignment_status` in (_cp850'Assigned',_cp850'In Progress',_cp850'Completed')))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `feedback` (
  `feedback_id` int NOT NULL AUTO_INCREMENT,
  `request_id` int NOT NULL,
  `rating` int NOT NULL,
  `comments` varchar(500) DEFAULT NULL,
  `feedback_date` datetime DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`feedback_id`),
  UNIQUE KEY `request_id` (`request_id`),
  CONSTRAINT `feedback_ibfk_1` FOREIGN KEY (`request_id`) REFERENCES `maintenance_requests` (`request_id`),
  CONSTRAINT `feedback_chk_1` CHECK ((`rating` between 1 and 5))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `material_usage` (
  `usage_id` int NOT NULL AUTO_INCREMENT,
  `request_id` int NOT NULL,
  `material_id` int NOT NULL,
  `quantity` decimal(10,2) NOT NULL,
  `used_date` datetime DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`usage_id`),
  KEY `request_id` (`request_id`),
  KEY `material_id` (`material_id`),
  CONSTRAINT `material_usage_ibfk_1` FOREIGN KEY (`request_id`) REFERENCES `maintenance_requests` (`request_id`),
  CONSTRAINT `material_usage_ibfk_2` FOREIGN KEY (`material_id`) REFERENCES `materials` (`material_id`),
  CONSTRAINT `material_usage_chk_1` CHECK ((`quantity` > 0))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `status_history` (
  `history_id` int NOT NULL AUTO_INCREMENT,
  `request_id` int NOT NULL,
  `old_status` varchar(30) DEFAULT NULL,
  `new_status` varchar(30) NOT NULL,
  `changed_at` datetime DEFAULT CURRENT_TIMESTAMP,
  `changed_by` int NOT NULL,
  PRIMARY KEY (`history_id`),
  KEY `request_id` (`request_id`),
  KEY `changed_by` (`changed_by`),
  CONSTRAINT `status_history_ibfk_1` FOREIGN KEY (`request_id`) REFERENCES `maintenance_requests` (`request_id`),
  CONSTRAINT `status_history_ibfk_2` FOREIGN KEY (`changed_by`) REFERENCES `users` (`user_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE TABLE `work_logs` (
  `work_log_id` int NOT NULL AUTO_INCREMENT,
  `request_id` int NOT NULL,
  `technician_id` int NOT NULL,
  `work_date` datetime DEFAULT CURRENT_TIMESTAMP,
  `work_description` varchar(500) NOT NULL,
  `hours_spent` decimal(5,2) NOT NULL,
  PRIMARY KEY (`work_log_id`),
  KEY `request_id` (`request_id`),
  KEY `technician_id` (`technician_id`),
  CONSTRAINT `work_logs_ibfk_1` FOREIGN KEY (`request_id`) REFERENCES `maintenance_requests` (`request_id`),
  CONSTRAINT `work_logs_ibfk_2` FOREIGN KEY (`technician_id`) REFERENCES `technicians` (`technician_id`),
  CONSTRAINT `work_logs_chk_1` CHECK ((`hours_spent` >= 0))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `pending_requests` AS select `mr`.`request_number` AS `request_number`,`mr`.`title` AS `title`,`u`.`full_name` AS `reported_by`,`c`.`category_name` AS `category_name`,`p`.`priority_name` AS `priority_name`,`mr`.`status` AS `status`,`b`.`building_name` AS `building_name`,`r`.`room_number` AS `room_number` from (((((`maintenance_requests` `mr` join `users` `u` on((`mr`.`user_id` = `u`.`user_id`))) join `categories` `c` on((`mr`.`category_id` = `c`.`category_id`))) join `priorities` `p` on((`mr`.`priority_id` = `p`.`priority_id`))) join `rooms` `r` on((`mr`.`room_id` = `r`.`room_id`))) join `buildings` `b` on((`r`.`building_id` = `b`.`building_id`))) where (`mr`.`status` not in ('Closed','Completed'));
