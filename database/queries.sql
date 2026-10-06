-- Read-only SELECT statements used by the Flask application.
-- The application binds user input through mysql-connector-python parameters.

-- Dashboard request totals by status.
SELECT
    COUNT(*) AS total_requests,
    SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'pending' THEN 1 ELSE 0 END) AS pending_requests,
    SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'assigned' THEN 1 ELSE 0 END) AS assigned_requests,
    SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'in progress'
              OR LOWER(COALESCE(status, '')) = 'in_progress' THEN 1 ELSE 0 END) AS in_progress_requests,
    SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'completed' THEN 1 ELSE 0 END) AS completed_requests,
    SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'closed' THEN 1 ELSE 0 END) AS closed_requests
FROM maintenance_requests;

-- Requests page: request, reporter, location, category, priority, and status.
SELECT mr.request_id, mr.request_number, mr.title, u.full_name AS user_name,
       b.building_name, r.room_number, c.category_name,
       p.priority_name, mr.status, mr.request_date
FROM maintenance_requests mr
LEFT JOIN users u ON u.user_id = mr.user_id
LEFT JOIN rooms r ON r.room_id = mr.room_id
LEFT JOIN buildings b ON b.building_id = r.building_id
LEFT JOIN categories c ON c.category_id = mr.category_id
LEFT JOIN priorities p ON p.priority_id = mr.priority_id
ORDER BY mr.request_date DESC;

-- Recent requests shown on the dashboard.
SELECT mr.request_id, mr.request_number, mr.title, mr.status, mr.request_date,
       u.full_name AS user_name,
       COALESCE(tech.technician_name, 'Unassigned') AS technician_name
FROM maintenance_requests mr
LEFT JOIN users u ON u.user_id = mr.user_id
LEFT JOIN assignments a ON a.request_id = mr.request_id
LEFT JOIN technicians tech ON tech.technician_id = a.technician_id
ORDER BY mr.request_date DESC
LIMIT 5;

-- Reference data used to populate the Create Request form.
SELECT user_id, full_name AS name
FROM users
ORDER BY full_name;

SELECT room_id, room_number, building_name
FROM rooms r
LEFT JOIN buildings b ON b.building_id = r.building_id
ORDER BY building_name, room_number;

SELECT category_id, category_name
FROM categories
ORDER BY category_name;

SELECT priority_id, priority_name
FROM priorities
ORDER BY priority_name;

-- Technician workload report.
SELECT tech.technician_name, tech.specialization,
       COUNT(a.assignment_id) AS assigned_requests
FROM technicians tech
LEFT JOIN assignments a ON a.technician_id = tech.technician_id
GROUP BY tech.technician_id, tech.technician_name, tech.specialization
ORDER BY assigned_requests DESC;

-- Building request-count report.
SELECT b.building_name, COUNT(mr.request_id) AS request_count
FROM buildings b
LEFT JOIN rooms r ON r.building_id = b.building_id
LEFT JOIN maintenance_requests mr ON mr.room_id = r.room_id
GROUP BY b.building_id, b.building_name
ORDER BY request_count DESC;

-- Average feedback rating shown on the dashboard.
SELECT ROUND(AVG(rating), 2) AS avg_rating
FROM feedback;
