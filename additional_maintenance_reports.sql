
-- Additional SQL reports
-- Project 41: Campus Facility Maintenance Request Management System

-- 1. Display all open maintenance requests
SELECT request_number, title, status, request_date
FROM maintenance_requests
WHERE status NOT IN ('Closed', 'Completed')
ORDER BY request_date ASC;

-- 2. Find the age of open maintenance requests
SELECT
    request_number,
    title,
    status,
    TIMESTAMPDIFF(DAY, request_date, CURRENT_TIMESTAMP)
        AS age_in_days
FROM maintenance_requests
WHERE status NOT IN ('Closed', 'Completed')
ORDER BY age_in_days DESC;

-- 3. Display active assignments for each technician
SELECT
    t.technician_name,
    COUNT(a.assignment_id) AS active_assignments
FROM technicians AS t
LEFT JOIN assignments AS a
    ON t.technician_id = a.technician_id
    AND a.assignment_status IN ('Assigned', 'In Progress')
GROUP BY t.technician_id, t.technician_name
ORDER BY active_assignments DESC;

-- 4. Calculate material usage and estimated cost
SELECT
    m.material_name,
    m.unit,
    SUM(mu.quantity) AS total_quantity_used,
    SUM(mu.quantity * m.unit_cost) AS estimated_total_cost
FROM material_usage AS mu
JOIN materials AS m
    ON mu.material_id = m.material_id
GROUP BY m.material_id, m.material_name, m.unit, m.unit_cost
ORDER BY estimated_total_cost DESC;
