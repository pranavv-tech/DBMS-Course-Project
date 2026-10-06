import os
from datetime import datetime

import mysql.connector
from flask import Flask, flash, redirect, render_template, request, url_for
from mysql.connector import Error

from config import get_db_config

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY") or os.urandom(32)


def get_db_connection():
    try:
        return mysql.connector.connect(**get_db_config())
    except Error:
        app.logger.exception("Could not connect to the configured MySQL database")
        raise


def db_error_message(error):
    app.logger.exception("Database operation failed: %s", error)
    message = str(error)
    if "Duplicate entry" in message and "request_id" in message:
        return "This request has already been assigned or recorded."
    if "Duplicate entry" in message and "feedback" in message:
        return "Feedback for this request already exists."
    if "foreign key" in message.lower():
        return "A referenced database record could not be found. Please check the selected values."
    if "Data too long" in message:
        return "One or more field values are too long for the database."
    return "A database error occurred. Please try again."


def fetch_all(query, params=()):
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute(query, params)
        return cursor.fetchall()
    finally:
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()


def fetch_one(query, params=()):
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute(query, params)
        return cursor.fetchone()
    finally:
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()


def execute_write(query, params=()):
    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        return cursor.lastrowid
    except Error as exc:
        if conn is not None:
            conn.rollback()
        raise exc
    finally:
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()


def get_admin_user_id(cursor):
    cursor.execute("SELECT user_id FROM users WHERE role = %s ORDER BY user_id LIMIT 1", ("Admin",))
    admin = cursor.fetchone()
    if not admin:
        raise ValueError("No administrator account exists in users.")
    return admin[0]


def next_request_number(cursor):
    year = datetime.now().year
    prefix = f"REQ-{year}-"
    cursor.execute(
        "SELECT request_number FROM maintenance_requests WHERE request_number LIKE %s FOR UPDATE",
        (prefix + "%",),
    )
    numbers = []
    for (number,) in cursor.fetchall():
        suffix = number.rsplit("-", 1)[-1]
        if suffix.isdigit():
            numbers.append(int(suffix))
    return f"{prefix}{max(numbers, default=0) + 1:03d}"


def get_request_by_id(request_id):
    query = """
        SELECT mr.*, u.full_name AS user_name, b.building_name, r.room_number,
               c.category_name, p.priority_name, tech.technician_name,
               CASE
                   WHEN mr.status IS NULL OR mr.status = '' THEN 'Pending'
                   ELSE mr.status
               END AS current_status
        FROM maintenance_requests mr
        LEFT JOIN users u ON u.user_id = mr.user_id
        LEFT JOIN rooms r ON r.room_id = mr.room_id
        LEFT JOIN buildings b ON b.building_id = r.building_id
        LEFT JOIN categories c ON c.category_id = mr.category_id
        LEFT JOIN priorities p ON p.priority_id = mr.priority_id
        LEFT JOIN assignments a ON a.request_id = mr.request_id
        LEFT JOIN technicians tech ON tech.technician_id = a.technician_id
        WHERE mr.request_id = %s
    """
    return fetch_one(query, (request_id,))


def get_dashboard_stats():
    query = """
        SELECT
            COUNT(*) AS total_requests,
            SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'pending' THEN 1 ELSE 0 END) AS pending_requests,
            SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'assigned' THEN 1 ELSE 0 END) AS assigned_requests,
            SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'in progress' OR LOWER(COALESCE(status, '')) = 'in_progress' THEN 1 ELSE 0 END) AS in_progress_requests,
            SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'completed' THEN 1 ELSE 0 END) AS completed_requests,
            SUM(CASE WHEN LOWER(COALESCE(status, '')) = 'closed' THEN 1 ELSE 0 END) AS closed_requests
        FROM maintenance_requests
    """
    stats = fetch_one(query)
    if not stats:
        stats = {"total_requests": 0, "pending_requests": 0, "assigned_requests": 0, "in_progress_requests": 0, "completed_requests": 0, "closed_requests": 0}
    for key in stats:
        if stats[key] is None:
            stats[key] = 0

    technicians_count = fetch_one("SELECT COUNT(*) AS total FROM technicians")
    average_rating = fetch_one("SELECT ROUND(AVG(rating), 2) AS avg_rating FROM feedback")

    recent_requests = fetch_all(
        """
        SELECT mr.request_id, mr.request_number, mr.title, mr.status, mr.request_date, u.full_name AS user_name,
               COALESCE(tech.technician_name, 'Unassigned') AS technician_name
        FROM maintenance_requests mr
        LEFT JOIN users u ON u.user_id = mr.user_id
        LEFT JOIN assignments a ON a.request_id = mr.request_id
        LEFT JOIN technicians tech ON tech.technician_id = a.technician_id
        ORDER BY mr.request_date DESC
        LIMIT 5
        """
    )

    return {
        "stats": stats,
        "technicians_count": technicians_count["total"] if technicians_count else 0,
        "average_rating": average_rating["avg_rating"] if average_rating and average_rating["avg_rating"] is not None else 0,
        "recent_requests": recent_requests,
    }


def get_users():
    return fetch_all("SELECT user_id, full_name AS name FROM users ORDER BY full_name")


def get_technicians():
    return fetch_all(
        """
        SELECT tech.technician_id, tech.technician_name AS name, tech.specialization, tech.status
        FROM technicians tech
        ORDER BY tech.technician_name
        """
    )


def get_materials():
    return fetch_all("SELECT material_id, material_name, unit, unit_cost FROM materials ORDER BY material_name")


def get_status_options():
    return ["Pending", "Assigned", "In Progress", "Completed", "Closed"]


@app.route("/")
def index():
    dashboard = get_dashboard_stats()
    return render_template("index.html", dashboard=dashboard)


@app.route("/requests")
def requests_list():
    search = request.args.get("search", "").strip()
    status_filter = request.args.get("status", "")
    category_filter = request.args.get("category", "")
    priority_filter = request.args.get("priority", "")

    query = """
        SELECT mr.request_id, mr.request_number, mr.title, u.full_name AS user_name,
               b.building_name, r.room_number, c.category_name,
               p.priority_name, mr.status, mr.request_date
        FROM maintenance_requests mr
        LEFT JOIN users u ON u.user_id = mr.user_id
        LEFT JOIN rooms r ON r.room_id = mr.room_id
        LEFT JOIN buildings b ON b.building_id = r.building_id
        LEFT JOIN categories c ON c.category_id = mr.category_id
        LEFT JOIN priorities p ON p.priority_id = mr.priority_id
        WHERE 1=1
    """
    params = []

    if search:
        query += " AND (mr.request_number LIKE %s OR mr.title LIKE %s)"
        like = f"%{search}%"
        params.extend([like, like])
    if status_filter:
        query += " AND mr.status = %s"
        params.append(status_filter)
    if category_filter:
        query += " AND mr.category_id = %s"
        params.append(category_filter)
    if priority_filter:
        query += " AND mr.priority_id = %s"
        params.append(priority_filter)

    query += " ORDER BY mr.request_date DESC"

    rows = fetch_all(query, tuple(params))
    categories = fetch_all("SELECT category_id, category_name FROM categories ORDER BY category_name")
    priorities = fetch_all("SELECT priority_id, priority_name FROM priorities ORDER BY priority_name")
    statuses = ["Pending", "Assigned", "In Progress", "Completed", "Closed"]
    return render_template(
        "requests.html",
        requests=rows,
        categories=categories,
        priorities=priorities,
        statuses=statuses,
        search=search,
        status_filter=status_filter,
        category_filter=category_filter,
        priority_filter=priority_filter,
    )


@app.route("/requests/new", methods=["GET", "POST"])
def add_request():
    if request.method == "POST":
        user_id = request.form.get("user_id")
        room_id = request.form.get("room_id")
        category_id = request.form.get("category_id")
        priority_id = request.form.get("priority_id")
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()

        if not all([user_id, room_id, category_id, priority_id, title, description]):
            flash("Please complete all required fields.", "error")
        else:
            conn = None
            cursor = None
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                conn.start_transaction()
                request_number = next_request_number(cursor)
                cursor.execute(
                    """
                    INSERT INTO maintenance_requests
                    (request_number, title, description, user_id, room_id, category_id, priority_id)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """,
                    (request_number, title, description, user_id, room_id, category_id, priority_id),
                )
                request_id = cursor.lastrowid
                cursor.execute(
                    "INSERT INTO status_history (request_id, old_status, new_status, changed_by) VALUES (%s, NULL, 'Pending', %s)",
                    (request_id, user_id),
                )
                conn.commit()
                flash(f"Request {request_number} created successfully.", "success")
                return redirect(url_for("request_details", request_id=request_id))
            except Error as exc:
                if conn is not None:
                    conn.rollback()
                flash(db_error_message(exc), "error")
            finally:
                if cursor is not None:
                    cursor.close()
                if conn is not None:
                    conn.close()

    users = get_users()
    rooms = fetch_all("SELECT room_id, room_number, building_name FROM rooms r LEFT JOIN buildings b ON b.building_id = r.building_id ORDER BY building_name, room_number")
    categories = fetch_all("SELECT category_id, category_name FROM categories ORDER BY category_name")
    priorities = fetch_all("SELECT priority_id, priority_name FROM priorities ORDER BY priority_name")
    return render_template("add_request.html", users=users, rooms=rooms, categories=categories, priorities=priorities)


@app.route("/requests/<int:request_id>")
def request_details(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    history = fetch_all(
        """
        SELECT sh.*, u.full_name AS changed_by_name
        FROM status_history sh
        LEFT JOIN users u ON u.user_id = sh.changed_by
        WHERE sh.request_id = %s
        ORDER BY sh.changed_at ASC
        """,
        (request_id,),
    )
    work_logs = fetch_all(
        """
        SELECT wl.*, tech.technician_name
        FROM work_logs wl
        LEFT JOIN technicians tech ON tech.technician_id = wl.technician_id
        WHERE wl.request_id = %s
        ORDER BY wl.work_date DESC
        """,
        (request_id,),
    )
    materials_used = fetch_all(
        """
        SELECT mu.*, m.material_name, m.unit, m.unit_cost,
               (mu.quantity * m.unit_cost) AS total_cost
        FROM material_usage mu
        LEFT JOIN materials m ON m.material_id = mu.material_id
        WHERE mu.request_id = %s
        ORDER BY m.material_name
        """,
        (request_id,),
    )
    feedback = fetch_one("SELECT * FROM feedback WHERE request_id = %s", (request_id,))
    technician_assignment = fetch_one(
        """
        SELECT a.*, tech.technician_name, tech.specialization
        FROM assignments a
        LEFT JOIN technicians tech ON tech.technician_id = a.technician_id
        WHERE a.request_id = %s
        """,
        (request_id,),
    )

    material_total = sum(float(row["total_cost"]) if row.get("total_cost") is not None else 0 for row in materials_used)
    return render_template(
        "request_details.html",
        request=request_row,
        history=history,
        work_logs=work_logs,
        materials_used=materials_used,
        feedback=feedback,
        technician_assignment=technician_assignment,
        material_total=material_total,
    )


@app.route("/requests/<int:request_id>/assign", methods=["GET", "POST"])
def assign_technician(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    if request.method == "POST":
        technician_id = request.form.get("technician_id")
        if not technician_id:
            flash("Please select a technician.", "error")
        else:
            conn = None
            cursor = None
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                conn.start_transaction()
                cursor.execute("SELECT request_id FROM assignments WHERE request_id = %s", (request_id,))
                if cursor.fetchone():
                    flash("This request already has an assigned technician.", "error")
                    conn.rollback()
                    return redirect(url_for("request_details", request_id=request_id))
                admin_user_id = get_admin_user_id(cursor)
                cursor.execute(
                    "INSERT INTO assignments (request_id, technician_id, assigned_at, assignment_status) VALUES (%s, %s, NOW(), 'Assigned')",
                    (request_id, technician_id),
                )
                cursor.execute(
                    "UPDATE maintenance_requests SET assigned_date = NOW(), status = 'Assigned' WHERE request_id = %s",
                    (request_id,),
                )
                cursor.execute("UPDATE technicians SET status = 'Busy' WHERE technician_id = %s", (technician_id,))
                cursor.execute(
                    "INSERT INTO status_history (request_id, old_status, new_status, changed_by) VALUES (%s, %s, 'Assigned', %s)",
                    (request_id, request_row.get("status") or "Pending", admin_user_id),
                )
                conn.commit()
                flash("Technician assigned successfully.", "success")
                return redirect(url_for("request_details", request_id=request_id))
            except Error as exc:
                if conn is not None:
                    conn.rollback()
                flash(db_error_message(exc), "error")
            finally:
                if cursor is not None:
                    cursor.close()
                if conn is not None:
                    conn.close()

    technicians = get_technicians()
    return render_template("assign_technician.html", request=request_row, technicians=technicians)


@app.route("/requests/<int:request_id>/status", methods=["GET", "POST"])
def update_status(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    if request.method == "POST":
        new_status = request.form.get("status")
        current_status = (request_row.get("status") or "Pending").strip()
        allowed = {
            "Pending": ["Assigned", "In Progress"],
            "Assigned": ["In Progress"],
            "In Progress": ["Completed"],
            "Completed": ["Closed"],
        }
        if new_status not in allowed.get(current_status, []) and current_status != new_status:
            flash("This status change is not allowed for the current request state.", "error")
            return redirect(url_for("update_status", request_id=request_id))
        if new_status == "Closed" and not request.form.get("closure_details", "").strip():
            flash("Closure details are required before a request can be marked as Closed.", "error")
            return redirect(url_for("update_status", request_id=request_id))

        conn = None
        cursor = None
        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            conn.start_transaction()
            admin_user_id = get_admin_user_id(cursor)
            update_query = "UPDATE maintenance_requests SET status = %s, closed_date = %s, closure_details = %s WHERE request_id = %s"
            closed_date = None
            closure_details = request.form.get("closure_details", "").strip() or None
            if new_status == "Closed":
                closed_date = request.form.get("closed_date") or datetime.now().strftime("%Y-%m-%d")
            else:
                if current_status == "Completed" and new_status != "Closed":
                    closed_date = None
                    closure_details = None
            cursor.execute(update_query, (new_status, closed_date, closure_details, request_id))
            cursor.execute(
                "INSERT INTO status_history (request_id, old_status, new_status, changed_by) VALUES (%s, %s, %s, %s)",
                (request_id, current_status, new_status, admin_user_id),
            )
            assignment_status = {"Assigned": "Assigned", "In Progress": "In Progress", "Completed": "Completed", "Closed": "Completed"}.get(new_status)
            if assignment_status:
                cursor.execute(
                    "UPDATE assignments SET assignment_status = %s WHERE request_id = %s",
                    (assignment_status, request_id),
                )
            if new_status in ["Completed", "Closed"]:
                cursor.execute(
                    "UPDATE technicians tech JOIN assignments a ON a.technician_id = tech.technician_id "
                    "SET tech.status = 'Available' WHERE a.request_id = %s",
                    (request_id,),
                )
            conn.commit()
            flash("Status updated successfully.", "success")
            return redirect(url_for("request_details", request_id=request_id))
        except Error as exc:
            if conn is not None:
                conn.rollback()
            flash(db_error_message(exc), "error")
        finally:
            if cursor is not None:
                cursor.close()
            if conn is not None:
                conn.close()

    return render_template("update_status.html", request=request_row, statuses=get_status_options())


@app.route("/requests/<int:request_id>/work", methods=["GET", "POST"])
def add_work_log(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    if request.method == "POST":
        technician_id = request.form.get("technician_id")
        description = request.form.get("work_description", "").strip()
        hours_spent = request.form.get("hours_spent", "0").strip()
        try:
            hours = float(hours_spent)
        except ValueError:
            flash("Hours spent must be a valid number.", "error")
            return redirect(url_for("add_work_log", request_id=request_id))
        if not technician_id or not description:
            flash("Technician and work description are required.", "error")
        elif hours < 0:
            flash("Hours spent cannot be negative.", "error")
        else:
            try:
                execute_write(
                    """
                    INSERT INTO work_logs (request_id, technician_id, work_description, hours_spent, work_date)
                    VALUES (%s, %s, %s, %s, NOW())
                    """,
                    (request_id, technician_id, description, hours),
                )
                flash("Work log added successfully.", "success")
                return redirect(url_for("request_details", request_id=request_id))
            except Error as exc:
                flash(db_error_message(exc), "error")

    technicians = get_technicians()
    return render_template("add_work.html", request=request_row, technicians=technicians)


@app.route("/requests/<int:request_id>/material", methods=["GET", "POST"])
@app.route("/requests/<int:request_id>/materials", methods=["GET", "POST"])
def add_material_usage(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    if request.method == "POST":
        material_id = request.form.get("material_id")
        quantity = request.form.get("quantity", "0").strip()
        try:
            quantity_value = float(quantity)
        except ValueError:
            flash("Quantity must be a valid numeric value.", "error")
            return redirect(url_for("add_material_usage", request_id=request_id))
        if not material_id:
            flash("Please select a material.", "error")
        elif quantity_value <= 0:
            flash("Quantity must be greater than zero.", "error")
        else:
            try:
                execute_write(
                    "INSERT INTO material_usage (request_id, material_id, quantity, used_date) VALUES (%s, %s, %s, NOW())",
                    (request_id, material_id, quantity_value),
                )
                flash("Material usage added successfully.", "success")
                return redirect(url_for("request_details", request_id=request_id))
            except Error as exc:
                flash(db_error_message(exc), "error")

    materials = get_materials()
    return render_template("add_material.html", request=request_row, materials=materials)


@app.route("/requests/<int:request_id>/close", methods=["GET", "POST"])
def close_request(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    if request.method == "POST":
        closure_details = request.form.get("closure_details", "").strip()
        closed_date = request.form.get("closed_date") or datetime.now().strftime("%Y-%m-%d")
        current_status = (request_row.get("status") or "Pending").strip()
        if current_status != "Completed":
            flash("A request must be completed before it can be closed.", "error")
            return redirect(url_for("close_request", request_id=request_id))
        if not closure_details:
            flash("Closure details are required.", "error")
        else:
            conn = None
            cursor = None
            try:
                conn = get_db_connection()
                cursor = conn.cursor()
                conn.start_transaction()
                admin_user_id = get_admin_user_id(cursor)
                cursor.execute(
                    "UPDATE maintenance_requests SET status = 'Closed', closed_date = %s, closure_details = %s WHERE request_id = %s",
                    (closed_date, closure_details, request_id),
                )
                cursor.execute(
                    "INSERT INTO status_history (request_id, old_status, new_status, changed_by) VALUES (%s, 'Completed', 'Closed', %s)",
                    (request_id, admin_user_id),
                )
                cursor.execute(
                    "UPDATE assignments SET assignment_status = 'Completed' WHERE request_id = %s",
                    (request_id,),
                )
                cursor.execute(
                    "UPDATE technicians tech JOIN assignments a ON a.technician_id = tech.technician_id "
                    "SET tech.status = 'Available' WHERE a.request_id = %s",
                    (request_id,),
                )
                conn.commit()
                flash("Request closed successfully.", "success")
                return redirect(url_for("request_details", request_id=request_id))
            except Error as exc:
                if conn is not None:
                    conn.rollback()
                flash(db_error_message(exc), "error")
            finally:
                if cursor is not None:
                    cursor.close()
                if conn is not None:
                    conn.close()

    return render_template("close_request.html", request=request_row)


@app.route("/requests/<int:request_id>/feedback", methods=["GET", "POST"])
def add_feedback(request_id):
    request_row = get_request_by_id(request_id)
    if not request_row:
        flash("The requested maintenance record could not be found.", "error")
        return redirect(url_for("requests_list"))

    current_status = (request_row.get("status") or "Pending").strip()
    if current_status not in ["Completed", "Closed"]:
        flash("Feedback can only be submitted for completed or closed requests.", "error")
        return redirect(url_for("request_details", request_id=request_id))

    if request.method == "POST":
        rating = request.form.get("rating")
        comments = request.form.get("comments", "").strip()
        try:
            rating_value = int(rating)
        except (TypeError, ValueError):
            flash("Rating must be a number from 1 to 5.", "error")
            return redirect(url_for("add_feedback", request_id=request_id))
        if rating_value < 1 or rating_value > 5:
            flash("Rating must be between 1 and 5.", "error")
            return redirect(url_for("add_feedback", request_id=request_id))
        if not comments:
            flash("Comments are required for feedback.", "error")
            return redirect(url_for("add_feedback", request_id=request_id))
        try:
            execute_write(
                "INSERT INTO feedback (request_id, rating, comments) VALUES (%s, %s, %s)",
                (request_id, rating_value, comments),
            )
            flash("Feedback submitted successfully.", "success")
            return redirect(url_for("request_details", request_id=request_id))
        except Error as exc:
            flash(db_error_message(exc), "error")

    return render_template("feedback.html", request=request_row)


@app.route("/technicians")
def technicians():
    technicians_list = get_technicians()
    return render_template("technicians.html", technicians=technicians_list)


@app.route("/reports")
def reports():
    reports_data = {
        "pending_requests": fetch_all(
            """
            SELECT mr.request_number, mr.title, c.category_name, p.priority_name,
                   b.building_name, r.room_number, mr.status
            FROM maintenance_requests mr
            LEFT JOIN rooms r ON r.room_id = mr.room_id
            LEFT JOIN buildings b ON b.building_id = r.building_id
            LEFT JOIN categories c ON c.category_id = mr.category_id
            LEFT JOIN priorities p ON p.priority_id = mr.priority_id
            WHERE LOWER(COALESCE(mr.status, '')) = 'pending'
            ORDER BY mr.request_date DESC
            """
        ),
        "aging_requests": fetch_all(
            """
            SELECT mr.request_number, mr.title, mr.request_date, mr.status,
                   DATEDIFF(NOW(), mr.request_date) AS age_days
            FROM maintenance_requests mr
            WHERE LOWER(COALESCE(mr.status, '')) NOT IN ('completed', 'closed')
            ORDER BY mr.request_date ASC
            """
        ),
        "technician_load": fetch_all(
            """
            SELECT tech.technician_name, tech.specialization,
                   COUNT(a.assignment_id) AS assigned_requests
            FROM technicians tech
            LEFT JOIN assignments a ON a.technician_id = tech.technician_id
            GROUP BY tech.technician_id, tech.technician_name, tech.specialization
            ORDER BY assigned_requests DESC
            """
        ),
        "building_counts": fetch_all(
            """
            SELECT b.building_name, COUNT(mr.request_id) AS request_count
            FROM buildings b
            LEFT JOIN rooms r ON r.building_id = b.building_id
            LEFT JOIN maintenance_requests mr ON mr.room_id = r.room_id
            GROUP BY b.building_id, b.building_name
            ORDER BY request_count DESC
            """
        ),
        "material_usage": fetch_all(
            """
            SELECT m.material_name, SUM(mu.quantity) AS quantity_used,
                   m.unit_cost,
                   SUM(mu.quantity * m.unit_cost) AS total_cost
            FROM material_usage mu
            LEFT JOIN materials m ON m.material_id = mu.material_id
            GROUP BY m.material_id, m.material_name, m.unit_cost
            ORDER BY total_cost DESC
            """
        ),
        "building_costs": fetch_all(
            """
            SELECT b.building_name,
                   COALESCE(SUM(mu.quantity * m.unit_cost), 0) AS total_cost
            FROM buildings b
            LEFT JOIN rooms r ON r.building_id = b.building_id
            LEFT JOIN maintenance_requests mr ON mr.room_id = r.room_id
            LEFT JOIN material_usage mu ON mu.request_id = mr.request_id
            LEFT JOIN materials m ON m.material_id = mu.material_id
            GROUP BY b.building_id, b.building_name
            ORDER BY total_cost DESC
            """
        ),
        "feedback_report": fetch_all(
            """
            SELECT mr.request_number, f.rating, f.comments
            FROM feedback f
            LEFT JOIN maintenance_requests mr ON mr.request_id = f.request_id
            ORDER BY f.feedback_date DESC
            """
        ),
        "response_times": fetch_all(
            """
            SELECT mr.request_number, DATEDIFF(a.assigned_at, mr.request_date) AS response_days
            FROM maintenance_requests mr
            LEFT JOIN assignments a ON a.request_id = mr.request_id
            WHERE a.assigned_at IS NOT NULL AND mr.request_date IS NOT NULL
            ORDER BY response_days
            """
        ),
    }
    return render_template("reports.html", reports=reports_data)


@app.errorhandler(404)
def page_not_found(error):
    return render_template("error.html", error_message="The page you requested could not be found."), 404


@app.errorhandler(Exception)
def handle_exception(error):
    if isinstance(error, Error):
        message = db_error_message(error)
    else:
        app.logger.exception("Unhandled app error")
        message = "Something went wrong. Please try again later."
    return render_template("error.html", error_message=message), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
