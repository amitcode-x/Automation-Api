import pymysql
import threading
import time

# ==========================================
# GLOBAL REUSABLE CONNECTION
# ==========================================
_connection = None
_connection_lock = threading.Lock()

# optional reconnect retry config
MAX_RETRIES = 2
RETRY_DELAY = 0.3


# ==========================================
# GET CONNECTION
# ==========================================
def get_connection():

    global _connection

    try:

        with _connection_lock:

            # create new connection if not exists
            if _connection is None:

                print("Creating new DB connection")

                _connection = pymysql.connect(
                    host="micro.satyukt.com",
                    user="dummy_api_writer_prod",
                    password="@mQ4rZ9!tNp3",
                    db="dummy_api",
                    autocommit=True,
                    charset="utf8mb4",
                    connect_timeout=5,
                    read_timeout=10,
                    write_timeout=10,
                    cursorclass=pymysql.cursors.Cursor
                )

                return _connection

            # validate existing connection
            try:

                _connection.ping(reconnect=False)

            except Exception as ping_error:

                print("Stale DB connection detected:", str(ping_error))

                try:
                    _connection.close()
                except:
                    pass

                _connection = None

                print("Recreating DB connection")

                _connection = pymysql.connect(
                    host="micro.satyukt.com",
                    user="dummy_api_writer_prod",
                    password="@mQ4rZ9!tNp3",
                    db="dummy_api",
                    autocommit=True,
                    charset="utf8mb4",
                    connect_timeout=5,
                    read_timeout=10,
                    write_timeout=10,
                    cursorclass=pymysql.cursors.Cursor
                )

            return _connection

    except Exception as e:

        print("DB CONNECTION ERROR:", str(e))

        try:
            if _connection:
                _connection.close()
        except:
            pass

        _connection = None

        raise e


# ==========================================
# INTERNAL EXECUTOR
# ==========================================
def execute_query(query, values=None, fetch_one=False, fetch_all=False):

    cursor = None

    for attempt in range(MAX_RETRIES):

        try:

            db = get_connection()

            cursor = db.cursor()

            cursor.execute(query, values or ())

            if fetch_one:
                return cursor.fetchone()

            if fetch_all:
                return cursor.fetchall()

            return cursor

        except pymysql.err.OperationalError as e:

            print(f"Operational DB Error (attempt {attempt+1}):", str(e))

            global _connection

            try:
                if _connection:
                    _connection.close()
            except:
                pass

            _connection = None

            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
                continue

            raise e

        finally:

            if cursor:
                cursor.close()


# ==========================================
# EXISTS
# ==========================================
def exists(table_name, **kwargs):

    try:

        sql_query = f"SELECT EXISTS(SELECT 1 FROM {table_name}"

        values = []

        if kwargs:

            conditions = []

            for key, value in kwargs.items():

                if isinstance(value, (list, tuple)):

                    placeholders = ', '.join(['%s'] * len(value))

                    conditions.append(f"`{key}` IN ({placeholders})")

                    values.extend(value)

                elif key.lower() == "date":

                    conditions.append(f"DATE({key}) = %s")

                    values.append(value)

                else:

                    conditions.append(f"`{key}` = %s")

                    values.append(value)

            sql_query += " WHERE " + " AND ".join(conditions)

        sql_query += " LIMIT 1)"

        result = execute_query(
            sql_query,
            tuple(values),
            fetch_one=True
        )

        return result if result else (0,)

    except pymysql.MySQLError as e:

        return [0, f"Database error: {e}"]

    except Exception as e:

        return [0, f"Error: {str(e)}"]


# ==========================================
# FETCH
# ==========================================
def fetch(table_name, columns="*", limit=None, offset=None, order_by=None, **kwargs):

    try:

        # SELECT clause
        if isinstance(columns, str) and columns == "*":

            select_clause = "*"

        elif isinstance(columns, list):

            select_clause = ', '.join([f"`{col}`" for col in columns])

        else:

            select_clause = columns

        sql_query = f"SELECT {select_clause} FROM {table_name}"

        values = []

        # WHERE clause
        if kwargs:

            conditions = []

            for key, value in kwargs.items():

                if isinstance(value, (list, tuple)):

                    placeholders = ', '.join(['%s'] * len(value))

                    conditions.append(f"`{key}` IN ({placeholders})")

                    values.extend(value)

                elif key.lower() == "date":

                    conditions.append(f"DATE({key}) = %s")

                    values.append(value)

                else:

                    conditions.append(f"`{key}` = %s")

                    values.append(value)

            sql_query += " WHERE " + " AND ".join(conditions)

        # ORDER BY
        if order_by is not None:

            sql_query += f" ORDER BY {order_by}"

        elif table_name == "polygonStore":

            sql_query += " ORDER BY Time DESC"

        # LIMIT OFFSET
        if limit is not None:

            sql_query += " LIMIT %s"

            values.append(int(limit))

            if offset is not None:

                sql_query += " OFFSET %s"

                values.append(int(offset))

        result = execute_query(
            sql_query,
            tuple(values),
            fetch_all=True
        )

        return result if result else []

    except pymysql.MySQLError as e:

        return [0, f"Database error: {e}"]

    except Exception as e:

        return [0, f"Error: {str(e)}"]


# ==========================================
# INSERT
# ==========================================
def insert(table_name, **kwargs):

    cursor = None

    try:

        db = get_connection()

        cursor = db.cursor()

        columns = ', '.join(kwargs.keys())

        placeholders = ', '.join(['%s' for _ in kwargs])

        values = tuple(kwargs.values())

        sql_query = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

        cursor.execute(sql_query, values)

        inserted_id = cursor.lastrowid

        if table_name == "polygonStore":

            return [[inserted_id, "successfully inserted"]]

        return [[1, "successfully inserted"]]

    except pymysql.MySQLError as e:

        if e.args[0] == 1062:

            return [[0, "Value already exists"]]

        return [[0, f"MySQL error: {e}"]]

    except Exception as e:

        return [[0, f"Error: {str(e)}"]]

    finally:

        if cursor:
            cursor.close()


# ==========================================
# UPDATE
# ==========================================
def update(table_name, conditions, **kwargs):

    cursor = None

    try:

        db = get_connection()

        cursor = db.cursor()

        set_clause = ', '.join([f"{key} = %s" for key in kwargs])

        values = list(kwargs.values())

        condition_clause = ' AND '.join([f"{key} = %s" for key in conditions])

        condition_values = list(conditions.values())

        all_values = values + condition_values

        sql_query = f"UPDATE {table_name} SET {set_clause} WHERE {condition_clause}"

        cursor.execute(sql_query, tuple(all_values))

        if cursor.rowcount > 0:

            return [[1, "Updated successfully"]]

        return [[0, "No rows matched the condition"]]

    except pymysql.MySQLError as e:

        return [[0, f"MySQL error: {e}"]]

    except Exception as e:

        return [[0, f"Error: {str(e)}"]]

    finally:

        if cursor:
            cursor.close()