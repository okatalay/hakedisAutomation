import sqlite3 as sql
def sql_into(table, entry_values):
    vt = None
    try:
        vt = sql.connect('mk_yapidenetim.db')
        cursor = vt.cursor()
        cursor.execute(f"INSERT INTO {table} VALUES ({','.join(['?' for _ in range(len(entry_values))])})", entry_values)
        results = cursor.fetchall()
        vt.commit()

    except sql.Error as e:
        print("SQL_INTO SORUNU:", e)

    finally:
        if vt:
            vt.close()
def sql_insert_or_update(table, entry_values):
    conn = None
    try:
        conn = sql.connect('mk_yapidenetim.db')
        cursor = conn.cursor()
        placeholders = ','.join(['?' for _ in entry_values])
        sql_query = f"INSERT OR REPLACE INTO {table} (ada, parsel, daire, hakedis, personel) VALUES (?, ?, ?, ?, ?)"
        cursor.execute(sql_query, entry_values)
        conn.commit()
    except sql.Error as e:
        print("SQL_INSERT_OR_UPDATE ERROR:", e)
    finally:
        if conn:
            conn.close()


def sql_into_beton(table, entry_values):
    vt = None
    try:
        vt = sql.connect('mk_yapidenetim.db')
        cursor = vt.cursor()
        cursor.execute(f"INSERT OR REPLACE INTO {table} VALUES ({','.join(['?' for _ in range(len(entry_values))])})", entry_values)
        results = cursor.fetchall()
        vt.commit()

    except sql.Error as e:
        print("SQL_INTO SORUNU:", e)

    finally:
        if vt:
            vt.close()
def sql_delete(table, column1, value1, column2, value2):
    vt = None
    try:
        vt = sql.connect('mk_yapidenetim.db')
        cursor = vt.cursor()
        cursor.execute(f"DELETE FROM {table} WHERE {column1}='{value1}' AND {column2}='{value2}'")
        vt.commit()
    except sql.Error as e:
        print("SQL_DELETE SORUNU:", e)
    finally:
        if vt:
            vt.close()

def sql_delete2(table, column1, value1, column2, value2, column3, value3, column4, value4):
    connection = None
    try:
        connection = sql.connect('mk_yapidenetim.db')
        cursor = connection.cursor()

        # Constructing the WHERE clause with AND between each condition
        query = f"DELETE FROM {table} WHERE {column1}=? AND {column2}=? AND {column3}=? AND {column4}=?"

        # Passing values as a tuple
        values = (value1, value2, value3, value4)

        cursor.execute(query, values)
        connection.commit()
        print("Rows deleted successfully.")
    except sql.Error as e:
        print("SQL_DELETE SORUNU:", e)
    finally:
        if connection:
            connection.close()


def sql_query(target, table, column1=None, value1=None, column2=None, value2=None, column3=None, value3=None, column4=None, value4=None):
    vt = None
    try:
        connect = sql.connect('mk_yapidenetim.db')
        vt = connect
        cursor = vt.cursor()

        if column1 is not None and column2 is not None and column3 is not None and column4 is not None:
            cursor.execute(f"SELECT {target} FROM {table} WHERE {column1}='{value1}'AND {column2}='{value2}' AND {column3}='{value3}' AND {column4}='{value4}'")
            results = cursor.fetchall()
            return results

        elif column1 is not None and column2 is not None and column3 is not None:

            cursor.execute(f"SELECT {target} FROM {table} WHERE {column1}='{value1}'AND {column2}='{value2}' AND {column3}='{value3}'")
            results = cursor.fetchall()
            return results

        elif column1 is not None and column2 is not None:

            cursor.execute(f"SELECT {target} FROM {table} WHERE {column1}='{value1}'AND {column2}='{value2}'")
            results = cursor.fetchall()
            return results

        elif column1 is not None:

            cursor.execute(f"SELECT {target} FROM {table} WHERE {column1}='{value1}'")
            results = cursor.fetchall()
            return results

        else:

            target = ', '.join(f'"{column}"' for column in target)
            cursor.execute(f"SELECT {target} FROM {table}")
            results = cursor.fetchall()
            return results

    except sql.Error as e:
        print("SQL_SORGU SORUNU:", e)

    finally:
        if vt:
            vt.close()
def sql_query_all(table, column=None):
    vt = None
    try:
        vt = sql.connect('mk_yapidenetim.db')
        cursor = vt.cursor()

        if column is None:
            cursor.execute(f"SELECT * FROM '{table}'")
            # cursor.execute(f"SELECT * FROM ydk_liste WHERE adi='{value1}' AND sifre='{value2}'")
            results = cursor.fetchall()
        else:
            cursor.execute(f"SELECT {column} FROM '{table}'")
            results = cursor.fetchall()

        return results

    except sql.Error as e:
        print("SQL_SORGU SORUNU:", e)

    finally:
        if vt:
            vt.close()

def sql_update(table: object, update_values: object, condition_columns: object, condition_values: object) -> object:

    conn = sql.connect('mk_yapidenetim.db')
    cursor = conn.cursor()

    cursor.execute(f"PRAGMA table_info({table})")
    columns = [column[1] for column in cursor.fetchall()]

    set_clause = ", ".join(f"{column} = ?" for column in columns)
    where_clause = " AND ".join(f"{column} = ?" for column in condition_columns)
    sql_query = f"UPDATE {table} SET {set_clause} WHERE {where_clause}"

    cursor.execute(sql_query, update_values + condition_values)
    conn.commit()
