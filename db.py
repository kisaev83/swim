from mysql.connector import connect, Error  # pip install mysql-connector-python
from read_config import read_db_config


def connect_to_db():
    db_config = read_db_config(section='mysql')
    conn = connect(**db_config)
    return conn


def close_connect_db(conn):
    conn.close()


def select_db(conn, query=None, fetchall=False, table=None, *column, **data):
    rows = ''
    data_return = {}
    column_names = []
    if table:
        id_ = list(data)[0]
        query = "SELECT"
        if data:
            for field in range(0, len(column)):
                query += " {f},".format(f=column[field])
            query = query.strip(",")
        else:
            query += ' *'
        query += f" FROM {table} WHERE {id_} = {data[id_]}"
        # print('def query_select', query)
    try:
        cursor = conn.cursor(buffered=True)
        cursor.execute(query)
        if fetchall:
            rows = cursor.fetchall()
        else:
            rows = cursor.fetchone()
            # print('ROWS_fetchone', rows)
            column_names = cursor.column_names
    except Error as error:
        print(error)
    finally:
        cursor.close()
        # print('def select_db: ', rows)
        if fetchall:
            return rows
        else:
            if rows:
                for i in range(len(rows)):
                    data_return[column_names[i]] = rows[i]
            else:
                data_return = None
            return data_return


def update_db(conn, query=None, table=None, **data):
    if table:
        query = f"UPDATE {table} SET"
        data_keys = list(data.keys())
        data_values = list(data.values())
        for field in range(1, len(data_keys)):
            query += " {f} = %s,".format(f=data_keys[field])
        query = query.strip(",")
        query += " WHERE " + data_keys[0] + " = %s" % data_values[0]
        del data[data_keys[0]]
    try:
        cursor = conn.cursor()
        cursor.execute(query, list(data.values()) if table else None)
        conn.commit()
    except Error as error:
        print(error)
    finally:
        # print('def update_db_new: ', query)
        cursor.close()


def insert_db(conn, table, **data):
    placeholders = ', '.join(['%s'] * len(data))
    columns = ', '.join(data.keys())
    query = "INSERT INTO %s ( %s ) VALUES ( %s )" % (table, columns, placeholders)
    try:
        cursor = conn.cursor()
        cursor.execute(query, list(data.values()))
        conn.commit()
    except Error as error:
        print(error)
    finally:
        # print('def insert_db: ', query)
        cursor.close()