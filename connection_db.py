import psycopg2

def check_password(username, password):
    conn = psycopg2.connect("dbname=norm_time user=postgres password=admin host=localhost port=5433")
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT app_password FROM users WHERE username = %(username)s', {'username': username})
        records = cursor.fetchall()
        if password == records[0][0]:
            return True
        else:
            return False
    except:
        pass

def get_tz(username):
    conn = psycopg2.connect("dbname=norm_time user=postgres password=admin host=localhost port=5433")
    try:
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM users WHERE username = %(username)s', {'username': username})
        records = cursor.fetchall()
        cursor.execute('SELECT * FROM information_tz WHERE user_id = %(user_id)s', {'user_id': records[0][0]})
        records = cursor.fetchall()

        return records
    except:
        pass