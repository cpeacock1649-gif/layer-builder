import sqlite3
import json
from datetime import datetime
import os


class InsuranceDB:
    def __init__(self, db_path="data/insurance.db"):
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute(
            """CREATE TABLE IF NOT EXISTS accounts
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      account_name TEXT UNIQUE NOT NULL,
                      created_date TEXT,
                      modified_date TEXT)"""
        )

        c.execute(
            """CREATE TABLE IF NOT EXISTS programs
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      account_id INTEGER,
                      program_data TEXT,
                      FOREIGN KEY (account_id) REFERENCES accounts(id))"""
        )

        c.execute(
            """CREATE TABLE IF NOT EXISTS carriers
                     (id INTEGER PRIMARY KEY AUTOINCREMENT,
                      carrier_name TEXT UNIQUE NOT NULL)"""
        )

        conn.commit()
        conn.close()

    def add_account(self, account_name):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        now = datetime.now().isoformat()
        try:
            c.execute(
                "INSERT INTO accounts (account_name, created_date, modified_date) VALUES (?, ?, ?)",
                (account_name, now, now),
            )
            account_id = c.lastrowid

            default_program = {"account": account_name, "layers": []}
            c.execute(
                "INSERT INTO programs (account_id, program_data) VALUES (?, ?)",
                (account_id, json.dumps(default_program)),
            )

            conn.commit()
            return account_id
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()

    def get_all_accounts(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute(
            "SELECT id, account_name, created_date, modified_date FROM accounts ORDER BY account_name"
        )
        accounts = c.fetchall()
        conn.close()
        return accounts

    def get_program(self, account_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT program_data FROM programs WHERE account_id=?", (account_id,))
        result = c.fetchone()
        conn.close()
        if result:
            return json.loads(result[0])
        return None

    def save_program(self, account_id, program_data):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        now = datetime.now().isoformat()
        c.execute(
            "UPDATE programs SET program_data=? WHERE account_id=?",
            (json.dumps(program_data), account_id),
        )
        c.execute("UPDATE accounts SET modified_date=? WHERE id=?", (now, account_id))
        conn.commit()
        conn.close()

    def delete_account(self, account_id):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM programs WHERE account_id=?", (account_id,))
        c.execute("DELETE FROM accounts WHERE id=?", (account_id,))
        conn.commit()
        conn.close()

    def clone_account(self, account_id, new_name):
        program = self.get_program(account_id)
        if program:
            new_id = self.add_account(new_name)
            if new_id:
                program["account"] = new_name
                self.save_program(new_id, program)
                return new_id
        return None

    def add_carrier(self, carrier_name):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        try:
            c.execute("INSERT INTO carriers (carrier_name) VALUES (?)", (carrier_name,))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def get_all_carriers(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT carrier_name FROM carriers ORDER BY carrier_name")
        carriers = c.fetchall()
        conn.close()
        return [c[0] for c in carriers]

    def delete_carrier(self, carrier_name):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM carriers WHERE carrier_name=?", (carrier_name,))
        conn.commit()
        conn.close()
