import sqlite3
import os
import pandas as pd
from datetime import datetime

class HouseholdDB:
    def __init__(self, db_path="household.db"):
        # src 폴더 안에 db 생성
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.db_path = os.path.join(current_dir, db_path)
        self.init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def init_db(self):
        """데이터베이스 테이블 생성"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    type TEXT NOT NULL,
                    category TEXT NOT NULL,
                    amount INTEGER NOT NULL,
                    note TEXT
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS recurring_templates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    type TEXT NOT NULL,
                    category TEXT NOT NULL,
                    amount INTEGER NOT NULL,
                    note TEXT
                )
            ''')
            conn.commit()

    def add_transaction(self, date, t_type, category, amount, note):
        """수입/지출 내역 추가"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO transactions (date, type, category, amount, note)
                VALUES (?, ?, ?, ?, ?)
            ''', (date, t_type, category, amount, note))
            conn.commit()

    def delete_transaction(self, t_id):
        """내역 삭제"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM transactions WHERE id = ?', (t_id,))
            conn.commit()

    def get_transactions_by_month(self, year, month):
        """특정 연/월의 내역 조회 (YYYY-MM 형식 필터링)"""
        search_pattern = f"{year}-{month:02d}-%"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, date, type, category, amount, note 
                FROM transactions 
                WHERE date LIKE ? 
                ORDER BY date DESC, id DESC
            ''', (search_pattern,))
            return cursor.fetchall()

    def search_transactions(self, keyword):
        """통합 검색 기능"""
        search_pattern = f"%{keyword}%"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, date, type, category, amount, note 
                FROM transactions 
                WHERE note LIKE ? OR category LIKE ?
                ORDER BY date DESC, id DESC
            ''', (search_pattern, search_pattern))
            return cursor.fetchall()
            
    def get_all_transactions(self):
        """전체 내역 조회"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, date, type, category, amount, note 
                FROM transactions 
                ORDER BY date DESC, id DESC
            ''')
            return cursor.fetchall()

    def get_monthly_summary(self, year, month):
        """특정 연/월의 총수입, 총지출, 카테고리별 지출 통계"""
        search_pattern = f"{year}-{month:02d}-%"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # 총수입
            cursor.execute('SELECT SUM(amount) FROM transactions WHERE type="수입" AND date LIKE ?', (search_pattern,))
            total_income = cursor.fetchone()[0] or 0
            
            # 총지출
            cursor.execute('SELECT SUM(amount) FROM transactions WHERE type="지출" AND date LIKE ?', (search_pattern,))
            total_expense = cursor.fetchone()[0] or 0
            
            # 카테고리별 지출 (차트용)
            cursor.execute('''
                SELECT category, SUM(amount) 
                FROM transactions 
                WHERE type="지출" AND date LIKE ?
                GROUP BY category
                ORDER BY SUM(amount) DESC
            ''', (search_pattern,))
            expense_by_category = dict(cursor.fetchall())
            
        return {
            'income': total_income,
            'expense': total_expense,
            'balance': total_income - total_expense,
            'expense_by_category': expense_by_category
        }

    def get_annual_summary(self, year):
        """연간 월별 수입/지출 통계 조회"""
        search_pattern = f"{year}-%"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT substr(date, 6, 2) as month, SUM(amount) 
                FROM transactions 
                WHERE type="수입" AND date LIKE ?
                GROUP BY month
            ''', (search_pattern,))
            income_rows = cursor.fetchall()
            
            cursor.execute('''
                SELECT substr(date, 6, 2) as month, SUM(amount) 
                FROM transactions 
                WHERE type="지출" AND date LIKE ?
                GROUP BY month
            ''', (search_pattern,))
            expense_rows = cursor.fetchall()
            
            return {
                'income': dict(income_rows),
                'expense': dict(expense_rows)
            }

    def get_unique_categories(self, t_type):
        """특정 유형(수입/지출)에 대해 사용자가 입력했던 모든 고유 카테고리 조회"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT category FROM transactions WHERE type=?', (t_type,))
            rows = cursor.fetchall()
            return [row[0] for row in rows]

    def export_to_excel(self, filename="가계부_내역.xlsx"):
        """전체 내역을 엑셀 파일로 내보내기"""
        with self.get_connection() as conn:
            df = pd.read_sql_query('SELECT date as "날짜", type as "분류", category as "카테고리", amount as "금액(원)", note as "메모" FROM transactions ORDER BY date DESC', conn)
            
        if not df.empty:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            filepath = os.path.join(current_dir, filename)
            try:
                df.to_excel(filepath, index=False, engine='openpyxl')
                return filepath
            except PermissionError:
                return "PERMISSION_ERROR"
            except Exception as e:
                return str(e)
        return None

    # --- 추가: 설정(예산) 및 고정 항목 관리 ---
    def get_setting(self, key, default=None):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT value FROM settings WHERE key=?', (key,))
            row = cursor.fetchone()
            return row[0] if row else default

    def set_setting(self, key, value):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)', (key, str(value)))
            conn.commit()

    def add_recurring_template(self, t_type, category, amount, note):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO recurring_templates (type, category, amount, note)
                VALUES (?, ?, ?, ?)
            ''', (t_type, category, amount, note))
            conn.commit()

    def get_recurring_templates(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, type, category, amount, note FROM recurring_templates')
            return cursor.fetchall()
            
    def delete_recurring_template(self, t_id):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM recurring_templates WHERE id=?', (t_id,))
            conn.commit()
