import os
from datetime import datetime
from sqlalchemy import (
    create_engine, MetaData, Table, Column,
    Integer, String, DateTime, ForeignKey, text
)
from sqlalchemy.orm import sessionmaker, relationship, declarative_base
from sqlalchemy.exc import SQLAlchemyError
from dotenv import load_dotenv


# Загружаем переменные окружения
load_dotenv()

# Получаем URL базы данных
DATABASE_URL = os.getenv('DATABASE_URL')

if not DATABASE_URL:
    raise ValueError("EXTERNAL_DATABASE_URL не найден в переменных окружения")

# Создаем подключение
engine = create_engine(DATABASE_URL)

# Создаем базовый класс для моделей (исправленная строка)
Base = declarative_base()


# Определяем модели с использованием ORM классов
class Employee(Base):
    __tablename__ = 'employees'

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), nullable=False)
    surname = Column(String(100), nullable=False)
    lastname = Column(String(100))
    team_number = Column(Integer, nullable=False)
    position = Column(String(100), nullable=False, default="Специалист НК II уровня")
    license = Column(String(100), nullable=False)
    license_number = Column(String(100), nullable=False)
    instrument_table = Column(String, default=None)

    # Связь с позициями - исправлено lazy loading
    # licenses = relationship("License", back_populates="employee", lazy='joined')


    def get_info(self):
        info = dict()
        info['id'] = self.id
        info['name'] = self.name
        info['surname'] = self.surname
        info['team_number'] = self.team_number
        info['table'] = self.instrument_table
        return info


class License(Base): # rename: position -> license
    __tablename__ = 'licenses'

    id = Column(Integer, primary_key=True, autoincrement=True, nullable=False)
    # employees_id = Column(Integer, ForeignKey('employees.id'), nullable=False)
    # license_number = Column(String(100), ForeignKey('employees.license_number'), nullable=False)
    license = Column(String(200), nullable=False)
    license_end_date = Column(DateTime, nullable=False)

    # Связь с сотрудником
    # employee = relationship("Employee", back_populates="licenses")

    def get_info(self):
        info = dict()
        info['id'] = self.id
        info['employees_id'] = self.employees_id
        info['license'] = self.license
        info['license_end_date'] = self.license_end_date
        return info


class DatabaseManager:
    def __init__(self):
        self.engine = create_engine(DATABASE_URL)
        self.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=self.engine)

    def create_tables(self):
        """Создает все таблицы в базе данных"""
        try:
            Base.metadata.create_all(self.engine)
            print("Таблицы успешно созданы")
            return True
        except SQLAlchemyError as e:
            print(f"Ошибка при создании таблиц: {e}")
            return False

    def delete_tables(self):
        try:
            Base.metadata.drop_all(self.engine, checkfirst="positions")
            print("Таблицы удалены")
            return True
        except SQLAlchemyError as e:
            print(f"Ошибка удаления: {e}")
            return False

    def get_session(self):
        """Возвращает сессию для работы с базой данных"""
        return self.SessionLocal()

    def add_employee(self, name, surname, team_number, license_name, lastname=None, instrument_table = None, position="Специалист НК II уровня", license_number="NONE"):
        """Добавляет нового сотрудника"""
        session = self.get_session()
        try:
            employee = Employee(
                name=name,
                surname=surname,
                lastname=lastname,
                team_number=team_number,
                position=position,
                license=license_name,
                license_number=license_number,
                instrument_table=instrument_table
            )

            session.add(employee)
            session.commit()
            session.refresh(employee)
            print(f"Сотрудник добавлен с ID: {employee.id}")
            return employee
        except SQLAlchemyError as e:
            session.rollback()
            print(f"Ошибка при добавлении сотрудника: {e}")
            return None
        finally:
            session.close()

    def add_license(self, license_number, license, license_end_date):
        """Добавляет позицию для сотрудника"""
        session = self.get_session()
        try:
            license_obj = License(
                license_number=license_number,
                license=license,
                license_end_date=license_end_date
            )

            session.add(license_obj)
            session.commit()
            session.refresh(license_obj)
            print(f"Позиция добавлена с ID: {license_obj.id}")
            return license_obj
        except SQLAlchemyError as e:
            session.rollback()
            print(f"Ошибка при добавлении позиции: {e}")
            return None
        finally:
            session.close()

    def get_all_employees(self):
        """Получает всех сотрудников"""
        session = self.get_session()
        try:
            employees = session.query(Employee).all()
            return employees
        except SQLAlchemyError as e:
            print(f"Ошибка при получении сотрудников: {e}")
            return []
        finally:
            session.close()

    def print_all_employees(self):
        session = self.get_session()
        try:
            employees = session.query(Employee).all()
            for employer in employees:
                print(employer.get_info())
        except SQLAlchemyError as e:
            print(f"Ошибка при получении сотрудников: {e}")
            return []
        finally:
            session.close()

    def get_employee_with_licenses(self, employee_id):
        """Получает сотрудника с его позициями"""
        session = self.get_session()
        try:
            # Используем joinedload для загрузки связанных данных
            from sqlalchemy.orm import joinedload
            employee = session.query(Employee) \
                .options(joinedload(Employee.licenses)) \
                .filter(Employee.id == employee_id) \
                .first()
            return employee
        except SQLAlchemyError as e:
            print(f"Ошибка при получении сотрудника: {e}")
            return None
        finally:
            session.close()

    def get_employee_with_licenses_alternative(self, employee_id):
        """Альтернативный способ получения сотрудника с позициями"""
        session = self.get_session()
        try:
            employee = session.query(Employee).filter(Employee.id == employee_id).first()
            if employee:
                # Принудительно загружаем связанные данные перед закрытием сессии
                _ = employee.licenses
            return employee
        except SQLAlchemyError as e:
            print(f"Ошибка при получении сотрудника: {e}")
            return None
        finally:
            session.close()

    def get_licenses_by_employee(self, employee_id):
        """Получает все позиции сотрудника"""
        session = self.get_session()
        try:
            licenses = session.query(License) \
                .filter(License.employees_id == employee_id) \
                .all()
            return licenses
        except SQLAlchemyError as e:
            print(f"Ошибка при получении позиций: {e}")
            return []
        finally:
            session.close()

    def get_all_employees_with_licenses(self):
        """Получает всех сотрудников с их позициями"""
        session = self.get_session()
        try:
            from sqlalchemy.orm import joinedload
            employees = session.query(Employee) \
                .options(joinedload(Employee.licenses)) \
                .all()
            return employees
        except SQLAlchemyError as e:
            print(f"Ошибка при получении сотрудников: {e}")
            return []
        finally:
            session.close()

    def update_employee(self, employee_id, name, surname, team_number, lastname=None):
        """Обновляет данные сотрудника"""
        session = self.get_session()
        try:
            employee = session.query(Employee).filter(Employee.id == employee_id).first()
            if employee:
                # for key, value in kwargs.items():
                #     if hasattr(employee, key):
                #         setattr(employee, key, value)
                employee.name = name
                employee.surname = surname
                if lastname:
                    employee.lastname = lastname
                employee.team_number = team_number
                session.commit()
                session.refresh(employee)
                return employee
            return None
        except SQLAlchemyError as e:
            session.rollback()
            print(f"Ошибка при обновлении сотрудника: {e}")
            return None
        finally:
            session.close()

    def delete_employee(self, employee_id):
        """Удаляет сотрудника"""
        session = self.get_session()
        try:
            employee = session.query(Employee).filter(Employee.id == employee_id).first()
            if employee:
                session.delete(employee)
                session.commit()
                return True
            return False
        except SQLAlchemyError as e:
            session.rollback()
            print(f"Ошибка при удалении сотрудника: {e}")
            return False
        finally:
            session.close()

    def raw_sql_query(self, sql_query, params=None):
        """Выполняет сырой SQL запрос"""
        session = self.get_session()
        try:
            result = session.execute(text(sql_query), params or {})
            if sql_query.strip().upper().startswith('SELECT'):
                return result.fetchall()
            else:
                session.commit()
                return result.rowcount
        except SQLAlchemyError as e:
            session.rollback()
            print(f"Ошибка при выполнении SQL запроса: {e}")
            return None
        finally:
            session.close()






# Пример использования
if __name__ == "__main__":

    db = DatabaseManager()
    # print(db.print_all_employees())
    # print("-------------------------------")
    # print(db.get_licenses_by_employee(1)[0].get_info())
    # print(db.get_licenses_by_employee(2)[0].get_info())
    #
    # print(db.get_all_employees_with_licenses())
    #db.update_employee(1, "Наиль", "Таджитдинов", 2, "Азаматович")
    #employ = db.get_all_employees()
    #print("test1")
    # Создаем менеджер базы данных
    # db = DatabaseManager()

    # Создаем таблицы
    # db.raw_sql_query("DROP TABLE employees")
    db.delete_tables()
    # print("Создание таблиц...")
    db.create_tables()
    #
    # # Добавляем сотрудников
    #----------------------------------------------------------
    print("\nДобавление сотрудников...")

    # from test import *
    # doc = Document('work_files/Для_ТО_БН_2025_приборы,_корочки_и_пр.docx')
    #
    #
    #
    # emp1 = db.add_employee("Наиль", "Тажитдинов", 2, "уд. № 0069-0377 по УК, ВИК до 01.03.2026;\nпо ЭК до 01.03.2027.", "Азаматович", doc.tables[6]._tbl.xml)
    # emp2 = db.add_employee("Артур", "Салимов", 2, "уд. № 0069-0963 по ВИК, УК до 30.04.2028.", "Рустемович")
    # emp3 = db.add_employee("Вячеслав", "Парфенов", 1, "уд. № НОАП-0069-0425 по УК до 29.03.2027;\nуд. № 0042-5807 по ПВК, МК, ВИК до 28.02.2028.", "Федорович", doc.tables[0]._tbl.xml)
    # emp4 = db.add_employee("Пономаренко", "Денис", 1, "уд. № 0042-5805 по УК, МК, ПВК, ВИК до 28.02.2028.", "Витальевич")
    # db.print_all_employees()

    #-----------------------------------------------------------
    #
    # Добавляем позиции
    # if emp1 and emp2:
    #     print("\nДобавление позиций...")
    #     db.add_license(
    #         "0069-0377",
    #         "по УК, ВИК",
    #         datetime(2026, 3, 1)
    #     )
    #     db.add_license(
    #         "0069-0377",
    #         "по ЭК",
    #         datetime(2027, 3, 1)
    #     )
    #     db.add_license(
    #         "0069-0963",
    #         "по ВИК, УК",
    #         datetime(2028, 4, 30)
    #     )
    # Получаем всех сотрудников
    # print("\nВсе сотрудники:")
    # employees = db.get_all_employees()
    # for emp in employees:
    #     print(f"{emp.id}: {emp.surname} {emp.name} {emp.lastname or ''}")
    #
    # # Получаем сотрудника с позициями (исправленный метод)
    # if employees:
    #     print(f"\nДетали сотрудника {employees[0].id}:")
    #     employee = db.get_employee_with_positions(employees[0].id)
    #     if employee:
    #         print(f"Сотрудник: {employee.surname} {employee.name}")
    #         if employee.positions:
    #             for pos in employee.positions:
    #                 print(f"  Позиция: {pos.position}, Лицензия до: {pos.license_end_date}")
    #         else:
    #             print("  Нет позиций")

    # Альтернативный способ получения всех сотрудников с позициями
    # print("\nВсе сотрудники с их позициями:")
    # all_employees = db.get_all_employees_with_positions()
    # for emp in all_employees:
    #     print(f"\n{emp.surname} {emp.name}:")
    #     if emp.positions:
    #         for pos in emp.positions:
    #             print(f"  - {pos.position} (лицензия до {pos.license_end_date})")
    #     else:
    #         print("  Нет позиций")

    # # Пример обновления
    # print("\nОбновление данных сотрудника...")
    # if employees:
    #     updated = db.update_employee(
    #         employees[0].id,
    #         lastname="НовоеОтчество"
    #     )
    #     if updated:
    #         print(f"Обновлен сотрудник: {updated.surname} {updated.name} {updated.lastname}")
    #
    # # Пример сырого SQL запроса
    # print("\nСтатистика (сырой SQL запрос):")
    # result = db.raw_sql_query("""
    #     SELECT
    #         e.surname || ' ' || e.name as full_name,
    #         COUNT(p.id) as positions_count
    #     FROM employees e
    #     LEFT JOIN positions p ON e.id = p.employees_id
    #     GROUP BY e.id, e.surname, e.name
    #     ORDER BY e.surname
    # """)
    #
    # if result:
    #     for row in result:
    #         print(f"{row[0]}: {row[1]} позиций")
