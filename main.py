from fastapi import FastAPI, HTTPException, status, Query, Depends
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field, field_validator
from typing import List
import pandas as pd
import io

# Импорты для работы с базой данных
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import declarative_base, sessionmaker, Session

# --- 1. НАСТРОЙКА БАЗЫ ДАННЫХ (SQLAlchemy) ---
DATABASE_URL = "sqlite:///./pstu_sport.db"  # Файл БД создастся сам прямо в папке проекта

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


# Зависимость (Dependency) для получения сессии БД в эндпоинтах
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# --- 2. ORM МОДЕЛЬ (Как таблица выглядит в Базе Данных) ---
class StudentDB(Base):
    __tablename__ = "students"

    id = Column(Integer, primary_key=True, index=True)
    fullname = Column(String, index=True)
    group_code = Column(String)
    attended_hours = Column(Integer, default=0)
    bonus_hours = Column(Integer, default=0)


# Создаем таблицы в файле базы данных
Base.metadata.create_all(bind=engine)


# --- 3. УРОВЕНЬ ООП (БИЗНЕС-ЛОГИКА) ---
class AttendanceCalculator:
    REQUIRED_HOURS = 72

    def __init__(self, attended_hours: int, bonus_hours: int = 0):
        self.attended_hours = attended_hours
        self.bonus_hours = bonus_hours

    def get_total_hours(self) -> int:
        return self.attended_hours + self.bonus_hours

    def is_eligible_for_credit(self) -> bool:
        return self.get_total_hours() >= self.REQUIRED_HOURS

    def get_remaining_hours(self) -> int:
        return max(0, self.REQUIRED_HOURS - self.get_total_hours())


# --- 4. СХЕМЫ ВАЛИДАЦИИ (PYDANTIC МОДЕЛИ) ---
class StudentCreate(BaseModel):
    fullname: str = Field(..., example="Иванов Иван Иванович")
    group_code: str = Field(..., example="ЭТЭД-26-1")
    attended_hours: int = Field(default=0, ge=0, example=12)
    bonus_hours: int = Field(default=0, ge=0, example=2)

    @field_validator('fullname')
    @classmethod
    def validate_name(cls, value: str) -> str:
        if len(value).strip() < 5:
            raise ValueError("ФИО слишком короткое")
        return value


class StudentResponse(BaseModel):
    id: int
    fullname: str
    group_code: str
    total_hours: int
    hours_left_to_earn: int
    status_for_credit: str

    class Config:
        from_attributes = True


# --- 5. СОЗДАНИЕ FastAPI ПРИЛОЖЕНИЯ ---
app = FastAPI(
    title="СпортУчет ПНИПУ v2.0",
    description="Профессиональное API с интеграцией СУБД SQLite и выгрузкой в Excel"
)


# --- 6. ЭНДПОИНТЫ (МАРШРУТЫ) ---

@app.post("/students/", response_model=StudentResponse, status_code=status.HTTP_201_CREATED)
def register_student(student: StudentCreate, db: Session = Depends(get_db)):
    """Регистрация нового студента и сохранение его в СУБД SQLite."""
    db_student = StudentDB(
        fullname=student.fullname,
        group_code=student.group_code,
        attended_hours=student.attended_hours,
        bonus_hours=student.bonus_hours
    )
    db.add(db_student)
    db.commit()
    db.refresh(db_student)

    calc = AttendanceCalculator(db_student.attended_hours, db_student.bonus_hours)
    return StudentResponse(
        id=db_student.id,
        fullname=db_student.fullname,
        group_code=db_student.group_code,
        total_hours=calc.get_total_hours(),
        hours_left_to_earn=calc.get_remaining_hours(),
        status_for_credit="ДОПУЩЕН К ЗАЧЕТУ" if calc.is_eligible_for_credit() else "ОТРАБОТКА"
    )


@app.get("/students/{student_id}", response_model=StudentResponse)
def get_student_status(student_id: int, db: Session = Depends(get_db)):
    """Поиск студента в базе данных по его ID."""
    db_student = db.query(StudentDB).filter(StudentDB.id == student_id).first()
    if not db_student:
        raise HTTPException(status_code=404, detail="Студент не найден в базе данных")

    calc = AttendanceCalculator(db_student.attended_hours, db_student.bonus_hours)
    return StudentResponse(
        id=db_student.id,
        fullname=db_student.fullname,
        group_code=db_student.group_code,
        total_hours=calc.get_total_hours(),
        hours_left_to_earn=calc.get_remaining_hours(),
        status_for_credit="ДОПУЩЕН К ЗАЧЕТУ" if calc.is_eligible_for_credit() else "ОТРАБОТКА"
    )


@app.put("/students/{student_id}/add-hours/")
def add_hours_for_attendance(student_id: int, hours: int = Query(..., gt=0, le=4), db: Session = Depends(get_db)):
    """Добавление часов за посещение занятия с сохранением изменений в базу данных."""
    db_student = db.query(StudentDB).filter(StudentDB.id == student_id).first()
    if not db_student:
        raise HTTPException(status_code=404, detail="Студент не найден")

    db_student.attended_hours += hours
    db.commit()
    return {"message": f"Успешно добавлено {hours} ч. Студенту {db_student.fullname}. Данные сохранены в SQLite."}


@app.get("/students/export/excel")
def export_students_to_excel(db: Session = Depends(get_db)):
    """Выгрузка всех актуальных записей из базы данных в документ Excel."""
    students = db.query(StudentDB).all()
    if not students:
        raise HTTPException(status_code=400, detail="База данных пуста. Нечего экспортировать.")

    report_data = []
    for s in students:
        calc = AttendanceCalculator(s.attended_hours, s.bonus_hours)
        report_data.append({
            "ID студента": s.id,
            "ФИО студента": s.fullname,
            "Учебная группа": s.group_code,
            "Всего набрано часов": calc.get_total_hours(),
            "Осталось отработать (ч)": calc.get_remaining_hours(),
            "Статус допуска": "ДОПУЩЕН" if calc.is_eligible_for_credit() else "НЕ ДОПУЩЕН (ОТРАБОТКА)"
        })

    df = pd.DataFrame(report_data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Ведомость Кафедры')

    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=sport_report_pstu.xlsx"}
    )
