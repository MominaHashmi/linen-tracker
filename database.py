from sqlalchemy import create_engine, Column, String, Integer, DateTime, Enum

from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import datetime
import enum

DATABASE_URL = "sqlite:///linen.db"
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

class TowelStatus(enum.Enum):
    REGISTERED = "registered"
    IN_LAUNDRY = "in_laundry"
    IN_USE = "in_use"
    ASSIGNED = "assigned"
    MISSING = "missing"
    RETIRED = "retired"

class Towel(Base):
    __tablename__ = "towels"
    tag_id = Column(String, primary_key=True)
    towel_type = Column(String)
    status = Column(String, default="registered")
    last_location = Column(String, nullable=True)
    wash_count = Column(Integer, default=0)
    dispatched_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

class Event(Base):
    __tablename__ = "events"
    event_id = Column(Integer, primary_key=True, autoincrement=True)
    tag_id = Column(String)
    event_type = Column(String)
    reader_id = Column(String, nullable=True)
    location = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.datetime.utcnow)

def init_db():
    Base.metadata.create_all(bind=engine)

if __name__ == "__main__":
    init_db()
    print("Database created successfully!")
