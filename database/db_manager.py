import os
import sqlite3
from sqlalchemy import create_engine, Column, Integer, String, Text, Date, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship

# Create base directory for database
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, 'data.db')

# Create SQLAlchemy engine and session
engine = create_engine(f'sqlite:///{DB_PATH}', echo=False)
Session = sessionmaker(bind=engine)
Base = declarative_base()

def init_db():
    """Initialize the database by creating all tables."""
    # Create directory for database if it doesn't exist
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    
    # Import models to ensure they are registered with Base
    from models.task import Task
    from models.person import Person
    from models.award import Award
    
    # Create all tables
    Base.metadata.create_all(engine)
    
    return engine

def get_session():
    """Get a new database session."""
    return Session()
