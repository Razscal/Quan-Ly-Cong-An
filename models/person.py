from sqlalchemy import Column, Integer, String, ForeignKey
from sqlalchemy.orm import relationship
from database.db_manager import Base

class Person(Base):
    """Person model representing an individual in a task."""
    __tablename__ = 'people'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    task_id = Column(Integer, ForeignKey('tasks.id'), nullable=False)
    
    # Relationships
    task = relationship("Task", back_populates="people")
    awards = relationship("Award", back_populates="person", cascade="all, delete-orphan")
    
    def __repr__(self):
        return f"<Person(id={self.id}, name='{self.name}')>"
