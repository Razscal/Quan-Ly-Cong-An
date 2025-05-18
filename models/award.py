from sqlalchemy import Column, Integer, String, ForeignKey
from sqlalchemy.orm import relationship
from database.db_manager import Base

class Award(Base):
    """Award model representing an award given to a person."""
    __tablename__ = 'awards'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    year = Column(Integer, nullable=False)
    person_id = Column(Integer, ForeignKey('people.id'), nullable=False)
    
    # Relationships
    person = relationship("Person", back_populates="awards")
    
    def __repr__(self):
        return f"<Award(id={self.id}, name='{self.name}', year={self.year})>"
