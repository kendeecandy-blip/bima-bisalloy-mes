from sqlalchemy import Column, Integer, String, Float, DateTime, Boolean, ForeignKey, Enum, Text, Date
from sqlalchemy.orm import declarative_base, relationship
from datetime import datetime
import enum
import pytz

def get_jakarta_time():
    tz = pytz.timezone('Asia/Jakarta')
    return datetime.now(tz).replace(tzinfo=None)

Base = declarative_base()

class VendorType(enum.Enum):
    INTERNAL = "Internal"
    EXTERNAL = "External"

class QCStatus(enum.Enum):
    PASS = "Pass"
    FAIL = "Fail"
    HOLD = "Hold"

class ProjectStatus(enum.Enum):
    WAITING_SCHEDULE = "Waiting Schedule"
    WAITING_DRAWING = "Waiting Drawing"
    IN_PRODUCTION = "In Production"
    COMPLETED = "Completed"

class Customer(Base):
    __tablename__ = "customers"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False, unique=True)
    
class Vendor(Base):
    __tablename__ = "vendors"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False, unique=True)
    type = Column(Enum(VendorType), nullable=False)

class ScopeMaster(Base):
    __tablename__ = "scope_masters"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(100), nullable=False, unique=True)

class ProjectScopeRouting(Base):
    __tablename__ = "project_scope_routings"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False)
    scope_id = Column(Integer, ForeignKey("scope_masters.id"), nullable=False)
    
    project = relationship("Project", back_populates="routings")
    scope = relationship("ScopeMaster")

class Project(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, index=True)
    customer_id = Column(Integer, ForeignKey("customers.id"), nullable=False)
    project_name = Column(String(200), nullable=False)
    so_number = Column(String(50), nullable=False, unique=True)
    sales_pic = Column(String(100), nullable=False)
    total_tonnage = Column(Float, nullable=False, default=0.0)
    
    material_list = Column(String(500), nullable=True) 
    is_jasa_only = Column(Boolean, default=False)
    
    risk_management_filled = Column(Boolean, default=False)
    risk_management_date = Column(DateTime, nullable=True)
    
    start_date = Column(Date, nullable=True) 
    target_selesai_global = Column(Date, nullable=True)
    yej_cycle = Column(String(20), nullable=True) 
    
    logistic_status = Column(String(100), nullable=True)
    delivery_number = Column(String(100), nullable=True)
    actual_delivery_date = Column(Date, nullable=True)
    
    drawing_approved = Column(Boolean, default=False)
    drawing_notes = Column(Text, nullable=True)
    status_proyek = Column(Enum(ProjectStatus), default=ProjectStatus.WAITING_SCHEDULE)
    
    customer = relationship("Customer")
    routings = relationship("ProjectScopeRouting", back_populates="project", cascade="all, delete-orphan")
    progress_logs = relationship("ProjectProgressLog", back_populates="project")

    @property
    def time_performance_status(self):
        if self.status_proyek == ProjectStatus.COMPLETED and self.target_selesai_global:
            end_date = self.actual_delivery_date
            if not end_date:
                last_log = max((log.log_date for log in self.progress_logs), default=None)
                if last_log:
                    end_date = last_log.date()
            if end_date:
                if end_date < self.target_selesai_global:
                    return "LEBIH CEPAT (EARLY)"
                elif end_date == self.target_selesai_global:
                    return "PAS (ON-TIME)"
                else:
                    return "LEBIH LAMA (DELAY)"
        return "ON-PROGRESS"

class ProjectProgressLog(Base):
    __tablename__ = "project_progress_logs"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"), nullable=False)
    scope_id = Column(Integer, ForeignKey("scope_masters.id"), nullable=False)
    vendor_id = Column(Integer, ForeignKey("vendors.id"), nullable=False)
    actual_tonnage_completed = Column(Float, nullable=False, default=0.0)
    log_date = Column(DateTime, default=get_jakarta_time)
    qc_notes = Column(Text, nullable=False)
    status = Column(Enum(QCStatus), nullable=False)
    image_path = Column(String(255), nullable=True)
    image_caption = Column(Text, nullable=True)  # FITUR BARU: Deskripsi Foto Arsip
    material_validated = Column(Boolean, default=False)
    despatch_to_bb = Column(String(255), nullable=True)
    po_vendor_number = Column(String(100), nullable=True)

    project = relationship("Project", back_populates="progress_logs")
    scope = relationship("ScopeMaster")
    vendor = relationship("Vendor")
