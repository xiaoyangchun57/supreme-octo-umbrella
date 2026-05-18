from dataclasses import dataclass, field
from typing import List, Optional, Dict
from datetime import datetime


@dataclass
class Point:
    x: float
    y: float
    elevation: float = 0.0
    marker: str = ""
    code: str = ""
    
    def __repr__(self) -> str:
        return f"Point({self.x:.2f}, {self.y:.2f}, {self.elevation:.3f})"


@dataclass
class DeepestPoint:
    x: float
    y: float
    elevation: float
    longitude: Optional[float] = None
    latitude: Optional[float] = None


@dataclass
class Section:
    name: str
    file_path: str
    points: List[Point] = field(default_factory=list)
    deepest_point: Optional[DeepestPoint] = None
    area: float = 0.0
    longitude: Optional[float] = None
    latitude: Optional[float] = None
    vertex_elevation: float = 0.0
    
    @property
    def has_valid_area(self) -> bool:
        return self.area > 1.0


@dataclass
class BridgeData:
    name: str
    original_elevation: float
    elevation_85: float
    section_id: Optional[str] = None
    subtract_value: Optional[float] = None


@dataclass
class CalculationResult:
    bridge_name: str
    sections: List[Section]
    volume_details: List[Dict] = field(default_factory=list)
    total_volume: float = 0.0
    success: bool = True
    error_message: str = ""


@dataclass
class ProcessingResult:
    success_count: int = 0
    failed_count: int = 0
    total_count: int = 0
    success_files: List[str] = field(default_factory=list)
    failed_files: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    start_time: Optional[datetime] = None
    end_time: Optional[datetime] = None
    
    @property
    def duration(self) -> float:
        if self.start_time and self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0.0
    
    @property
    def success_rate(self) -> float:
        if self.total_count > 0:
            return (self.success_count / self.total_count) * 100
        return 0.0


@dataclass
class ProjectConfig:
    root_dir: str = ""
    data_dir: str = ""
    template_dir: str = ""
    output_dir: str = ""
    last_modified: datetime = field(default_factory=datetime.now)
    
    def validate(self) -> List[str]:
        errors = []
        if not self.root_dir or not self.root_dir.exists():
            errors.append("根目录不存在")
        if not self.data_dir or not self.data_dir.exists():
            errors.append("数据目录不存在")
        if not self.template_dir or not self.template_dir.exists():
            errors.append("模板目录不存在")
        return errors


@dataclass
class FileInfo:
    name: str
    path: str
    size: int = 0
    modified_time: datetime = field(default_factory=datetime.now)
    file_type: str = "unknown"
    
    @property
    def size_str(self) -> str:
        if self.size < 1024:
            return f"{self.size} B"
        elif self.size < 1024 * 1024:
            return f"{self.size / 1024:.2f} KB"
        else:
            return f"{self.size / (1024 * 1024):.2f} MB"