from typing import List, Dict, Optional, Callable
from abc import ABC, abstractmethod
from .models import Section, BridgeData, CalculationResult, ProcessingResult, Point, DeepestPoint
from .events import EventBus, Event, EventType
from ..utils import log_message, log_error
import os
import math


class BaseService(ABC):
    def __init__(self):
        self._event_bus = EventBus()
    
    def _publish_progress(self, message: str, progress: int = None) -> None:
        data = {'message': message}
        if progress is not None:
            data['progress'] = progress
        self._event_bus.publish(Event(EventType.PROCESS_PROGRESS, data))
    
    def _publish_error(self, message: str) -> None:
        self._event_bus.publish(Event(EventType.ERROR_OCCURRED, message))
    
    def _publish_warning(self, message: str) -> None:
        self._event_bus.publish(Event(EventType.WARNING_OCCURRED, message))


class StorageCalculationService(BaseService):
    def __init__(self):
        super().__init__()
        self._stop_flag = False
    
    def stop(self) -> None:
        self._stop_flag = True
        self._publish_progress("收到停止信号，正在停止计算...")
    
    def _check_stop(self) -> bool:
        if self._stop_flag:
            self._publish_progress("计算已停止")
            return True
        return False
    
    def _calculate_section_area(self, points: List[Point], vertex_elevation: float) -> float:
        if not points or len(points) < 2:
            return 0.0
        
        sorted_points = sorted(points, key=lambda p: p.x)
        area = 0.0
        
        for i in range(len(sorted_points) - 1):
            p1, p2 = sorted_points[i], sorted_points[i + 1]
            depth1 = max(0, vertex_elevation - p1.elevation)
            depth2 = max(0, vertex_elevation - p2.elevation)
            
            if depth1 > 0 or depth2 > 0:
                width = abs(p2.x - p1.x)
                avg_depth = (depth1 + depth2) / 2
                area += width * avg_depth
        
        return round(area, 4)
    
    def _find_deepest_point(self, points: List[Point]) -> Optional[DeepestPoint]:
        if not points:
            return None
        
        min_elevation = float('inf')
        deepest = None
        
        for p in points:
            if p.elevation < min_elevation:
                min_elevation = p.elevation
                deepest = p
        
        if deepest:
            return DeepestPoint(
                x=deepest.x,
                y=deepest.y,
                elevation=deepest.elevation
            )
        return None
    
    def _calculate_distance(self, point1: DeepestPoint, point2: DeepestPoint) -> float:
        if not point1 or not point2:
            return 0.0
        return math.sqrt((point1.x - point2.x) ** 2 + (point1.y - point2.y) ** 2)
    
    def _calculate_distance_by_latlon(self, lon1: float, lat1: float, lon2: float, lat2: float) -> float:
        R = 6371000
        lat1_rad = math.radians(lat1)
        lon1_rad = math.radians(lon1)
        lat2_rad = math.radians(lat2)
        lon2_rad = math.radians(lon2)
        
        dlat = lat2_rad - lat1_rad
        dlon = lon2_rad - lon1_rad
        
        a = math.sin(dlat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2) ** 2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        
        return R * c
    
    def process_bridge(self, bridge_data: BridgeData, sections: List[Section]) -> CalculationResult:
        if self._check_stop():
            return CalculationResult(
                bridge_name=bridge_data.name,
                sections=[],
                success=False,
                error_message="计算已停止"
            )
        
        self._publish_progress(f"处理桥: {bridge_data.name}")
        
        for sec in sections:
            sec.area = self._calculate_section_area(sec.points, bridge_data.elevation_85)
            
            if sec.area < 1.0:
                avg_elev = sum(p.elevation for p in sec.points) / len(sec.points)
                self._publish_warning(f"断面 {sec.name} 面积异常，已自动调整")
                if bridge_data.elevation_85 < avg_elev - 0.1:
                    new_vertex = avg_elev + 1.0
                    sec.area = self._calculate_section_area(sec.points, new_vertex)
            
            sec.vertex_elevation = bridge_data.elevation_85
            sec.deepest_point = self._find_deepest_point(sec.points)
        
        sections.sort(key=lambda x: x.deepest_point.elevation if x.deepest_point else float('inf'))
        
        total_volume = 0.0
        volume_details = []
        
        for i in range(len(sections) - 1):
            if self._check_stop():
                return CalculationResult(
                    bridge_name=bridge_data.name,
                    sections=sections,
                    success=False,
                    error_message="计算已停止"
                )
            
            sec1, sec2 = sections[i], sections[i + 1]
            
            if sec1.deepest_point and sec2.deepest_point:
                if sec1.deepest_point.longitude and sec2.deepest_point.longitude:
                    distance = self._calculate_distance_by_latlon(
                        sec1.deepest_point.longitude,
                        sec1.deepest_point.latitude,
                        sec2.deepest_point.longitude,
                        sec2.deepest_point.latitude
                    )
                    distance_source = '经纬度'
                else:
                    distance = self._calculate_distance(sec1.deepest_point, sec2.deepest_point)
                    distance_source = '平面坐标'
            else:
                continue
            
            if distance < 0.1:
                continue
            
            avg_area = (sec1.area + sec2.area) / 2
            volume = avg_area * distance
            total_volume += volume
            
            volume_details.append({
                'section1': sec1.name,
                'section2': sec2.name,
                'distance': distance,
                'distance_source': distance_source,
                'avg_area': avg_area,
                'volume': volume
            })
        
        return CalculationResult(
            bridge_name=bridge_data.name,
            sections=sections,
            volume_details=volume_details,
            total_volume=total_volume
        )


class ConversionService(BaseService):
    def __init__(self):
        super().__init__()
    
    def convert_csv_to_report(self, csv_files: List[str], output_dir: str, 
                            template_dir: str, progress_callback: Optional[Callable] = None) -> ProcessingResult:
        from ..modules.csv_to_report import CsvToReportConverter
        
        result = ProcessingResult()
        result.start_time = __import__('datetime').datetime.now()
        result.total_count = len(csv_files)
        
        converter = CsvToReportConverter(output_dir=output_dir, template_dir=template_dir)
        
        def progress(msg):
            self._publish_progress(msg)
            if progress_callback:
                progress_callback(msg)
        
        results = converter.process_all(csv_files, progress)
        
        result.success_count = len(results['success'])
        result.failed_count = len(results['failed'])
        result.success_files = results['success']
        result.failed_files = [f[0] for f in results['failed']]
        result.end_time = __import__('datetime').datetime.now()
        
        return result


class DataCheckService(BaseService):
    def __init__(self):
        super().__init__()
    
    def check_sections(self, xlsx_files: List[str], progress_callback: Optional[Callable] = None) -> ProcessingResult:
        from ..modules.data_check import DataChecker
        
        result = ProcessingResult()
        result.start_time = __import__('datetime').datetime.now()
        
        checker = DataChecker()
        
        def progress(msg):
            self._publish_progress(msg)
            if progress_callback:
                progress_callback(msg)
        
        for xlsx_file in xlsx_files:
            checker.check_report_file(xlsx_file, progress)
        
        result.success_count = len(checker.results['success'])
        result.failed_count = len(checker.results['failed'])
        result.success_files = checker.results['success']
        result.failed_files = [f[0] for f in checker.results['failed']]
        result.warnings = [f"{item[0]}: {'; '.join(item[1])}" for item in checker.results['warnings']]
        result.end_time = __import__('datetime').datetime.now()
        
        return result