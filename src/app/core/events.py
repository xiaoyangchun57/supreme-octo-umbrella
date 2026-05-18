from typing import Callable, Any
from dataclasses import dataclass
from enum import Enum


class EventType(Enum):
    FILE_LOADED = "file_loaded"
    PROCESS_STARTED = "process_started"
    PROCESS_COMPLETED = "process_completed"
    PROCESS_PROGRESS = "process_progress"
    ERROR_OCCURRED = "error_occurred"
    WARNING_OCCURRED = "warning_occurred"
    STATUS_CHANGED = "status_changed"


@dataclass
class Event:
    type: EventType
    data: Any = None
    sender: str = ""


class EventBus:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._listeners = {}
        return cls._instance
    
    def subscribe(self, event_type: EventType, listener: Callable[[Event], None]) -> None:
        if event_type not in self._listeners:
            self._listeners[event_type] = []
        self._listeners[event_type].append(listener)
    
    def unsubscribe(self, event_type: EventType, listener: Callable[[Event], None]) -> None:
        if event_type in self._listeners:
            self._listeners[event_type].remove(listener)
    
    def publish(self, event: Event) -> None:
        if event.type in self._listeners:
            for listener in self._listeners[event.type]:
                try:
                    listener(event)
                except Exception as e:
                    print(f"Error in event listener: {e}")


class ProgressTracker:
    def __init__(self, total_steps: int = 100):
        self.total_steps = total_steps
        self.current_step = 0
        self.message = ""
        self._event_bus = EventBus()
    
    def set_total(self, total: int) -> None:
        self.total_steps = total
    
    def advance(self, message: str = "") -> None:
        self.current_step += 1
        if message:
            self.message = message
        self._notify_progress()
    
    def set_progress(self, step: int, message: str = "") -> None:
        self.current_step = min(step, self.total_steps)
        if message:
            self.message = message
        self._notify_progress()
    
    def set_message(self, message: str) -> None:
        self.message = message
        self._notify_progress()
    
    def reset(self) -> None:
        self.current_step = 0
        self.message = ""
    
    def _notify_progress(self) -> None:
        progress = {
            'current': self.current_step,
            'total': self.total_steps,
            'percentage': int((self.current_step / self.total_steps) * 100),
            'message': self.message
        }
        self._event_bus.publish(Event(EventType.PROCESS_PROGRESS, progress))


event_bus = EventBus()