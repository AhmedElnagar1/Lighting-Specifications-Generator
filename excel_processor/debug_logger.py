"""
Debug logger module for writing debug messages to a file.
"""

from typing import Optional


class DebugLogger:
    """
    Handles debug logging to a file.
    
    Attributes:
        _debug_file: Optional file handle for the debug file
    """
    
    def __init__(self) -> None:
        """
        Initialize the debug logger.
        """
        self._debug_file: Optional[object] = None
    
    def write(self, message: str) -> None:
        """
        Write a debug message to the debug file.
        
        Args:
            message (str): The debug message to write
        """
        if self._debug_file is None:
            # Open debug file in write mode (overwrites existing)
            self._debug_file = open("debug.txt", "w", encoding="utf-8")
        
        self._debug_file.write(message + "\n")
        self._debug_file.flush()  # Ensure it's written immediately
    
    def close(self) -> None:
        """
        Close the debug file if it's open.
        """
        if self._debug_file is not None:
            self._debug_file.close()
            self._debug_file = None


