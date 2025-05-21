"""
modules/logging_setup.py
Handles logging configuration for the application, including both console and file logging.
"""

import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler
import colorama
from colorama import Fore, Style

# Initialize colorama for Windows
colorama.init()

class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds colors to the log output"""
    
    COLORS = {
        'DEBUG': Fore.BLUE,
        'INFO': Fore.GREEN,
        'WARNING': Fore.YELLOW,
        'ERROR': Fore.RED,
        'CRITICAL': Fore.RED + Style.BRIGHT
    }

    def format(self, record):
        # Add color to the level name
        if record.levelname in self.COLORS:
            record.levelname = f"{self.COLORS[record.levelname]}{record.levelname}{Style.RESET_ALL}"
        return super().format(record)

def setup_logging():
    """Configure logging for both console and file output"""
    
    # Create logs directory if it doesn't exist
    logs_dir = "logs"
    if not os.path.exists(logs_dir):
        os.makedirs(logs_dir)

    # Create a new log file for each session
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(logs_dir, f"app_{timestamp}.log")

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Console handler with colors
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = ColoredFormatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(console_formatter)

    # File handler
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setLevel(logging.INFO)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_formatter)

    # Add handlers to root logger
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    # Log the start of a new session
    logging.info("="*80)
    logging.info(f"Starting new logging session at {timestamp}")
    logging.info("="*80)

    return log_file

def log_user_action(action, details=None):
    """Log a user action with optional details"""
    message = f"User Action: {action}"
    if details:
        message += f" - Details: {details}"
    logging.info(message)

def log_file_operation(operation, file_path):
    """Log a file operation"""
    logging.info(f"File Operation: {operation} - File: {file_path}")

def log_operation_result(operation, status, details=None):
    """Log the result of an operation"""
    message = f"Operation Result: {operation} - Status: {status}"
    if details:
        message += f" - Details: {details}"
    logging.info(message)

def log_error(error, context=None):
    """Log an error with optional context"""
    message = f"Error: {str(error)}"
    if context:
        message += f" - Context: {context}"
    logging.error(message, exc_info=True) 