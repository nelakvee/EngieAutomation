# utils/logger_setup.py
"""
Utility for setting up a centralized logger.
"""
import logging
import sys
from config import LOG_FILE_PATH

def setup_logger():
    """
    Sets up a logger that outputs to both console and a file.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)-8s] %(message)s",
        handlers=[]
    )