import os
import logging

logger = logging.getLogger(__name__)
if not logger.handlers:
    logging.basicConfig(level=logging.INFO)


def error(msg, *args, **kwargs):
    """Log an error message via module logger."""
    logger.error(msg, *args, **kwargs)


sheet_name = os.getenv('SHEET_NAME', 'Sheet1')
