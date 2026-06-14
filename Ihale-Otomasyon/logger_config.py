"""Logging konfigürasyonu ve çalışma yönetimi."""
import logging
import logging.handlers
from pathlib import Path
from datetime import datetime


ROOT = Path(__file__).parent
LOGS_DIR = ROOT / "logs"


def setup_logger(name: str = "ihale_otomasyonu") -> logging.Logger:
    """Logging sistemini başlat. RotatingFileHandler ile günlük döndürme."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Dosya handler (RotatingFileHandler - 5MB başına döner, 5 dosya tutar)
    log_file = LOGS_DIR / f"{name}.log"
    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5,
        encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    
    # Console handler (sadece WARNING ve üstü)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.WARNING)
    
    # Format
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Handler'ları ekle (eğer yoksa)
    if not logger.handlers:
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    
    return logger


def log_program_start(logger: logging.Logger) -> None:
    """Programın başlangıcını logla."""
    logger.info("=" * 60)
    logger.info(f"Program başlatıldı: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 60)


def log_program_end(logger: logging.Logger, exit_code: int = 0) -> None:
    """Programın sonlanışını logla."""
    status = "BAŞARILI" if exit_code == 0 else "HATA İLE SONLANDI"
    logger.info(f"Program sonlandırıldı ({status}): {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 60)


def log_file_operation(logger: logging.Logger, operation: str, file_path: Path) -> None:
    """Dosya operasyonlarını logla."""
    logger.info(f"{operation}: {file_path}")


def log_file_not_found(logger: logging.Logger, file_path: Path, context: str = "") -> None:
    """Dosya bulunamayan hatası logla."""
    ctx = f" ({context})" if context else ""
    logger.error(f"Dosya bulunamadı{ctx}: {file_path}")


def log_summary(logger: logging.Logger, summary: dict) -> None:
    """Operasyon özetini logla."""
    logger.info("=" * 60)
    logger.info("OPERASYON ÖZETİ")
    for key, value in summary.items():
        logger.info(f"  {key}: {value}")
    logger.info("=" * 60)
