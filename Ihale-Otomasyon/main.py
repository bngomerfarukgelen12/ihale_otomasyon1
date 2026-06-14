"""Ana giriş noktası."""
import sys

from hesapla_teklif import main
from logger_config import setup_logger

logger = setup_logger("ihale_otomasyonu")

if __name__ == "__main__":
    raise SystemExit(main(["main.py", *sys.argv[1:]]))
