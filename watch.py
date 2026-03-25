"""
Downloads watcher — auto-converts CR_*_stretto-parser_*.json to Excel
as soon as the browser scraper drops a new file in Downloads.

Usage: python watch.py
       (keep running in terminal, press Ctrl+C to stop)
"""

import time
import json
from pathlib import Path

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from format_saks_claims import create_excel, COLUMNS

DOWNLOADS = Path.home() / "Downloads"
OUTPUT_DIR = Path(__file__).parent / "output"


class DownloadHandler(FileSystemEventHandler):

    def on_created(self, event):
        if event.is_directory:
            return
        self._check(Path(event.src_path))

    def on_moved(self, event):
        # Chrome renames .crdownload → final filename on completion
        if event.is_directory:
            return
        self._check(Path(event.dest_path))

    def _check(self, path):
        if not path.match("CR_*_stretto-parser_*.json"):
            return

        print(f"\n[Watcher] Detected: {path.name}")
        print("[Watcher] Waiting for download to finish writing...")
        time.sleep(3)

        OUTPUT_DIR.mkdir(exist_ok=True)
        output_path = OUTPUT_DIR / f"{path.stem}.xlsx"

        print(f"[Watcher] Converting → output/{output_path.name}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            count = create_excel(data, str(output_path))
            print(f"[Watcher] Done! {count} claims → output/{output_path.name}\n")
        except Exception as e:
            print(f"[Watcher] Error during conversion: {e}\n")


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"[Watcher] Watching {DOWNLOADS}")
    print("[Watcher] Will auto-convert any CR_*_stretto-parser_*.json that lands there.")
    print("[Watcher] Press Ctrl+C to stop.\n")

    handler = DownloadHandler()
    observer = Observer()
    observer.schedule(handler, str(DOWNLOADS), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    print("\n[Watcher] Stopped.")
