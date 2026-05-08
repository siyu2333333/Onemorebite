import hashlib
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

import config


def download_batch(urls, target_dir=None):
    target_dir = Path(target_dir or config.IMAGE_DIR)
    target_dir.mkdir(parents=True, exist_ok=True)

    unique_urls = list(set(u for u in urls if u))
    if not unique_urls:
        return {}

    print(f"下载 {len(unique_urls)} 张图片 (并发:{config.DOWNLOAD_WORKERS})...")
    url_map = {}
    failed = 0

    with ThreadPoolExecutor(max_workers=config.DOWNLOAD_WORKERS) as pool:
        futures = {pool.submit(_download_one, url, target_dir): url for url in unique_urls}
        for i, future in enumerate(as_completed(futures)):
            url, local_path, ok = future.result()
            if ok:
                url_map[url] = local_path
            else:
                failed += 1
            if (i + 1) % 10 == 0 or i == len(futures) - 1:
                print(f"  进度: {i + 1}/{len(unique_urls)} (失败:{failed})")

    print(f"  完成: {len(url_map)} 成功, {failed} 失败")
    return url_map


def _download_one(url, target_dir):
    url_hash = hashlib.md5(url.encode()).hexdigest()[:12]
    ext = _guess_ext(url)
    local_path = target_dir / f"{url_hash}{ext}"

    if local_path.exists():
        return url, str(local_path), True

    for attempt in range(config.DOWNLOAD_RETRIES):
        try:
            resp = requests.get(url, timeout=30, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            })
            resp.raise_for_status()
            local_path.write_bytes(resp.content)
            return url, str(local_path), True
        except Exception:
            if attempt < config.DOWNLOAD_RETRIES - 1:
                time.sleep(2 ** attempt)
    return url, "", False


def _guess_ext(url):
    url_lower = url.split("?")[0].lower()
    if ".webp" in url_lower:
        return ".webp"
    if ".png" in url_lower:
        return ".png"
    if ".jpg" in url_lower or ".jpeg" in url_lower:
        return ".jpg"
    return ".jpg"
