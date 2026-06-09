import urllib.request, json, ssl, os, re
ssl._create_default_https_context = ssl._create_unverified_context

queries = {
    "88mutian": "88 mu tian songjiang shanghai b&b",
    "jijiadun": "jijiadun ideal village kunshan",
    "moganshan": "moganshan naked heart resort",
    "qingshancun": "qingshan village hangzhou design library",
    "anji": "anji taohuayuan greentown banyan tree",
}

out_dir = "docs/layer1-基础研究母本/competitor_images"
os.makedirs(out_dir, exist_ok=True)
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

for name, q in queries.items():
    try:
        url = f"https://www.google.com/search?tbm=isch&q={urllib.parse.quote(q)}"
        req = urllib.request.Request(url, headers=headers)
        resp = urllib.request.urlopen(req, timeout=15)
        html = resp.read().decode("utf-8", errors="replace")
        pattern = r'https?://[^"''\s\[\]()<>]+\.(?:jpg|jpeg|png|webp)'
        urls = list(set(re.findall(pattern, html, re.I)))[:5]
        if urls:
            for j, img_url in enumerate(urls[:2]):
                try:
                    img_req = urllib.request.Request(img_url, headers=headers)
                    img_data = urllib.request.urlopen(img_req, timeout=15).read()
                    if len(img_data) > 3000:
                        fname = os.path.join(out_dir, f"{name}_{j}.jpg")
                        with open(fname, "wb") as f: f.write(img_data)
                        print(f"OK: {name} ({len(img_data)/1024:.0f}KB)")
                        break
                except: pass
        else:
            print(f"No URLs: {name}")
    except Exception as e:
        print(f"Error {name}: {str(e)[:80]}")

total = len([f for f in os.listdir(out_dir) if os.path.getsize(os.path.join(out_dir,f))>100])
print(f"\nDownloaded: {total}/5 competitor images")
