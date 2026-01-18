import os
import requests

# List of image URLs collected from Google Images
image_urls = [
    "https://simplep.net/wp-content/uploads/2019/02/Blanc_free_ppt_title.jpg",
    "https://simplep.net/wp-content/uploads/2018/08/zero-ppt-template-title-1280x640.jpg",
    "https://file.tooldi.com/template_upload/57/992/7749582.png",
    "https://pptbizcam.co.kr/wp-content/uploads/2024/12/%ED%8C%8C%EC%9B%8C%ED%8F%AC%EC%9D%B8%ED%8A%B8-%ED%85%9C%ED%94%8C%EB%A6%BF-%EC%9B%90%EB%B3%B8-%ED%8C%8C%EC%9D%BC-%EB%8B%A4%EC%9A%B4%EB%A1%9C%EB%93%9C-%EB%B0%9B%EA%B8%B0-free-ppt-template-1585.jpg",
    "https://cdn.imweb.me/upload/S202105159ce8646407e94/ac5da39ff6361.png",
    "https://identity.snu.ac.kr/webdata/uploads/identity/image/2021/06/5-12-1.png",
    "https://www.pptshop.co.kr/wp-content/uploads/2022/09/Business-Plan-Template-Cover.jpg",
    "https://marketplace.canva.com/EAGfKUG7Zks/2/0/1600w/canva-%EC%B4%88%EB%A1%9D-%EB%B2%A0%EC%9D%B4%EC%A7%80-%EA%B9%94%EB%81%94%ED%95%9C-%ED%8C%80-%ED%94%84%EB%A1%9C%EC%A0%9D%ED%8A%B8-%EB%B3%B4%EA%B3%A0%EC%84%9C-%ED%94%84%EB%A0%88%EC%A0%A0%ED%85%8C%EC%9D%B4%EC%85%98-unCMuOk3Mkg.jpg",
    "https://cdn.imweb.me/upload/S20211115ba735066ce08d/ed0f56735351a.jpg",
    "https://pptbizcam.co.kr/wp-content/uploads/2024/04/%ED%8C%8C%EC%9B%8C%ED%8F%AC%EC%9D%B8%ED%8A%B8-%ED%85%9C%ED%94%8C%EB%A6%BF-%EC%9B%90%EB%B3%B8-%ED%8C%8C%EC%9D%BC-%EB%8B%A4%EC%9A%B4%EB%A1%9C%EB%93%9C-%EB%B0%9B%EA%B8%B0-free-ppt-template-1479.jpg"
]

# Target directory
target_dir = r"C:\project\pptlib\images-gemini"
if not os.path.exists(target_dir):
    os.makedirs(target_dir)

# Headers to mimic a browser request
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

print(f"Starting download of {len(image_urls)} images to {target_dir}...")

for i, url in enumerate(image_urls):
    try:
        # Determine file extension from URL or default to .jpg
        if ".png" in url.lower():
            ext = ".png"
        else:
            ext = ".jpg"
            
        filename = f"slide_{i+1:02d}{ext}"
        filepath = os.path.join(target_dir, filename)
        
        print(f"Downloading {url} as {filename}...")
        
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        with open(filepath, "wb") as f:
            f.write(response.content)
            print(f"Successfully saved {filename}")
            
    except Exception as e:
        print(f"Failed to download {url}: {e}")

print("Download process completed.")
