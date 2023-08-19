from requests.adapters import HTTPAdapter, Retry
from tqdm import tqdm
import os
import requests
import string
import openpyxl

session = requests.Session()
retry = Retry(connect=3, backoff_factor=0.5)
adapter = HTTPAdapter(max_retries=retry)
session.mount('http://', adapter)
session.mount('https://', adapter)

linkvidtt = input('Masukan File Link Video tiktok Ex (link.txt): ')

# Create an Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Video Info'

# Set column headers
worksheet['A1'] = 'Title'
worksheet['B1'] = 'Video URL'

# Open the file and count the lines
with open(linkvidtt, "r") as linkvid:
    lines = linkvid.readlines()  # Read all lines into a list

    for i, line in enumerate(lines, start=2):  # Start from row 2
        nomer = f"{i - 1}"
        linkvideo = f"{line.strip()}"
        print("=====================================")
        print(nomer, linkvideo)
        headers = {
            'accept': 'application/json, text/javascript, */*; q=0.01',
        }

        params = {
            'url': linkvideo,
            'update': '1',
        }

        response = session.get(
            'https://dl1.tikmate.cc/listFormats', params=params, headers=headers).json()
        video = response['formats']['video'][0]['url']
        tittlesr = response['formats']['title']
        creator = response['formats']['creator']
        tittles = tittlesr.replace(" ", "_")  # Replace spaces with underscores
        valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
        longtitle = ''.join(c for c in tittles if c in valid_chars)
        if len(longtitle) > 100:
            tittle = longtitle[:100]
        print(creator)
        print(tittlesr)

        if not os.path.exists('video'):
            os.makedirs('video')

        # Download the video and show a progress bar
        video_filename = f"video/{tittle}.mp4"
        response = requests.get(video, stream=True)

        # Get the total size of the file from the content-length header
        total_size = int(response.headers.get('content-length', 0))

        # Use the 'with' statement for both file and tqdm
        with open(video_filename, 'wb') as file, tqdm(
                desc=f"Downloading",
                total=total_size,
                unit='B', unit_scale=True, unit_divisor=1024,
                ascii=True
        ) as bar:
            for data in response.iter_content(chunk_size=1024):
                file.write(data)
                bar.update(len(data))

        print("Download completed!")

        # Write title and video URL to the Excel worksheet
        worksheet[f'A{i}'] = tittlesr
        worksheet[f'B{i}'] = f'video/{tittle}.mp4'

# Save the workbook as an Excel file
excel_filename = 'video_info.xlsx'
workbook.save(excel_filename)
print(f'Video information saved in {excel_filename}')
