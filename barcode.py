import io
import tempfile

import requests
from fake_useragent import UserAgent
from PIL import Image, ImageDraw, ImageFont

def get_barcode(barcode_param: str):
	url = f"https://barcode.tec-it.com/barcode.ashx?data={barcode_param}&code=EANUCC128&translate-esc=true"
	res = requests.get(url, headers={'user-agent': ua.random}, timeout=45, stream=True)
	if res.ok:
		with tempfile.SpooledTemporaryFile(max_size=1e9) as buffer:
			for chunk in res.iter_content(chunk_size=1024):
			    buffer.write(chunk)
			
			# write image to a pillow object 
			buffer.seek(0)
			image = Image.open(io.BytesIO(buffer.read()))
	return image

def draw_text_on_barcode(barcode_text, barcode_param):
	barcode_image = get_barcode(barcode_param)
	barcode_image = barcode_image.resize((barcode_image.width+350, barcode_image.height+180))

	width, height = barcode_image.size
	margin = 25
	new_height = height + (2*margin)

	# empty image for code and text - it needs margins for text
	new_image = Image.new('RGB', (width+60, new_height), (255, 255, 255))
	new_image.paste(barcode_image, (0, margin))

	# object to draw text
	draw = ImageDraw.Draw(new_image)

	# create cover for the barcode number
	cover_image = Image.new('RGB', (width, 72), color='white')
	new_image.paste(cover_image, (0, 245))	 
	
	# put space in between the barcode text
	display_text = ' '.join([(barcode_text[i:i+4]) for i in range(0, len(barcode_text), 4)])

	# draw text
	fnt = ImageFont.truetype("./fonts/arialbd.ttf", 52)
	draw.text((100, new_height - 70), display_text, fill=(0, 0, 0), font=fnt)   	
	return new_image

ua = UserAgent(fallback="Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36")