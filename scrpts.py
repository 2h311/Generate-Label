from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from xlsxr import XlsxReader
from barcode import draw_text_on_barcode

def cover_weight(weight: str) -> None: 
	weight_cover_image = Image.new(mode, (60, 70), color=color)
	im.paste(weight_cover_image, (420, 110))

	# set the x coordinate to place the weight according to the length of the weight text
	if len(weight) == 1:
		weight_x_coordinate = 466
	elif len(weight) == 2:
		weight_x_coordinate = 446
	else:
		weight_x_coordinate = 430
	draw.text((weight_x_coordinate, 104), weight, fill=9, font=font)

def cover_date(date: str) -> None:
	date = date.split()[0]
	# cover the date
	date_cover_image = Image.new(mode, (160, 40), color=color)
	im.paste(date_cover_image, (1025, 280))
	draw.text((950, 280), date, fill=9, font=font)

def cover_sender(sender: str) -> None:
	sender = sender.upper()
	sender_list = [string.strip() for string in sender.split(',')]
	sedner = f"{' '.join(sender_list[:-3])}\n{' '.join(sender_list[-3:])}" 

	sender_cover_image = Image.new(mode, (1000, 250), color=color)
	im.paste(sender_cover_image, (70, 460))
	draw.text((70, 460), sedner, fill=9, font=ImageFont.truetype(font_path.as_posix(), 62))

def cover_receiver(receiver: str):
	receiver = receiver.upper()
	# cover the receiver
	receiver_cover_image = Image.new(mode, (950, 240), color=color)
	im.paste(receiver_cover_image, (220, 715))
	draw.text((220, 715), receiver, fill=9, font=ImageFont.truetype(font_path.as_posix(), 62))

def cover_orderid(orderid: str):
	orderid_cover_iamge = Image.new(mode, (400, 72), color=color)
	im.paste(orderid_cover_iamge, (220, 1650))
	draw.text((220, 1676), orderid, fill=9, font=ImageFont.truetype(font_path.as_posix(), 47))

def cover_sku(sku: str):
	# cover the sku
	sku_cover_image = Image.new(mode, (420, 72), color=color)
	im.paste(sku_cover_image, (760, 1650))
	draw.text((800, 1655), sku, fill=9, font=ImageFont.truetype(r'fonts/arialbd.ttf', 60))

def generate_label_from_dict(dict_):
	cover_weight(dict_.get('weight'))
	cover_date(dict_.get('order create time'))
	cover_sender(dict_.get('sender'))
	receiver = dict_.get("First name") + " " + dict_.get("Last name")
	cover_receiver(receiver)
	orderid = dict_.get('Order ID')
	cover_orderid(orderid)
	cover_sku(dict_.get('SKU'))

	ai1 = '420'
	zipp = dict_['Postal code']
	barcode_text = dict_['Tracking number']

	barcode_param = f"({ai1}){zipp}({barcode_text[:2]}){barcode_text[2:]}"
	barcode_image = draw_text_on_barcode(barcode_text, barcode_param)

	# paste barcode 
	im.paste(barcode_image, (100, 1210))
	im.save(f"{barcode_text}.png")

def main():
	excel_filename = "order templat.xlsx"
	fields = [
		'sender',
		'Last name',
		'First name',
		'order create time',
		'weight',
		'Order ID',
		'SKU',
		'Postal code',
		'Tracking number',
	]
	reader = XlsxReader(filename=excel_filename, fields=fields)
	data = reader.grab_excel_body()
	for dict_ in data:
		generate_label_from_dict(dict_)

im = Image.open("TMPLATE.png").convert('RGB')
draw = ImageDraw.Draw(im)
mode, color = im.mode, 'white' 
font_path = Path.cwd() / Path('fonts') / Path("arial.ttf")
font = ImageFont.truetype(font_path.as_posix(), 37)
main()