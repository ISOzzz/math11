from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

source = Path(r"D:\第一次模拟\qa_final_pages_v3")
pages = sorted(source.glob("page-*.png"))
thumb_w = 280
margin = 18
label_h = 30
cols = 5

opened = []
for page in pages:
    image = Image.open(page).convert("RGB")
    thumb_h = round(image.height * thumb_w / image.width)
    image.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
    opened.append((page, image.copy()))

rows = (len(opened) + cols - 1) // cols
cell_h = max(image.height for _, image in opened) + label_h
sheet = Image.new(
    "RGB",
    (cols * thumb_w + (cols + 1) * margin, rows * cell_h + (rows + 1) * margin),
    "white",
)
draw = ImageDraw.Draw(sheet)
font = ImageFont.load_default()

for index, (path, image) in enumerate(opened):
    row, col = divmod(index, cols)
    x = margin + col * (thumb_w + margin)
    y = margin + row * (cell_h + margin)
    sheet.paste(image, (x, y))
    draw.text((x + 4, y + image.height + 6), path.stem, fill="black", font=font)

sheet.save(source / "contact_sheet.png", quality=92)
