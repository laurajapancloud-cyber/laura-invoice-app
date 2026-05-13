import barcode
from barcode.writer import ImageWriter

rv = barcode.get('code128', 'TEST', writer=ImageWriter())
print(dir(rv.writer))
