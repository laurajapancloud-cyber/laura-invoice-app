import barcode
from barcode.writer import ImageWriter
from io import BytesIO

def test():
    try:
        rv = barcode.get('code128', 'TEST1234', writer=ImageWriter())
        # Check supported options by looking at default_options
        print("Default options:", rv.writer.default_options)
        
        # Test writing with options
        bio = BytesIO()
        rv.write(bio, options={"write_text": False})
        print("Write with write_text=False succeeded")
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    test()
