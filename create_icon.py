from PIL import Image

# Create a new image with a transparent background
img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))

# Save as ICO
img.save('icon.ico', format='ICO')
