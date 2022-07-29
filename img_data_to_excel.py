from PIL import Image
from PIL.ExifTags import TAGS

# path to the image or video
imagename = "_Z.jpg"

# read the image data using PIL
image = Image.open(imagename)

# extract other basic metadata
info_dict = {
    "Image Name": image.filename,
    "Image Extension": image.format,
    "Image Size": image.size,
    "Image Height": image.height,
    "Image Width": image.width    
}

for label,value in info_dict.items():
    print(f"{label:25}: {value}")

    # extract EXIF data
exifdata = image.getexif()

# iterating over all EXIF data fields
for tag_id in exifdata:
    # get the tag name, instead of human unreadable tag id
    tag = TAGS.get(tag_id, tag_id)
    data = exifdata.get(tag_id)
    # decode bytes 
    if isinstance(data, bytes):
        data = data.decode()
    print(f"{tag:25}: {data}")