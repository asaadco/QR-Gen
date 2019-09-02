import dlib
from scipy import misc
from PIL import Image, ImageTk
from skimage import io
import matplotlib.pyplot as plt


def detect_faces(image):

    # Create a face detector
    face_detector = dlib.get_frontal_face_detector()

    # Run detector and get bounding boxes of the faces on image.
    detected_faces = face_detector(image, 1)
   
    return detected_faces


# Load image
def face_crop(img_path):
    image = io.imread(img_path)

# Detect faces
# detected_faces = detect_faces(image)
    filename = img_path.split(".")[0]+" Cropped"+".jpg"
    dets = detect_faces(image) # a list of faces (Rectangles)
    print("Number of faces detected: {}".format(len(dets)))
    for k, d in enumerate(dets):
        print("Detection {}: Left: {} Top: {} Right: {} Bottom: {}".format(k, d.left(), d.top(), d.right(), d.bottom()))
        vOffset = round(d.top()*0.25)
        hOffset = round(d.left()*0.25)
        # Could this lead to errors? MAYBE
        crop = image[d.top()-vOffset:d.bottom()+vOffset, d.left()-hOffset:d.right()+hOffset]

        io.imsave(filename, crop)
    if(dets):
        return filename
    else: 
        return img_path
        #
# Crop faces and plot
# for i, d in enumerate(detect_faces):
#     print("Detection {}".format(d))
#     io.imwrite("output.jpg", corp)

