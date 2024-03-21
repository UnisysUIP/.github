import cv2
from ultralytics import YOLO

model = YOLO('yolov8n-cls.pt')
model.train(data='mypath', epochs=30, imgsz=64)
