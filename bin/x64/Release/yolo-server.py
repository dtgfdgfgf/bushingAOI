from flask import Flask, request, jsonify
from ultralytics import YOLO
from PIL import Image
import io
import base64
import numpy as np
import cv2
import sys
import argparse
app = Flask(__name__)

# 設定命令行參數解析
parser = argparse.ArgumentParser()
parser.add_argument("--model_path", type=str, help="Model path")
parser.add_argument("--port", type=int, default=5000, help="Port to run server on")
args = parser.parse_args()

# 獲取模型路徑和端口
model_path = args.model_path
port = args.port

# 全域變數，用於儲存當前載入的 YOLOv8 模型 (與之前相同)
model = None

# 函式：載入 YOLOv8 模型 (與之前相同)
def load_yolo_model(model_path):
    global model
    try:
        model = YOLO(model_path)
        return True, f"模型 '{model_path}' 載入成功"
    except Exception as e:
        return False, f"模型載入失敗： {e}"

# 伺服器啟動時載入預設模型 (與之前相同)
#model_path = './model/yolov8s.pt'

# if len(sys.argv) > 1:
#     model_path = sys.argv[1]
#     print("模型路徑為:", model_path)
# else:
#     print("沒有參數傳入")
success, message = load_yolo_model(model_path)
if success:
    print(message)
else:
    print(f"警告：預設模型載入失敗，將無法進行物件偵測。錯誤訊息：{message}")


# 新增處理根路徑 "/" GET 請求的路由和函數
@app.route('/', methods=['GET'])
def index():
    """
    處理根路徑 "/" 的 GET 請求。
    用於簡單的伺服器健康檢查。
    """
    return jsonify({'status': 'ok', 'message': 'YOLOv8 Inference Server is running'}), 200 # 返回 200 OK 和 JSON 訊息


@app.route('/load_model', methods=['POST'])
def load_model_api():
    # load_model_api 函數 (與之前相同，不需要修改)
    try:
        data = request.get_json()
        new_model_path = data.get('model_path')

        if not new_model_path:
            return jsonify({'error': '缺少 "model_path" 參數'}), 400

        success, message = load_yolo_model(new_model_path)

        if success:
            return jsonify({'message': message}), 200
        else:
            return jsonify({'error': message}), 500

    except Exception as e:
        error_message = f"載入模型 API 發生錯誤: {e}"
        print(error_message)
        return jsonify({'error': error_message}), 500


@app.route('/detect', methods=['POST'])
def detect_objects():
    # detect_objects 函數 (與之前相同，不需要修改)
    if model is None:
        return jsonify({'error': 'YOLOv8 模型尚未載入，請先呼叫 /load_model API 載入模型'}), 503

    try:
        image_base64 = request.json['image']
        image_bytes = base64.b64decode(image_base64)
        image = Image.open(io.BytesIO(image_bytes))
        results = model(image)

        detections = []
        for result in results[0].boxes:
            x1, y1, x2, y2 = map(int, result.xyxy[0])
            class_id = int(result.cls)
            score = float(result.conf)
            class_name = model.names[class_id]

            detections.append({
                'box': [x1, y1, x2, y2],
                'class_id': class_id,
                'class_name': class_name,
                'score': score
            })

        return jsonify({'detections': detections})

    except Exception as e:
        error_message = f"物件偵測過程中發生錯誤: {e}"
        print(error_message)
        return jsonify({'error': error_message}), 500
@app.route('/split_detect', methods=['POST'])
def split_detect_objects():
    if model is None:
        return jsonify({'error': 'YOLOv8 model not loaded. Please call /load_model API first.'}), 503

    try:
        data = request.get_json()
        image_base64 = data['image']
        img_size = tuple(data.get('img_size', (1600, 1600)))  # (width, height), default (640, 640)
        sub_size = tuple(data.get('sub_size', (700, 700)))  # (width, height), default (320, 320)
        step = data.get('step', 500)  # Default step size
        conf_threshold = float(data.get('conf_threshold', 0.5))
        nms_threshold = float(data.get('nms_threshold', 0.1))


        image_bytes = base64.b64decode(image_base64)
        image = Image.open(io.BytesIO(image_bytes))
        src = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)  # Convert PIL Image to OpenCV Mat


        
        bboxes, classes, confidences = split_yolo(model, src, img_size, sub_size, step, conf_threshold, nms_threshold)

        detections = []
        for i in range(len(bboxes)):
            x1, y1, x2, y2 = map(int, bboxes[i])
            class_id = int(classes[i])
            score = float(confidences[i])
            class_name = model.names[class_id] # assuming model.names is available

            detections.append({
                'box': [x1, y1, x2, y2],
                'class_id': class_id,
                'class_name': class_name,
                'score': score
            })


        return jsonify({'detections': detections})

    except Exception as e:
        error_message = f"Error during split object detection: {e}"
        print(error_message)
        return jsonify({'error': error_message}), 500


def split_yolo(model, src, img_size, sub_size, step, conf_threshold, nms_threshold):
    """
    Splits the image into smaller patches, performs YOLOv8 detection on each patch,
    and combines the results.

    Args:
        model: YOLOv8 model.
        src: OpenCV Mat image.
        img_size: Tuple (width, height) for resizing the input image.
        sub_size: Tuple (width, height) for the size of each sub-image.
        step: Step size for sliding the window.
        conf_threshold: Confidence threshold for YOLOv8 detection.
        nms_threshold: NMS threshold for YOLOv8 detection.

    Returns:
        A tuple containing lists of bounding boxes, class IDs, and confidences.
    """
    #resized_src = cv2.resize(src, img_size)
    resized_src = src
    roi_amount = int(np.ceil((resized_src.shape[1] - sub_size[0]) / step) + 1)  # Shape is (height, width, channels)
    temp_mats = []
    index_list = []

    # Split the image into patches
    for i in range(roi_amount):
        for j in range(roi_amount):
            xmin = j * step
            ymin = i * step
            xmax = min(xmin + sub_size[0], resized_src.shape[1])
            ymax = min(ymin + sub_size[1], resized_src.shape[0])

            roi = resized_src[ymin:ymax, xmin:xmax]

            # Pad the ROI if it's smaller than sub_size
            if roi.shape[1] < sub_size[0]:
                pad_width = sub_size[0] - roi.shape[1]
                roi = cv2.copyMakeBorder(roi, 0, 0, 0, pad_width, cv2.BORDER_CONSTANT, value=(0, 0, 0))
            if roi.shape[0] < sub_size[1]:
                pad_height = sub_size[1] - roi.shape[0]
                roi = cv2.copyMakeBorder(roi, 0, pad_height, 0, 0, cv2.BORDER_CONSTANT, value=(0, 0, 0))

            temp_mats.append(roi)
            index_list.append((j, i))


    candidates = []
    confidences = []
    classes = []
    adjusted_boxes = []
    # Perform YOLO detection on each patch
    for i, roi in enumerate(temp_mats):
        results = model(roi,imgsz=640, conf=conf_threshold)  # Perform inference with confidence threshold

        for result in results[0].boxes: # Access boxes attribute directly
            x1, y1, x2, y2 = map(int, result.xyxy[0])
            candidates.append([x1, y1, x2, y2])  # Store box coordinates
            confidences.append(float(result.conf))  # Store confidence score
            classes.append(int(result.cls))  # Store class ID
            jj, ii = index_list[i]
            xmin = jj * step
            ymin = ii * step
            adjusted_boxes.append([x1 + xmin, y1 + ymin, x2 - x1 , y2 -y1])

    # Perform NMS
    if adjusted_boxes:  # Ensure there are boxes to process
        adjusted_boxes = np.array(adjusted_boxes)
        confidences = np.array(confidences)
        classes = np.array(classes)
        nms_indices = cv2.dnn.NMSBoxes(adjusted_boxes.tolist(), confidences.tolist(), conf_threshold, nms_threshold) #Need to convert to list

        final_boxes = adjusted_boxes[nms_indices.flatten()].tolist() if len(nms_indices) > 0 else [] # Convert to list
        final_classes = classes[nms_indices.flatten()].tolist() if len(nms_indices) > 0 else []
        final_confidences = confidences[nms_indices.flatten()].tolist() if len(nms_indices) > 0 else []
    else:
        final_boxes = []
        final_classes = []
        final_confidences = []

    return final_boxes, final_classes, final_confidences

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=port, debug=True)
    print(f"收到的命令行參數: {sys.argv}")
    print(f"解析後的端口參數: {args.port}")
    print(f"最終使用的端口: {port}")