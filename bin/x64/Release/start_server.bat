call "C:\Users\Chernger\anaconda3\Scripts\activate.bat" "diffusion"
python yolo-server.py --model_path "%1" --port %2