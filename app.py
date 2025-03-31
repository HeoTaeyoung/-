from flask import Flask, render_template, request
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    upload_path = 'uploads'
    if "file" not in request.files:
        return "파일이 없습니다.", 400
    
    file = request.files["file"]

    if file.filename == "":
        return "파일이 선택되지 않았습니다.", 400
    
    file.save(os.path.join("uploads", file.filename))
    return "업로드 완료"
    


if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)