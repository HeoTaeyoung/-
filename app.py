from flask import Flask, render_template, request, jsonify
import os
import olefile, re
import zlib
import struct

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
person_pattern = re.compile(r'\d{6}[-]\d{7}\b')
num_pattern = re.compile(r'\b(01[016789]-?\d{4}-?\d{4}|0\d{1,2}-?\d{3}-?\d{4})\b')
addr_pattern = re.compile(r'([가-힣]{2,6}(시|도)\s?[가-힣]{1,4}(군|구|시)\s?[가-힣0-9\-]+(읍|리|로|길)\s?\d{1,4})')
card_pattern = re.compile(r'\b(?:\d{4}-){3}\d{4}\b')

def get_hwp_text(filename):
    f = olefile.OleFileIO(filename)
    dirs = f.listdir()

    if ["FileHeader"] not in dirs or ["\x05HwpSummaryInformation"] not in dirs:
        raise Exception("Not Valid HWP.")

    header = f.openstream("FileHeader")
    header_data = header.read()
    is_compressed = (header_data[36] & 1) == 1

    nums = []
    for d in dirs:
        if d[0] == "BodyText":
            nums.append(int(d[1][len("Section"):]))

    sections = ["BodyText/Section" + str(x) for x in sorted(nums)]
    text = ""

    for section in sections:
        bodytext = f.openstream(section)
        data = bodytext.read()
        if is_compressed:
            try:
                unpacked_data = zlib.decompress(data, -15)
            except Exception as e:
                print(f"[압축 해제 오류] {e}")
                continue
        else:
            unpacked_data = data

        section_text = ""
        i = 0
        size = len(unpacked_data)

        while i < size:
            try:
                header = struct.unpack_from("<I", unpacked_data, i)[0]
                rec_type = header & 0x3ff
                rec_len = (header >> 20) & 0xfff
            except:
                break

            if rec_type == 67:
                rec_data = unpacked_data[i+4:i+4+rec_len]
                try:
                    section_text += rec_data.decode('utf-16')
                except:
                    pass
                section_text += "\n"

            i += 4 + rec_len

        text += section_text
        text += "\n"

    return text

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "파일이 없습니다."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "파일이 선택되지 않았습니다."}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    patterns = {
        'email': email_pattern,
        'person': person_pattern,
        'num': num_pattern,
        'addr': addr_pattern,
        'card': card_pattern
    }

    try:
        txt = get_hwp_text(filepath)
        result = {}
        for key, pattern in patterns.items():
            result[key] = re.findall(pattern, txt)

        summary = []
        total_count = 0

        for category, items in result.items():
            if items:
                count = len(items)
                summary.append(f"{category}: {count}개")
                total_count += count

        if total_count > 0:
            return {"total": total_count, "summary": summary, "matches": result}
        else:
            return jsonify({"message": "민감정보가 없습니다."})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
