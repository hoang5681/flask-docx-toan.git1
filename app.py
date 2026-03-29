import os
import re
import hashlib
from flask import Flask, render_template_string, jsonify
import docx
from docx.enum.text import WD_COLOR_INDEX

app = Flask(__name__, static_folder='static')

# Lấy thư mục gốc chứa file app.py hiện tại
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ghép với tên file (Sửa lại thành 'Decuong.docx' cho đúng chữ c thường như trong ảnh)
DOCX_FILE = os.path.join(BASE_DIR, 'Decuong.docx')

# Tạo thư mục lưu ảnh (Cũng dùng BASE_DIR cho an toàn)
os.makedirs(os.path.join(BASE_DIR, 'static', 'images'), exist_ok=True)

def is_correct_format(run):
    """
    Hàm kiểm tra xem một đoạn text (run) có được highlight màu vàng hay không.
    Sử dụng try-except để tránh lỗi khi highlight trong Word được set là 'None'.
    """
    try:
        # Kiểm tra nếu thuộc tính highlight_color là màu vàng
        if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
            return True
    except ValueError:
        # Bắt lỗi "ValueError: WD_COLOR_INDEX has no XML mapping for 'none'"
        # Bỏ qua và ngầm hiểu là không có highlight màu vàng
        pass
        
    return False

def parse_docx(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    part = 1 

    for para in doc.paragraphs:
        text_raw = para.text
        text_stripped = text_raw.strip()
        img_html = ""
        
        # Trích xuất ảnh
        for run in para.runs:
            for drawing in run._element.xpath('.//w:drawing'):
                for pic in drawing.xpath('.//pic:pic'):
                    try:
                        embed_id = pic.xpath('.//a:blip/@r:embed')[0]
                        image_part = doc.part.related_parts[embed_id]
                        img_hash = hashlib.md5(image_part.blob).hexdigest()
                        img_filename = f"img_{img_hash}.png"
                        img_path = os.path.join('static', 'images', img_filename)
                        
                        with open(img_path, 'wb') as f:
                            f.write(image_part.blob)
                        img_html += f'<div class="text-center my-3"><img src="/static/images/{img_filename}" class="img-fluid rounded shadow-sm" style="max-height: 400px; border: 1px solid #dee2e6;"></div>'
                    except Exception:
                        pass

        if not text_stripped and not img_html:
            continue

        if "PHẦN 2" in text_stripped.upper() or "ĐÚNG SAI" in text_stripped.upper() or "PHẦN II" in text_stripped.upper():
            part = 2
            continue

        # Bắt đầu một câu hỏi mới
        is_question = bool(re.match(r'^câu\s*\d+', text_stripped.lower()))
        if is_question:
            if current_q:
                questions.append(current_q)
            current_q = {
                "question_text": "",
                "type": "mc" if part == 1 else "tf",
                "options": []
            }

        if current_q:
            # Lấy định dạng màu sắc của từng ký tự để bóc tách đáp án dính chùm
            char_formats = []
            for run in para.runs:
                fmt = is_correct_format(run)
                char_formats.extend([fmt] * len(run.text))
            
            # Đảm bảo độ dài mảng format khớp với text
            if len(char_formats) < len(text_raw):
                char_formats.extend([False] * (len(text_raw) - len(char_formats)))
                
            is_auto_list = bool(para._element.xpath('.//w:numPr'))
            # Dùng Regex để tách các chữ A. B. C. D. dù nó nằm chung một dòng
            matches = list(re.finditer(r'(?:^|\s)([A-Da-d][.\)])(?=\s|$)', text_raw))
            
            if is_auto_list and not matches:
                # Trường hợp Word tự đánh số list (A, B, C bị ẩn đi)
                idx = len(current_q["options"])
                prefix = f"<b>{chr(65+idx)}.</b> " if current_q['type'] == 'mc' else f"<b>{chr(97+idx)})</b> "
                current_q["options"].append({
                    "text": prefix + text_stripped + img_html,
                    "is_correct": any(char_formats)
                })
            elif matches:
                # Có các chữ A. B. C. D. thủ công
                opt_starts = [m.start(1) for m in matches]
                
                # Nếu đằng trước chữ A. có text, thì text đó thuộc về câu hỏi
                if opt_starts[0] > 0:
                    prefix_text = text_raw[0:opt_starts[0]].strip()
                    if prefix_text:
                        current_q["question_text"] += ("<br>" + prefix_text) if current_q["question_text"] else prefix_text
                
                # Tách từng đoạn đáp án ra
                for i in range(len(opt_starts)):
                    start_idx = opt_starts[i]
                    end_idx = opt_starts[i+1] if i+1 < len(opt_starts) else len(text_raw)
                    
                    opt_text = text_raw[start_idx:end_idx].strip()
                    opt_is_correct = any(char_formats[start_idx:end_idx])
                    opt_text_styled = re.sub(r'^([A-Da-d][.\)])', r'<b>\1</b>', opt_text)
                    final_img = img_html if i == len(opt_starts) - 1 else ""
                    
                    current_q["options"].append({
                        "text": opt_text_styled + final_img,
                        "is_correct": opt_is_correct
                    })
            else:
                # Không phải đáp án -> Nối tiếp nội dung vào câu hỏi
                if not current_q["options"]:
                    current_q["question_text"] += ("<br>" + text_stripped + img_html) if current_q["question_text"] else (text_stripped + img_html)
    
    if current_q:
        questions.append(current_q)
        
    return questions

# ================= GIAO DIỆN HTML/JS =================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Luyện Đề Cương</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">
    <style>
        body { background-color: #f4f7f6; font-family: 'Segoe UI', system-ui, sans-serif; color: #2d3436; margin-bottom: 100px; }
        
        /* Header */
        .header-bar { background: #ffffff; padding: 18px 0; box-shadow: 0 4px 15px rgba(0,0,0,0.03); position: sticky; top: 0; z-index: 1000; }
        .progress-container { height: 4px; background: #dfe6e9; width: 100%; position: absolute; bottom: 0; left: 0; }
        .progress-bar-custom { height: 100%; background: #0984e3; width: 0%; transition: width 0.5s ease; }

        /* Khung câu hỏi */
        .question-card { background: white; border-radius: 16px; padding: 30px; margin-bottom: 30px; box-shadow: 0 4px 20px rgba(0,0,0,0.04); border: 2px solid transparent; transition: all 0.3s; }
        .question-card.graded-correct { border-color: #00b894; background-color: #fafffc; }
        .question-card.graded-wrong { border-color: #d63031; background-color: #fffafb; }
        .question-title { font-size: 1.15rem; font-weight: 600; margin-bottom: 25px; line-height: 1.6; }
        
        /* Đáp án Trắc nghiệm */
        .mc-option { display: block; margin-bottom: 12px; }
        .mc-label { display: flex; align-items: center; padding: 14px 20px; background: #ffffff; border: 2px solid #dfe6e9; border-radius: 12px; cursor: pointer; transition: 0.2s ease; font-weight: 500; margin: 0; }
        .mc-label:hover { background: #f8f9fa; border-color: #b2bec3; transform: translateY(-2px); }
        .mc-input { display: none; }
        .mc-input:checked + .mc-label { background: #e3f2fd; border-color: #0984e3; color: #0984e3; box-shadow: 0 4px 10px rgba(9, 132, 227, 0.15); }
        
        /* Sau khi chấm */
        .mc-label.correct-ans { background-color: #00b894 !important; border-color: #00b894 !important; color: #fff !important; }
        .mc-label.wrong-ans { background-color: #d63031 !important; border-color: #d63031 !important; color: #fff !important; }

        /* Đáp án Đúng/Sai */
        .tf-option { display: flex; justify-content: space-between; align-items: center; padding: 15px 20px; background: #ffffff; border: 2px solid #dfe6e9; border-radius: 12px; margin-bottom: 10px; transition: 0.2s;}
        .tf-option:hover { border-color: #b2bec3; }
        .tf-text { flex-grow: 1; padding-right: 20px; font-weight: 500; }
        .tf-btns label { min-width: 80px; border-radius: 8px !important; }

        /* Text phản hồi */
        .feedback-text { font-weight: bold; font-size: 1rem; margin-top: 20px; padding: 12px 18px; border-radius: 10px; display: none; }

        /* Thanh Nộp Bài ghim ở dưới */
        .bottom-bar { position: fixed; bottom: 0; left: 0; width: 100%; background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(5px); padding: 15px; box-shadow: 0 -4px 15px rgba(0,0,0,0.05); z-index: 1000; display: flex; justify-content: center; border-top: 1px solid #eee;}
    </style>
</head>
<body>

<div class="header-bar">
    <div class="container d-flex justify-content-between align-items-center flex-wrap gap-2">
        <h4 class="m-0 fw-bold" style="color: #0984e3;"><i class="bi bi-journal-bookmark-fill"></i> Luyện Đề Cương</h4>
        <div class="d-flex align-items-center gap-2">
            <span class="badge bg-light text-primary border fs-6 px-3 py-2 shadow-sm" id="progress-text">Đang tải...</span>
            <button class="btn btn-outline-secondary btn-sm fw-bold px-3 shadow-sm" onclick="prevBatch()" id="btn-prev" disabled><i class="bi bi-chevron-left"></i> Lùi lại</button>
            <button class="btn btn-warning btn-sm fw-bold px-3 shadow-sm" onclick="retryBatch()"><i class="bi bi-shuffle"></i> Đảo câu</button>
            <button class="btn btn-primary btn-sm fw-bold px-3 shadow-sm" onclick="nextBatch()" id="btn-next" disabled>Đi tiếp <i class="bi bi-chevron-right"></i></button>
        </div>
    </div>
    <div class="progress-container"><div class="progress-bar-custom" id="top-progress"></div></div>
</div>

<div class="container py-4">
    <div id="quiz-container"></div>
    <div id="result-container" class="alert alert-success fs-5 text-center shadow border-0 rounded-4 py-4 mt-4" style="display: none; background: #00b894; color: white;"></div>
</div>

<div class="bottom-bar" id="submit-area">
    <button class="btn btn-primary btn-lg px-5 fw-bold rounded-pill shadow" id="btn-submit" onclick="submitQuiz()" style="background: #0984e3; border:none;">
        <i class="bi bi-send-check-fill"></i> Nộp Bài Ngay
    </button>
</div>

<script>
    let allQuestions = [];
    let currentBatchData = [];
    let batchIndex = 0;
    const batchSize = 50;

    fetch('/api/questions')
        .then(response => response.json())
        .then(data => {
            allQuestions = data;
            loadBatch(true);
        });

    function shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [array[i], array[j]] = [array[j], array[i]];
        }
    }

    function updateProgress() {
        let pct = Math.min(((batchIndex + 1) * batchSize) / allQuestions.length * 100, 100);
        document.getElementById("top-progress").style.width = pct + "%";
    }

    function loadBatch(doShuffle = true) {
        let slice = allQuestions.slice(batchIndex * batchSize, (batchIndex + 1) * batchSize);
        if (doShuffle) {
            currentBatchData = JSON.parse(JSON.stringify(slice)); 
            shuffleArray(currentBatchData);
        }

        let startNum = batchIndex * batchSize + 1;
        let endNum = Math.min((batchIndex + 1) * batchSize, allQuestions.length);
        document.getElementById("progress-text").innerHTML = `<i class="bi bi-card-list"></i> Câu ${startNum} - ${endNum} / ${allQuestions.length}`;
        updateProgress();

        document.getElementById("btn-prev").disabled = (batchIndex === 0);
        document.getElementById("btn-next").disabled = true;
        document.getElementById("submit-area").style.display = "flex";
        document.getElementById("result-container").style.display = "none";

        renderQuiz(currentBatchData);
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function renderQuiz(questions) {
        const container = document.getElementById("quiz-container");
        container.innerHTML = "";

        questions.forEach((q, qIndex) => {
            let html = `<div class="question-card" id="qcard-${qIndex}">`;
            html += `<div class="question-title"><span class="badge bg-primary rounded-pill me-2 fs-6">Câu ${qIndex + 1}</span>${q.question_text}</div>`;

            if (q.type === 'mc') {
                html += `<div class="mc-options-container">`;
                q.options.forEach((opt, oIndex) => {
                    html += `
                        <div class="mc-option">
                            <input class="mc-input" type="radio" name="q-${qIndex}" value="${opt.is_correct}" id="opt-${qIndex}-${oIndex}">
                            <label class="mc-label" for="opt-${qIndex}-${oIndex}" id="label-mc-${qIndex}-${oIndex}">${opt.text}</label>
                        </div>
                    `;
                });
                html += `</div>`;
            } else if (q.type === 'tf') {
                q.options.forEach((opt, oIndex) => {
                    html += `
                        <div class="tf-option">
                            <div class="tf-text">${opt.text}</div>
                            <div class="btn-group tf-btns" role="group">
                                <input type="radio" class="btn-check" name="q-${qIndex}-opt-${oIndex}" id="t-${qIndex}-${oIndex}" value="true" data-correct="${opt.is_correct}">
                                <label class="btn btn-outline-success fw-bold" for="t-${qIndex}-${oIndex}" id="label-t-${qIndex}-${oIndex}"><i class="bi bi-check-lg"></i> Đúng</label>
                                
                                <input type="radio" class="btn-check" name="q-${qIndex}-opt-${oIndex}" id="f-${qIndex}-${oIndex}" value="false" data-correct="${opt.is_correct}">
                                <label class="btn btn-outline-danger fw-bold" for="f-${qIndex}-${oIndex}" id="label-f-${qIndex}-${oIndex}"><i class="bi bi-x-lg"></i> Sai</label>
                            </div>
                        </div>
                    `;
                });
            }
            
            html += `<div id="feedback-${qIndex}" class="feedback-text shadow-sm"></div></div>`;
            container.innerHTML += html;
        });
        container.dataset.questions = JSON.stringify(questions);
    }

    function submitQuiz() {
        // KHÔNG BẮT BUỘC XÁC NHẬN NỮA. Bấm là chấm luôn!
        const container = document.getElementById("quiz-container");
        const questions = JSON.parse(container.dataset.questions);
        
        let correctMC = 0; let totalMC = 0;
        let correctTF = 0; let totalTF = 0;

        questions.forEach((q, qIndex) => {
            let feedbackBox = document.getElementById(`feedback-${qIndex}`);
            let qCard = document.getElementById(`qcard-${qIndex}`);
            feedbackBox.style.display = "block";
            
            if (q.type === 'mc') {
                totalMC++;
                let selected = document.querySelector(`input[name="q-${qIndex}"]:checked`);
                let isCorrect = selected ? (selected.value === "true") : false;
                
                if (isCorrect) correctMC++;

                q.options.forEach((opt, oIndex) => {
                    let label = document.getElementById(`label-mc-${qIndex}-${oIndex}`);
                    let radio = document.getElementById(`opt-${qIndex}-${oIndex}`);
                    radio.disabled = true;

                    if (opt.is_correct) {
                        label.classList.add("correct-ans");
                    } else if (selected && selected.id === `opt-${qIndex}-${oIndex}` && !isCorrect) {
                        label.classList.add("wrong-ans");
                    }
                });

                if (isCorrect) {
                    qCard.classList.add("graded-correct");
                    feedbackBox.innerHTML = '<i class="bi bi-stars"></i> Tuyệt vời! Bạn chọn rất chính xác.';
                    feedbackBox.className = "feedback-text text-success bg-white border border-success mt-3";
                } else if (!selected) {
                    qCard.classList.add("graded-wrong");
                    feedbackBox.innerHTML = `<i class="bi bi-info-circle-fill"></i> Bạn chưa chọn đáp án. Ý màu Xanh lá là đáp án đúng!`;
                    feedbackBox.className = "feedback-text text-danger bg-white border border-danger mt-3";
                } else {
                    qCard.classList.add("graded-wrong");
                    feedbackBox.innerHTML = `<i class="bi bi-x-circle-fill"></i> Sai rồi! Ý màu Xanh lá mới là đáp án đúng.`;
                    feedbackBox.className = "feedback-text text-danger bg-white border border-danger mt-3";
                }

            } else if (q.type === 'tf') {
                let allCorrectInQuestion = true;
                
                q.options.forEach((opt, oIndex) => {
                    totalTF++;
                    let inputT = document.getElementById(`t-${qIndex}-${oIndex}`);
                    let inputF = document.getElementById(`f-${qIndex}-${oIndex}`);
                    let labelT = document.getElementById(`label-t-${qIndex}-${oIndex}`);
                    let labelF = document.getElementById(`label-f-${qIndex}-${oIndex}`);
                    
                    inputT.disabled = true; inputF.disabled = true;
                    let selected = document.querySelector(`input[name="q-${qIndex}-opt-${oIndex}"]:checked`);
                    
                    if (opt.is_correct) labelT.classList.replace('btn-outline-success', 'btn-success');
                    else labelF.classList.replace('btn-outline-danger', 'btn-danger');

                    if (selected) {
                        let userChoice = selected.value === "true";
                        if (userChoice !== opt.is_correct) {
                            allCorrectInQuestion = false;
                            if (userChoice) labelT.classList.replace('btn-outline-success', 'btn-danger');
                            else labelF.classList.replace('btn-outline-danger', 'btn-danger');
                        } else {
                            correctTF++;
                        }
                    } else {
                        allCorrectInQuestion = false;
                    }
                });

                if (allCorrectInQuestion) {
                    qCard.classList.add("graded-correct");
                    feedbackBox.innerHTML = '<i class="bi bi-stars"></i> Chính xác toàn bộ các ý!';
                    feedbackBox.className = "feedback-text text-success bg-white border border-success mt-3";
                } else {
                    qCard.classList.add("graded-wrong");
                    feedbackBox.innerHTML = `<i class="bi bi-x-circle-fill"></i> Có ý bạn chọn sai (hoặc chưa chọn). Các nút đang sáng màu (Xanh/Đỏ) chính là đáp án chuẩn!`;
                    feedbackBox.className = "feedback-text text-danger bg-white border border-danger mt-3";
                }
            }
        });

        const resBox = document.getElementById("result-container");
        resBox.innerHTML = `
            <div class="display-6 fw-bold mb-3"><i class="bi bi-trophy-fill text-warning"></i> ĐÃ CHẤM ĐIỂM XONG!</div>
            <div class="row text-center fs-5">
                <div class="col-6 border-end border-light">Nhiều lựa chọn:<br><strong class="fs-1 fw-bold">${correctMC}</strong> / ${totalMC}</div>
                <div class="col-6">Đúng/Sai:<br><strong class="fs-1 fw-bold">${correctTF}</strong> / ${totalTF}</div>
            </div>
        `;
        resBox.style.display = "block";
        document.getElementById("submit-area").style.display = "none";
        
        if ((batchIndex + 1) * batchSize < allQuestions.length) {
            document.getElementById("btn-next").disabled = false;
        }
    }

    function retryBatch() {
        loadBatch(true); 
    }

    function nextBatch() {
        batchIndex++;
        loadBatch(true);
    }
    
    function prevBatch() {
        batchIndex--;
        loadBatch(true);
    }
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/questions')
def get_questions():
    if not os.path.exists(DOCX_FILE):
        return jsonify({"error": f"Không tìm thấy file {DOCX_FILE}."})
    questions = parse_docx(DOCX_FILE)
    return jsonify(questions)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
