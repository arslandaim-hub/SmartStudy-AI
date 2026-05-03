from flask import Flask, jsonify, render_template, request, flash, session, send_file, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import requests
from dotenv import load_dotenv
import os
import re
import io
import base64
from datetime import datetime

# ORIGINAL SPECIALIZED IMPORTS
import img2pdf
from PyPDF2 import PdfReader
import yt_dlp
from faster_whisper import WhisperModel

# ADDED FOR POWERPOINT GENERATION
from pptx import Presentation
from pptx.util import Inches

# ChromaDB and Local Embeddings
import chromadb
from chromadb.utils import embedding_functions

load_dotenv()

app = Flask(__name__)
app.secret_key = "smartstudy-secret-key-2026"

# ==================== DATABASE & AUTH CONFIGURATION ====================
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    subjects = db.relationship('Subject', backref='user', lazy=True)

# NEW MODELS FOR DAILY STUDY TOOL
class Subject(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    lectures = db.relationship('Lecture', backref='subject', cascade="all, delete-orphan", lazy=True)

class Lecture(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    topic = db.Column(db.String(200), nullable=False)
    sub_topics = db.Column(db.Text)  # Detailed material or sub-points
    ai_concepts = db.Column(db.Text) # AI generated core concepts
    ai_quiz = db.Column(db.Text)     # AI generated MCQs
    date_added = db.Column(db.DateTime, default=datetime.utcnow)
    subject_id = db.Column(db.Integer, db.ForeignKey('subject.id'), nullable=False)

with app.app_context():
    db.create_all()

# ==================== DIRECTORIES & MODELS ====================
OLLAMA_URL = "http://localhost:11434/api/generate"
DOWNLOADS_DIR = "downloads"
os.makedirs(DOWNLOADS_DIR, exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# Initialize Faster-Whisper
model_size = "base"
whisper_model = WhisperModel(model_size, device="cpu", compute_type="int8")

# --- LOCAL EMBEDDING ENGINE ---
model_path = os.path.abspath("./local_model")
try:
    local_ef = embedding_functions.SentenceTransformerEmbeddingFunction(
        model_name=model_path,
        device="cpu"
    )
    print("✅ Local embedding model loaded successfully.")
except Exception as e:
    print(f"❌ Error loading local model: {e}")
    local_ef = embedding_functions.DefaultEmbeddingFunction()

db_client = chromadb.PersistentClient(path="./study_db")

# Helper to keep user data separate
def get_user_collection():
    if 'user_email' in session:
        safe_email = re.sub(r'[^a-zA-Z0-9]', '_', session['user_email'])
        collection_name = f"user_data_{safe_email}"
    else:
        collection_name = "guest_data_collection"
        
    return db_client.get_or_create_collection(
        name=collection_name, 
        embedding_function=local_ef
    )

# ==================== CORE LLM FUNCTION ====================

def ask_local_llm(prompt, model_type="fast", image_data=None):
    if image_data:
        selected_model = "llava"
    else:
        selected_model = "gemma2:9b" if model_type == "heavy" else "llama3.2"
    
    payload = {"model": selected_model, "prompt": prompt, "stream": False}
    
    if image_data:
        payload["images"] = [image_data]

    try:
        timeout_val = 150 if (model_type == "heavy" or image_data) else 60
        response = requests.post(OLLAMA_URL, json=payload, timeout=timeout_val)
        return response.json().get('response', "Error: Empty response.")
    except Exception as e:
        return f"Error: {str(e)}"

def chunk_text(text, size=1000):
    return [text[i:i+size] for i in range(0, len(text), size)]

# ==================== AUTH ROUTES ====================

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        name = request.form.get('name')
        email = request.form.get('email')
        password = request.form.get('password')
        user_exists = User.query.filter_by(email=email).first()
        if user_exists:
            flash("Email already exists!", "danger")
        else:
            hashed_pw = generate_password_hash(password, method='pbkdf2:sha256')
            new_user = User(name=name, email=email, password=hashed_pw)
            db.session.add(new_user)
            db.session.commit()
            flash("Signup successful! Please login.", "success")
            return redirect(url_for('login'))
    return render_template('signup.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        user = User.query.filter_by(email=email).first()
        if user and check_password_hash(user.password, password):
            session['user_id'] = user.id
            session['user_name'] = user.name
            session['user_email'] = user.email
            flash(f"Welcome back, {user.name}!", "success")
            return redirect(url_for('index'))
        else:
            flash("Login failed. Check your email and password.", "danger")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for('index'))

# ==================== NEW FEATURE: DAILY COURSE COMPENDIUM ====================

# ... [Keep all previous imports and model definitions exactly as they are] ...

# ==================== NEW FEATURE: DAILY COURSE COMPENDIUM ====================

@app.route('/courses', methods=['GET', 'POST'])
def courses():
    if 'user_id' not in session:
        flash("Please login to use Course Compendium", "warning")
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        subject_name = request.form.get('subject_name')
        if subject_name:
            new_sub = Subject(name=subject_name, user_id=session['user_id'])
            db.session.add(new_sub)
            db.session.commit()
            flash(f"Subject '{subject_name}' added!", "success")

    user_subjects = Subject.query.filter_by(user_id=session['user_id']).all()
    return render_template('courses.html', subjects=user_subjects)

@app.route('/rename_subject/<int:subject_id>', methods=['POST'])
def rename_subject(subject_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    subject = Subject.query.get_or_404(subject_id)
    # Security: Ensure subject belongs to current user
    if subject.user_id != session['user_id']:
        flash("Unauthorized action.", "danger")
        return redirect(url_for('courses'))

    new_name = request.form.get('new_name')
    if new_name:
        subject.name = new_name
        db.session.commit()
        flash("Folder renamed successfully!", "success")
    
    return redirect(url_for('courses'))

@app.route('/delete_subject/<int:subject_id>')
def delete_subject(subject_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    subject = Subject.query.get_or_404(subject_id)
    # Security: Ensure subject belongs to current user
    if subject.user_id != session['user_id']:
        flash("Unauthorized action.", "danger")
        return redirect(url_for('courses'))

    db.session.delete(subject)
    db.session.commit()
    flash("Folder and all contained lectures deleted.", "info")
    return redirect(url_for('courses'))

# Lecture Delection Function....

@app.route('/delete_lecture/<int:lecture_id>')
def delete_lecture(lecture_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    lecture = Lecture.query.get_or_404(lecture_id)
    subject_id = lecture.subject_id
    
    # Security check: Ensure the lecture belongs to a subject owned by the user
    if lecture.subject.user_id != session['user_id']:
        flash("Unauthorized action.", "danger")
        return redirect(url_for('courses'))

    db.session.delete(lecture)
    db.session.commit()
    flash("Topic deleted successfully.", "info")
    return redirect(url_for('view_subject', subject_id=subject_id))


@app.route('/subject/<int:subject_id>', methods=['GET', 'POST'])
def view_subject(subject_id):
    subject = Subject.query.get_or_404(subject_id)
    if request.method == 'POST':
        topic = request.form.get('topic')
        sub_topics = request.form.get('sub_topics')
        action = request.form.get('action') # 'save' or 'save_gen'

        concepts = ""
        quiz = ""

        # Only generate if the user clicked the "Save & Generate" button
        if action == 'save_gen':
            concept_prompt = f"Given these lecture sub-topics about {topic}: {sub_topics}. Extract the most important core concepts in bullet points."
            concepts = ask_local_llm(concept_prompt)
            
            quiz_prompt = f"Generate 3 MCQs based on these lecture notes: {sub_topics}. Topic: {topic}. Include correct answers."
            quiz = ask_local_llm(quiz_prompt)

        new_lecture = Lecture(
            topic=topic, 
            sub_topics=sub_topics, 
            ai_concepts=concepts,
            ai_quiz=quiz,
            subject_id=subject_id
        )
        db.session.add(new_lecture)
        db.session.commit()
        
        msg = "Work logged!" if action == 'save' else "Work logged and AI kit generated!"
        flash(msg, "success")

    return render_template('view_subject.html', subject=subject)

@app.route('/generate_kit/<int:lecture_id>')
def generate_kit(lecture_id):
    lecture = Lecture.query.get_or_404(lecture_id)
    
    # Generate content
    concept_prompt = f"Given these lecture notes about {lecture.topic}: {lecture.sub_topics}. Extract core concepts in bullet points."
    lecture.ai_concepts = ask_local_llm(concept_prompt)
    
    quiz_prompt = f"Generate 3 MCQs based on these lecture notes: {lecture.sub_topics}. Topic: {lecture.topic}."
    lecture.ai_quiz = ask_local_llm(quiz_prompt)
    
    db.session.commit()
    flash(f"AI Study Kit for {lecture.topic} updated!", "success")
    return redirect(url_for('view_subject', subject_id=lecture.subject_id))

# ==================== RESTORED FUNCTIONAL ROUTES ====================

@app.route('/')
def index():
    return render_template('index.html', user_name=session.get('user_name'))

@app.route('/diagram')
def diagram_studio():
    return render_template('diagram.html')

@app.route('/diagram-ai', methods=['POST'])
def diagram_ai():
    data = request.json
    user_prompt = data.get('prompt', '')

    system_instruction = """
        Generate ONLY Mermaid.js code starting with 'graph TD'.
        RULES:
        1. Arrow labels MUST use the pipe format: A -->|Label Text| B.
        2. Node IDs must be single letters or underscores (A, B_1).
        3. Use brackets for visible text: A[Egg] -->|Hatching| B[Tadpole].
        4. NO colons (:) for descriptions. NO markdown backticks.
    """

    full_prompt = f"{system_instruction}\n\nUser Request: {user_prompt}"
    mermaid_code = ask_local_llm(full_prompt, model_type="fast")

    cleaned = re.sub(r'```mermaid|```', '', mermaid_code).strip()
    
    match = re.search(r"(graph|flowchart)[\s\S]+", cleaned)
    if match:
        cleaned = match.group(0).strip()

    cleaned = re.sub(r"(\S+)\s*-->\s*(\S+)\s*:\s*(.*)", r"\1 -->|\3| \2", cleaned)

    return jsonify({"code": cleaned})

@app.route('/diagram-sketch', methods=['POST'])
def diagram_sketch():
    if 'sketch' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['sketch']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    try:
        image_data = base64.b64encode(file.read()).decode('utf-8')
        vision_prompt = """
            Analyze this sketch and convert it into Mermaid.js code.
            Start with 'graph TD'. Use brackets for labels like A[Start].
            Output ONLY the Mermaid code, no other text.
        """
        mermaid_code = ask_local_llm(vision_prompt, image_data=image_data)
        cleaned = re.sub(r'```mermaid|```', '', mermaid_code).strip()
        match = re.search(r"(graph|flowchart)[\s\S]+", cleaned)
        if match:
            cleaned = match.group(0).strip()

        return jsonify({
            "code": cleaned,
            "message": "Sketch processed using Llava!"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/get_response', methods=['POST'])
def get_response():
    user_name = session.get('user_name', 'Guest')
    
    if request.is_json:
        data = request.get_json()
        user_msg = data.get('message', '')
        image_file = None
    else:
        user_msg = request.form.get('message', '')
        image_file = request.files.get('image')
    
    if not user_msg and not image_file:
        return jsonify({"response": "How can I help?"})
    
    base64_image = None
    if image_file:
        try:
            base64_image = base64.b64encode(image_file.read()).decode('utf-8')
        except Exception as e:
            return jsonify({"response": f"Error processing image: {str(e)}"})
    
    collection = get_user_collection()
    results = collection.query(query_texts=[user_msg], n_results=3)
    context = "\n".join(results['documents'][0]) if results['documents'] else ""
    
    if session.get('last_summary'):
        context += f"\nVideo Summary Context: {session.get('last_summary')}"
    
    if base64_image:
        prompt = f"Analyze this image. User: {user_name}. Question: {user_msg}"
    else:
        prompt = f"Use this context to answer. User: {user_name}. Context: {context}\n\nQuestion: {user_msg}"
    
    reply = ask_local_llm(prompt, model_type="fast", image_data=base64_image)
    return jsonify({"response": reply})

@app.route('/notes', methods=['GET', 'POST'])
def notes():
    answer = None
    question = None
    collection = get_user_collection()
    
    if request.method == 'POST':
        if 'pdf' in request.files:
            file = request.files['pdf']
            if file and file.filename.endswith('.pdf'):
                filepath = os.path.join("uploads", file.filename)
                file.save(filepath)
                reader = PdfReader(filepath)
                text = "".join([page.extract_text() or "" for page in reader.pages])
                chunks = chunk_text(text)
                ids = [f"{file.filename}_{i}" for i in range(len(chunks))]
                collection.upsert(
                    documents=chunks, 
                    ids=ids, 
                    metadatas=[{"source": file.filename} for _ in chunks]
                )
                session['notes_uploaded'] = True
                flash(f"✅ {file.filename} indexed to your private library!", "success")
        
        elif 'question' in request.form:
            question = request.form.get('question')
            results = collection.query(query_texts=[question], n_results=4)
            context = "\n---\n".join(results['documents'][0])
            prompt = f"Based on these study notes:\n{context}\n\nQuestion: {question}"
            answer = ask_local_llm(prompt, model_type="fast")
            
    return render_template('notes.html', answer=answer, question=question)

@app.route('/images-to-pdf', methods=['GET', 'POST'])
def images_to_pdf():
    if request.method == 'POST':
        uploaded_files = request.files.getlist("images")
        output_format = request.form.get("format")
        
        if not uploaded_files:
            flash("No images selected", "danger")
            return redirect(request.url)

        try:
            img_bytes_list = [file.read() for file in uploaded_files]
            
            if output_format == 'ppt':
                prs = Presentation()
                prs.slide_width = Inches(13.33)
                prs.slide_height = Inches(7.5)
                for img_bytes in img_bytes_list:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    image_stream = io.BytesIO(img_bytes)
                    prs.slides[prs.slides.index(slide)].shapes.add_picture(
                        image_stream, 0, 0, width=prs.slide_width
                    )
                out = io.BytesIO()
                prs.save(out)
                out.seek(0)
                return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation', as_attachment=True, download_name='SmartStudy_Slides.pptx')
            else:
                pdf_bytes = img2pdf.convert(img_bytes_list)
                return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf', as_attachment=True, download_name='SmartStudy_Images.pdf')
        except Exception as e:
            flash(f"Error: {str(e)}", "danger")
            
    return render_template('images_to_pdf.html')

@app.route('/planner', methods=['GET', 'POST'])
def planner():
    plan = None
    if request.method == 'POST':
        subjects = request.form.get('subjects', '')
        hours = request.form.get('hours', '4')
        days = request.form.get('days', '7')
        deadline = request.form.get('deadline', '')
        prompt = f"Create a study plan for: {subjects} for {hours}h/day over {days} days. Deadline: {deadline}."
        plan = ask_local_llm(prompt, model_type="heavy")
    return render_template('planner.html', plan=plan)

@app.route('/quiz', methods=['GET', 'POST'])
def quiz():
    quiz_content = None
    if request.method == 'POST':
        topic = request.form.get('topic', '')
        num = request.form.get('num', '5')
        prompt = f"Generate {num} MCQs on {topic}. Format: Q, A/B/C/D, Correct Answer."
        quiz_content = ask_local_llm(prompt, model_type="fast")
    return render_template('quiz.html', quiz=quiz_content)

@app.route('/summarize', methods=['GET', 'POST'])
def summarize():
    summary = None
    video_id = None
    if request.method == 'POST':
        video_url = request.form.get('video_url', '')
        match = re.search(r"(?:v=|\/)([0-9A-Za-z_-]{11})", video_url)
        if match:
            video_id = match.group(1)
            try:
                ydl_opts = {
                    'format': 'm4a/bestaudio/best', 
                    'outtmpl': os.path.join(DOWNLOADS_DIR, '%(id)s.%(ext)s'), 
                    'quiet': True
                }
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(video_url, download=True)
                    audio_path = ydl.prepare_filename(info)
                segments, _ = whisper_model.transcribe(audio_path, beam_size=5)
                full_text = " ".join([segment.text for segment in segments])
                if os.path.exists(audio_path): 
                    os.remove(audio_path)
                prompt = f"Summarize this lecture transcript: {full_text[:7000]}"
                summary = ask_local_llm(prompt, model_type="fast")
                session['last_summary'] = summary
            except Exception as e:
                flash(f"Error: {str(e)}", "danger")
    return render_template('summary.html', summary=summary, video_id=video_id)

@app.route('/chat', methods=['GET', 'POST'])
def chat():
    user_name = session.get('user_name', 'Guest')
    if request.method == 'POST':
        if request.is_json:
            data = request.get_json()
            user_msg = data.get('message')
            image_file = None
        else:
            user_msg = request.form.get('message')
            image_file = request.files.get('image')

        if not user_msg and not image_file:
            return jsonify({"response": "How can I help?"})

        base64_image = None
        if image_file:
            try:
                base64_image = base64.b64encode(image_file.read()).decode('utf-8')
            except Exception as e:
                return jsonify({"response": f"Image Error: {str(e)}"})

        summary_context = session.get('last_summary', '')
        if base64_image:
            final_prompt = f"Analyze this image. User: {user_name}. Question: {user_msg}"
        else:
            final_prompt = f"Context: {summary_context}\n\nUser {user_name}: {user_msg}" if summary_context else f"User {user_name}: {user_msg}"

        reply = ask_local_llm(final_prompt, model_type="fast", image_data=base64_image)       
        return jsonify({"response": reply})       
    return render_template('chat.html', user_name=user_name)

if __name__ == '__main__':
    app.run(port=5000)