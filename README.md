# Evalsea – Learning Management System

Evalsea is a Learning Management System (LMS) designed to support academic workflows by connecting parents, students, faculty, and evaluation processes through an integrated platform.

## 🚀 Features

- **Parent Portal:** Provides parent-oriented access to student-related information.
- **Faculty Module:** Supports faculty-related academic workflows.
- **Evaluation Module:** Handles evaluation-related functionality.
- **Assignment & Assessment Generation:** Includes a generator module for academic content workflows.
- **Duplicate Detection:** Includes a duplicate detection engine to identify potentially repeated content.
- **Backend Services:** Python-based backend implementation.

## 🛠️ Technology Stack

- **Programming Language:** Python
- **Backend:** Python
- **Testing:** Pytest / Jupyter Notebook (where applicable)
- **Dependencies:** Python packages listed in `requirements.txt`

## 📁 Project Structure

```text
evalsea/
│
├── backend/
├── evalsea-duplicate-detection-engine/
├── evaluator/
├── faculty/
├── generator/
│
├── main.py
├── requirements.txt
├── test.ipynb
├── test_add_child.py
└── .gitignore
```

## ⚙️ Installation & Setup

### 1. Clone the repository

```bash
git clone https://github.com/vishuuu09/evalsea.git
cd evalsea
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

### 3. Activate the virtual environment

**Windows:**
```bash
venv\Scripts\activate
```

**Linux / macOS:**
```bash
source venv/bin/activate
```

### 4. Install dependencies

```bash
pip install -r requirements.txt
```

### 5. Run the application

```bash
python main.py
```

Refer to the individual module instructions if additional configuration or services are required.

## 🧪 Testing

Run the available Python test:

```bash
python test_add_child.py
```

Jupyter Notebook tests can be explored using:

```bash
jupyter notebook test.ipynb
```

## 🔮 Future Improvements

- Enhanced student and parent dashboards.
- Improved academic performance tracking.
- Advanced duplicate detection capabilities.
- Better evaluation and assignment workflows.
- Improved API integration and scalability.

## 👨‍💻 Maintainer

**Sai Vishal**

GitHub: [@vishuuu09](https://github.com/vishuuu09)

---

*This repository builds upon an existing LMS codebase. Original authorship and applicable licensing should be acknowledged. Contributions and modifications should be documented transparently.*
