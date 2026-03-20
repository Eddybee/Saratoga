# Saratoga Application Setup Guide

This guide will help you set up and run the Saratoga application (backend and frontend) for local development and testing.

---

## Prerequisites
- Python 3.9+
- Node.js (v16+ recommended) & npm
- Git

---

## 1. Clone the Repository
```
git clone https://github.com/Eddybee/Saratoga.git
cd Saratoga
```

---

## 2. Backend Setup

### a. Create and Activate Virtual Environment
```
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### b. Install Python Dependencies
```
pip install -r app/backend/requirements.txt
```

### c. Start Backend Server
```
cd app/backend
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

---

## 3. Frontend Setup

### a. Install Node Dependencies
```
cd ../../app/frontend
npm install
```

### b. Start Frontend Dev Server
```
npm run dev -- --host 0.0.0.0
```

---

## 4. Access the Application
- Frontend: http://localhost:5173
- Backend API: http://localhost:8000

---

## 5. Notes
- Ensure both servers are running for full functionality.
- Place any required files (PDFs, Excel, etc.) in the appropriate backend folders as needed.
- For any issues, check the logs or contact the maintainer.
