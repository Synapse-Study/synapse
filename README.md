# Synapse 🧠

**AI-Powered Medical Study Studio & Question Bank for King's College London MBBS**

Synapse is an intelligent, offline-capable medical education workstation engineered for KCL MBBS students. It automates high-yield Anki flashcard drafting, performs on-device diagram optical character recognition (OCR) for Image Occlusion, synthesizes clinical lecture notes with automatic diagram cropping, and connects directly to the Synapse SBA Question Bank.

---

## ✨ Features

- **🎯 MBBS Calibrated Flashcard Studio:** Generates atomic Cloze and Basic flashcards strictly grounded in Discrete Learning Outcomes. Automatically excludes administrative clutter, Nobel prizes, and non-examinable trivia.
- **🖼️ Offline Image Occlusion Engine:** On-device deep learning computer vision stack (PyTorch, CRAFT, EasyOCR, OpenCV) that auto-detects labels and anatomical structures on lecture diagrams without sending slide images to cloud servers.
- **📝 Clinical Lecture Notes Workspace:** Automatically extracts and aligns high-resolution slide diagrams with Notion notes, stripping headers and footer citations with millimeter precision.
- **🐦 AI Medical Tutor (Dr. P.I. Geon):** King's College London MBBS curriculum-grounded medical tutor with deep reasoning across Foundations of Biomedical Sciences (FBS), Physiology and Anatomy of Systems (PAS), and Values-Based Clinical Practice (IVBCP).
- **🔒 Secure Architecture:** Zero credentials bundled. Integrates with Supabase Auth (server-side fail-closed whitelist) and stores credentials securely in macOS Keychain.
- **⚡ GitHub Release Updater:** Automatically checks for and alerts you to new updates from [Synapse-Study/synapse](https://github.com/Synapse-Study/synapse).

---

## 🚀 Getting Started

### Download for macOS (Apple Silicon)

Download the latest disk image from [GitHub Releases](https://github.com/Synapse-Study/synapse/releases/latest):

1. Download **`Synapse.dmg`**.
2. Double-click the `.dmg` and drag **Synapse** into your **Applications** folder.
3. Open Synapse from Applications (or Spotlight).

> [!NOTE]
> Synapse is natively compiled for **Apple Silicon Macs (M1/M2/M3/M4)** running macOS 13 (Ventura) or later.

---

## 🛠️ Development & Building from Source

### Prerequisites
- macOS with Python 3.12+
- Anki with [AnkiConnect](https://foosoft.net/projects/anki-connect/) installed (optional, for direct Anki sync)

### Setup Virtual Environment
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Run Locally
```bash
./run_app.sh
```

### Build macOS App & DMG
```bash
./build_app.sh
```
This will compile the standalone `dist/Synapse.app` and generate the compressed `dist/Synapse.dmg`.

---

## 📄 License
Internal medical study suite developed for King's College London MBBS.
