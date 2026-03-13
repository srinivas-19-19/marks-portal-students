# Student Marks Portal (SVCE)

A fast, simple, and responsive React application for faculty to process student marks and generate formal Award Lists.

## ✨ Features

- **Excel Import**: Automatically import student lists from Excel files.
- **Smart Calculations**:
  - Automatically picks `max(Q1, Q2)`, `max(Q3, Q4)`, and `max(Q5, Q6)`.
  - Calculates `ceil(Total / 2)` for Descriptive marks.
  - Adds Objective marks for the Final score.
- **Two-Mid Consolidation**: Optional mode for weighted mid-exam results (`max*0.8 + min*0.2`).
- **Formal Award List**: Generates a print-ready document formatted exactly for SV College of Engineering (Autonomous).
- **Mobile Responsive**: Fully adaptive design for use on both desktop and mobile devices.

## 🚀 Getting Started

### Prerequisites
- Node.js (v18 or higher)
- npm or yarn

### Installation
1. Clone the repository:
   ```bash
   git clone <your-repo-link>
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Run the development server:
   ```bash
   npm run dev
   ```

## 🛠 Tech Stack
- **React** (Vite)
- **Lucide React** (Icons)
- **XLSX** (Excel Parsing)
- **CSS3** (Custom Responsive Layout)

---
*Built for SV College of Engineering Evaluation Process.*
