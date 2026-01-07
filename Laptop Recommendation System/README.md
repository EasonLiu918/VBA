# Laptop Recommendation System

A powerful Excel-VBA system for assisting users in selecting the most suitable laptop based on customizable criteria. Designed for retail stores and internal staff, this tool integrates benchmark scores, filtering logic, and optimization using Solver.

---

## Features

- **VBA-Powered Logic** — Fully implemented using Excel VBA with user-friendly forms and interfaces.
- **Smart Filtering** — Automatically filters laptops based on:
  - Brand
  - Usage type (e.g., Gaming, Business)
  - Screen size range
  - Resolution requirement
  - Condition (New, Refurbished, Open Box)
  - Price range
- **Hardware Scoring System**
  - CPU and GPU benchmark values are normalized to a 1-10 score.
  - Dynamic recalculation based on inserted data.
- **Best Laptop Selection Modes**
  - **Performance Mode** – Finds highest CPU + GPU scores.
  - **Cost-Effectiveness Mode** – Balances performance with price.
  - **Weighted Solver Mode** – Customize importance of CPU vs GPU and use Solver to find the optimal choice.
- **Password-Protected Admin Access** — Modify inventory only with staff password.
- **Form-Based Data Entry** — Add new inventory, CPU, and GPU records via modern user forms.

---

## File Structure

- `Group_11_A2_RecomSystem.xlsm` — Main Excel file with all VBA modules and forms.
- **Sheets Included**:
  - `CoverPage` – Main UI and instructions
  - `Inventory` – Laptop database
  - `Main` – Filtered results and best choice display
  - `CPU_Benchmark` – CPU scores (Benchmark + Scoring)
  - `GPU_Benchmark` – GPU scores (Benchmark + Scoring)

---

## Solver Logic (Weighted Optimization)

A best option is selected by solving:

```math
Maximize: Σ (CPU_weight × CPU_score + GPU_weight × GPU_score)
