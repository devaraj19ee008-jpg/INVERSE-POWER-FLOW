# ⚡ Inverse Power Flow System

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![NumPy](https://img.shields.io/badge/NumPy-1.19+-orange.svg)](https://numpy.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Stable-brightgreen.svg)]()

> A modular power system analysis tool using the **Inverse Power Flow Method** for fast, non-iterative power flow calculations

---

## 📌 Overview

Traditional power flow methods (Newton-Raphson, Gauss-Seidel) require iterative solutions and often face convergence issues. This system takes the **inverse approach**: 

> **Given voltage magnitudes and angles → Calculate power injections directly**

This single-step, non-iterative calculation achieves:
- ✅ **100% convergence rate** (never fails)
- ✅ **5-30x faster** than traditional methods
- ✅ **Perfect for Monte Carlo simulation**
- ✅ **Ideal for machine learning training data**

---

## 🎯 Key Features

### Core Capabilities
| Feature | Description |
|---------|-------------|
| **Non-iterative Power Flow** | Single-step calculation with guaranteed convergence |
| **Bidirectional Flows** | Complete from→to and to→from analysis for all branches |
| **Monte Carlo Simulation** | Generate thousands of valid operating samples |
| **Configurable Variation** | 5%, 10%, or 15% voltage variation strength |
| **Automatic Validation** | Built-in sample quality checking |

### Supported Components
| Component | Input Fields | Output Analysis |
|-----------|--------------|-----------------|
| 🔌 **Buses** | Number, name, type, base kV, V_init, angle, shunt G/B, area, zone, owner | Voltage (pu & kV), angle, status (normal/low/high) |
| ⚙️ **Generators** | Name, type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone | P/Q output, limit violation detection |
| 📉 **Loads** | Name, bus, P_demand, Q_demand, model, area, zone | MVA demand, supply status |
| 🔗 **Transmission Lines** | Name, from/to bus, length(km), R/km, X/km, B/km, rating, status, area, zone, owner | P/Q/MVA flows (bidirectional), losses, loss/km, loading % |
| 🔄 **Transformers** | Name, from/to bus, R, X, tap_ratio, rating, phase_shift, min/max tap, step, status, area, zone, owner | Flows, losses, tap position, loading % |
| 🔷 **Capacitors** | Name, bus, Q_rating, status, area, zone | Voltage-dependent Q injection |
| 🔶 **Reactors** | Name, bus, Q_rating, status, area, zone | Voltage-dependent Q absorption |
| ⚡ **Series Compensation** | Name, from/to bus, R, X, comp_percentage, status, area, zone, owner | Effective reactance |
| 🧲 **Series Reactors** | Name, from/to bus, R, X, status, area, zone, owner | Status monitoring |
| 🔰 **Shunt Compensation** | Name, bus, Q_rating, status, area, zone | Voltage-dependent Q injection |

### Output Formats
| Format | Extension | Description |
|--------|-----------|-------------|
| **CSV** | `.csv` | Tabular data for spreadsheet analysis |
| **Python** | `.py` | Full dataset for reproducibility |
| **Text Summary** | `.invout` | Comprehensive formatted report with zone analysis |


### Analysis Features
- 📊 Zone-wise power summaries (generation, load, losses)
- 📈 Zone-to-zone power flow matrices (active & reactive)
- 🔄 Inter-zone power exchange analysis
- ⚠️ Voltage violation detection (configurable thresholds)
- 🚨 Reactive power limit checking
- 📉 Loss calculations (MW, Mvar, MVA)
- 🔧 Transformer tap position tracking
- 📊 Loading percentage analysis

---

## 🏗️ Architecture

The system follows strict **separation of concerns** with four independent modules:



### Design Principles

- 🔒 **Engine NEVER modified** when changing input/output formats
- 🔧 **Input Reader is the only file to customize** for new data sources
- 📝 **Report Writer can be extended** for additional output formats
- 🎯 **Main Runner orchestrates** the entire workflow

---

## 📥 Installation

### Prerequisites
- Python 3.7 or higher
- NumPy 1.19 or higher

### Steps

```bash
# Clone the repository
git clone https://github.com/yourusername/inverse-power-flow.git
cd inverse-power-flow

# Install required package (only NumPy needed)
pip install numpy

🚀 Usage
# Command Line Mode
# Basic usage (10% variation, 200 samples)
python run_power_flow.py -i my_system.py -v 10 -n 200

# With 5% variation and 500 samples
python run_power_flow.py -i my_system.py -v 5 -n 500

# With 15% variation (default 200 samples)
python run_power_flow.py -i my_system.py -v 15

# Minimum arguments (will prompt for missing)
python run_power_flow.py -i my_system.py


## 🖥️ Command Line Arguments

The system accepts three command line arguments for non-interactive operation.

### Argument Summary Table

| Argument | Short | Type | Default | Required | Description |
|----------|-------|------|---------|----------|-------------|
| `--input` | `-i` | String | None | **Yes** | Path to the input Python file containing power system data |
| `--variation` | `-v` | Float | None* | No* | Voltage variation strength as percentage (5, 10, or 15) |
| `--samples` | `-n` | Integer | None* | No* | Number of Monte Carlo samples to generate |

*If not provided, system enters interactive mode to prompt for these values

### Detailed Argument Descriptions

#### `--input` or `-i`

**Description:** Specifies the path to the input Python file that contains all power system component data including buses, generators, loads, lines, transformers, and optional components.

**Format:** String path (absolute or relative)

**Example:**
```bash
python run_power_flow.py -i "E:\power_systems\case39.py"
python run_power_flow.py -i ./data/my_system.py

python run_power_flow.py

================================================================================
INTERACTIVE MODE
================================================================================

Enter Python input file path:
(Supports AREA, ZONE, OWNER fields)

Path: my_system.py

================================================================================
SIMULATION PARAMETERS
================================================================================

Variation Strength:
  [1] Small (5%)
  [2] Medium (10%) - Recommended
  [3] Large (15%)

Choice (1-3) [default: 2]: 2

Number of samples [default: 200]: 200

# Output Formate
my_system/
├── my_system_var10_20241201_120000_ALL_DATA.csv      # Complete CSV data
├── my_system_var10_20241201_120000_ALL_DATA.py       # Full Python dataset
├── my_system_var10_20241201_120000_SUMMARY.invout    # Text summary report


