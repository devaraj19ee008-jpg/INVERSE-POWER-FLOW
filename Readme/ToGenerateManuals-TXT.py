"""
FALLBACK: Creates TXT documents if python-docx is not available
"""

import os
from datetime import datetime

def create_all_txt():
    
    # DOCUMENT 1
    with open('01_Main_Calculations.txt', 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("POWER FLOW CALCULATION METHODOLOGY\n")
        f.write("Complete Mathematical Formulations\n")
        f.write("="*80 + "\n\n")
        
        f.write("1. ADMITTANCE MATRIX (Ybus) CONSTRUCTION\n")
        f.write("-"*40 + "\n\n")
        
        f.write("1.1 Transmission Line π-Model\n")
        f.write("""
For a line between buses i and j:
    Z_series = R + jX
    y_series = 1/Z_series = 1/(R + jX) = (R - jX)/(R² + X²)
    y_shunt = jB/2 (per side)

Ybus Contributions:
    Ybus[i][i] += y_series + jB
    Ybus[j][j] += y_series + jB
    Ybus[i][j] -= y_series
    Ybus[j][i] -= y_series
""")
        
        f.write("1.2 Transformer Model\n")
        f.write("""
For transformer with tap ratio 'a':
    Ybus[i][i] += y_series / a²
    Ybus[j][j] += y_series
    Ybus[i][j] -= y_series / a
    Ybus[j][i] -= y_series / a
""")
        
        f.write("\n2. POWER INJECTION CALCULATIONS\n")
        f.write("-"*40 + "\n\n")
        f.write("""
Fundamental Equation:
    S = V × I*
    
Where I = Ybus × V

Therefore:
    S = V × (Ybus × V)*

In matrix form:
    I_injection = Ybus @ V_complex
    S_complex = V_complex × conj(I_injection)
    
    P_MW = real(S) × baseMVA
    Q_Mvar = imag(S) × baseMVA

SIGN CONVENTION:
    Positive P,Q → Power INTO bus (generation)
    Negative P,Q → Power OUT of bus (load)
""")
        
        f.write("\n3. BRANCH FLOW CALCULATIONS\n")
        f.write("-"*40 + "\n\n")
        
        f.write("3.1 Transmission Line Flows\n")
        f.write("""
Forward Flow (i → j):
    I_series = (V_i - V_j) × y
    I_shunt = V_i × j(B/2)
    S_fwd = V_i × (I_series + I_shunt)*

Reverse Flow (j → i):
    I_series = (V_j - V_i) × y
    I_shunt = V_j × j(B/2)
    S_rev = V_j × (I_series + I_shunt)*

Losses = S_fwd + S_rev
""")
        
        f.write("3.2 Transformer Flows\n")
        f.write("""
Forward Flow (i → j, i is tap side):
    I_fwd = (V_i/a - V_j) × y
    S_fwd = V_i × I_fwd*

Reverse Flow (j → i):
    I_rev = (V_j × a - V_i) × (y / a²)
    S_rev = V_j × I_rev*

Losses = S_fwd + S_rev
""")
        
        f.write("\n4. LOADING CALCULATIONS\n")
        f.write("-"*40 + "\n\n")
        f.write("""
MVA Flow = √(P² + Q²)
Loading_% = (MVA_flow / rateA) × 100

Status:
    > 100%: OVERLOADED
    > 80%:  HIGH
    > 60%:  MODERATE
    else:   NORMAL

Tap Position:
    tap_step = int((tap_ratio - min_tap) / step_size)
""")
        
        f.write("\n5. SHUNT COMPENSATION\n")
        f.write("-"*40 + "\n\n")
        f.write("""
Capacitor/Reactor output varies with V²:
    Q_actual = Q_rated × V_actual²
""")
        
        f.write("\n6. COMPLETE PYTHON IMPLEMENTATION\n")
        f.write("-"*40 + "\n\n")
        f.write("""
import numpy as np

def build_ybus(n, branches, idx_map):
    Y = np.zeros((n,n), dtype=complex)
    for br in branches:
        i, j = idx_map[br['from_bus']], idx_map[br['to_bus']]
        y = 1.0 / complex(br['r'], br['x'])
        if br['ratio'] == 0:
            Y[i,i] += y + complex(0, br['b'])
            Y[j,j] += y + complex(0, br['b'])
            Y[i,j] -= y; Y[j,i] -= y
        else:
            a = br['ratio']
            Y[i,i] += y/a**2; Y[j,j] += y
            Y[i,j] -= y/a; Y[j,i] -= y/a
    return Y

def solve_power_flow(Y, V_mag, V_deg, baseMVA):
    V = V_mag * np.exp(1j * np.radians(V_deg))
    I = Y @ V
    S = V * np.conj(I)
    return S.real * baseMVA, S.imag * baseMVA

def calc_flows(Y, branches, idx_map, V, baseMVA):
    flows = []
    for br in branches:
        i, j = idx_map[br['from_bus']], idx_map[br['to_bus']]
        Vi, Vj = V[i], V[j]
        y = 1.0 / complex(br['r'], br['x'])
        if br['ratio'] == 0:
            S_fwd = Vi * np.conj((Vi-Vj)*y + Vi*complex(0, br['b']/2))
            S_rev = Vj * np.conj((Vj-Vi)*y + Vj*complex(0, br['b']/2))
        else:
            a = br['ratio']
            S_fwd = Vi * np.conj((Vi/a - Vj)*y)
            S_rev = Vj * np.conj((Vj*a - Vi)*(y/a**2))
        flows.append({
            'P_fwd': S_fwd.real * baseMVA,
            'Q_fwd': S_fwd.imag * baseMVA,
            'P_rev': S_rev.real * baseMVA,
            'loss': (S_fwd + S_rev).real * baseMVA
        })
    return flows
""")
    
    # DOCUMENT 2
    with open('02_User_Manual.txt', 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("INVERSE POWER FLOW SYSTEM - USER MANUAL\n")
        f.write("="*80 + "\n\n")
        
        f.write("1. INTRODUCTION\n")
        f.write("-"*40 + "\n")
        f.write("""
This system calculates power flows from given voltage profiles.
It generates multiple random samples for statistical analysis.

System Files:
  frompypowerformatetothis.py - Converts PYPOWER files
  input_reader.py             - Reads input files
  power_flow_engine.py        - Core calculations (DO NOT MODIFY)
  report_writer.py            - Generates reports
  run_power_flow.py           - Main execution script
""")
        
        f.write("\n2. INPUT FORMAT\n")
        f.write("-"*40 + "\n")
        f.write("""
Required variables in input .py file:

BASE_MVA = 100

BUS_DATA = [
    [bus_num, name, type, base_kV, V_init, angle, G, B, area, zone, owner],
    ...
]

GENERATOR_DATA = [
    [gen_num, name, type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone],
    ...
]

LOAD_DATA = [
    [load_num, name, bus, P_demand, Q_demand, model, area, zone],
    ...
]

LINE_DATA = [
    [line_num, name, from_bus, to_bus, length_km, R_per_km, X_per_km, B_per_km, rateA, status, area, zone, owner],
    ...
]

TRANSFORMER_DATA = [
    [xfmr_num, name, from_bus, to_bus, R_pu, X_pu, tap_ratio, rateA, phase_shift, min_tap, max_tap, step_size, status, area, zone, owner],
    ...
]

Bus Types: 1=PQ, 2=PV, 3=Slack
Units: V in pu, angle in degrees, R/X/B in pu, P/Q in MW/Mvar
""")
        
        f.write("\n3. HOW TO RUN\n")
        f.write("-"*40 + "\n")
        f.write("""
Interactive Mode:
    python run_power_flow.py
    (Follow prompts)

Command Line Mode:
    python run_power_flow.py -i file.py -v 10 -n 200
    
    Arguments:
    -i : Input file path
    -v : Variation (5, 10, or 15 percent)
    -n : Number of samples
""")
        
        f.write("\n4. OUTPUT FILES\n")
        f.write("-"*40 + "\n")
        f.write("""
Created in folder named after input file:

*_ALL_DATA.csv       - Complete data in CSV
*_ALL_DATA.py        - Complete data as Python module
*_OUTPUTS_ONLY.py    - Only calculation results
*_SUMMARY.txt        - System summary

Output Fields:
  Bus: V_pu, V_kV, angle, voltage_status
  Generator: P_out, Q_out, V_actual, Q_status
  Line: P_fwd, Q_fwd, MVA_fwd, P_rev, Q_rev, losses, loading_%
  Transformer: Same + tap_ratio, tap_step
""")
        
        f.write("\n5. CONVERTING PYPOWER FILES\n")
        f.write("-"*40 + "\n")
        f.write("""
    python frompypowerformatetothis.py
    Choose 1 (single) or 2 (batch mode)
    Output: converted_<filename>.py
""")
        
        f.write("\n6. TROUBLESHOOTING\n")
        f.write("-"*40 + "\n")
        f.write("""
File not found:        Check path, use "E:\\\\path\\\\file.py"
No valid samples:      Reduce variation, check for zero impedances
NaN in output:         Check for disconnected buses
Loading > 100%:        Verify rateA values are correct
""")
    
    # DOCUMENT 3
    with open('03_Simplified_Summary.txt', 'w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("POWER FLOW SYSTEM - QUICK REFERENCE (All 5 Files)\n")
        f.write("="*80 + "\n\n")
        
        f.write("FILE 1: frompypowerformatetothis.py\n")
        f.write("-"*40 + "\n")
        f.write("Purpose: Convert PYPOWER → Custom format\n")
        f.write("Key: convert_pypower_to_custom(file_path)\n")
        f.write("Usage: Run → Choose 1 or 2 → Provide path\n\n")
        
        f.write("FILE 2: input_reader.py\n")
        f.write("-"*40 + "\n")
        f.write("Purpose: Parse input files for engine\n")
        f.write("Key: read_from_python() → Data dict\n")
        f.write("     convert_to_engine_format() → bus_data, branch_data\n\n")
        
        f.write("FILE 3: power_flow_engine.py [DO NOT MODIFY]\n")
        f.write("-"*40 + "\n")
        f.write("Purpose: Core calculations\n")
        f.write("Key Formulas:\n")
        f.write("  Ybus: y=1/(R+jX), lines add jB, xfmr divide by tap²\n")
        f.write("  Power: S = V × (Ybus × V)*\n")
        f.write("  Line flow: S_fwd = Vi × [(Vi-Vj)×y + Vi×jB/2]*\n")
        f.write("  Xfmr flow: S_fwd = Vi × [(Vi/a - Vj)×y]*\n\n")
        
        f.write("FILE 4: report_writer.py\n")
        f.write("-"*40 + "\n")
        f.write("Purpose: Generate output files\n")
        f.write("Methods: write_csv_all_data(), write_py_all_data(),\n")
        f.write("         write_py_output_only(), write_txt_summary()\n")
        f.write("Outputs: Bus V/angle, Gen P/Q, Load supply, Line/Xfmr flows\n\n")
        
        f.write("FILE 5: run_power_flow.py\n")
        f.write("-"*40 + "\n")
        f.write("Purpose: Main orchestrator\n")
        f.write("Usage: python run_power_flow.py -i file.py -v 10 -n 200\n")
        f.write("Workflow: Read → Convert → Initialize → Simulate → Report\n\n")
        
        f.write("DATA FLOW SUMMARY\n")
        f.write("-"*40 + "\n")
        f.write("PYPOWER.py → [File 1] → Custom.py → [File 2] → Data dict\n")
        f.write("                                                     ↓\n")
        f.write("Output files ← [File 4] ← Samples ← [File 3] ← Engine format\n")
        f.write("                    ↑\n")
        f.write("              [File 5 orchestrates all]\n\n")
        
        f.write("KEY FORMULAS\n")
        f.write("-"*40 + "\n")
        f.write("y = 1/(R+jX) = (R-jX)/(R²+X²)\n")
        f.write("S = V × (Ybus × V)*\n")
        f.write("Line: S_fwd = Vi × [(Vi-Vj)×y + Vi×jB/2]*\n")
        f.write("Xfmr: S_fwd = Vi × [(Vi/a - Vj)×y]*\n")
        f.write("Loss = S_fwd + S_rev\n")
        f.write("Loading% = |S_fwd| × baseMVA / rateA × 100\n")
        f.write("Q_cap = Q_rated × V²\n")
        f.write("V_kV = V_pu × base_kV\n")
    
    print("Created: 01_Main_Calculations.txt")
    print("Created: 02_User_Manual.txt")
    print("Created: 03_Simplified_Summary.txt")

if __name__ == "__main__":
    create_all_txt()