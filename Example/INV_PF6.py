"""
ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)
INVERSE POWER FLOW TEMPLATE-THIS IS A EXAMPLE FILE TO SHOW THE FORMAT OF THE INPUT DATA
"""

BASE_MVA = 100.0

# ========== BUS DATA ==========
# Format: [bus_num, bus_name, bus_type, base_kV, V_init_pu, angle_init_deg, shunt_G, shunt_B, area, zone, owner]
BUS_DATA = [
    [1, "Bus_1", 3, 230, 1.0500, 0, 0, 0, 1, 1, 1],
    [2, "Bus_2", 2, 230, 1.0500, 0, 0, 0, 1, 1, 1],
    [3, "Bus_3", 2, 230, 1.0700, 0, 0, 0, 1, 1, 1],
    [4, "Bus_4", 1, 230, 1, 0, 0, 0, 1, 1, 1],
    [5, "Bus_5", 1, 230, 1, 0, 0, 0, 1, 1, 1],
    [6, "Bus_6", 1, 230, 1, 0, 0, 0, 1, 1, 1],
]

# ========== GENERATOR DATA ==========
# Format: [gen_num, gen_name, gen_type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone]
GENERATOR_DATA = [
    [1, "Gen_1_1", "SYNC", 1, 0, 0, 1.0500, -100, 100, 1, 1, 1],
    [2, "Gen_2_2", "SYNC", 2, 50, 0, 1.0500, -100, 100, 1, 1, 1],
    [3, "Gen_3_3", "SYNC", 3, 60, 0, 1.0700, -100, 100, 1, 1, 1],
]

# ========== LOAD DATA ==========
# Format: [load_num, load_name, bus, P_demand, Q_demand, model, area, zone]
LOAD_DATA = [
    [1, "Load_4", 4, 70, 70, "constant_PQ", 1, 1],
    [2, "Load_5", 5, 70, 70, "constant_PQ", 1, 1],
    [3, "Load_6", 6, 70, 70, "constant_PQ", 1, 1],
]

# ========== TRANSMISSION LINE DATA ==========
# Format: [line_num, line_name, from_bus, to_bus, I_length_km, I_R_per_km, I_X_per_km, I_B_per_km, I_rateA_MVA, I_status, area, zone, owner]
LINE_DATA = [
    [1, "Line_1_2", 1, 2, 1, 0.1000, 0.2000, 0.0400, 40, 1, 1, 1, 1],
    [2, "Line_1_4", 1, 4, 1, 0.0500, 0.2000, 0.0400, 60, 1, 1, 1, 1],
    [3, "Line_1_5", 1, 5, 1, 0.0800, 0.3000, 0.0600, 40, 1, 1, 1, 1],
    [4, "Line_2_3", 2, 3, 1, 0.0500, 0.2500, 0.0600, 40, 1, 1, 1, 1],
    [5, "Line_2_4", 2, 4, 1, 0.0500, 0.1000, 0.0200, 60, 1, 1, 1, 1],
    [6, "Line_2_5", 2, 5, 1, 0.1000, 0.3000, 0.0400, 30, 1, 1, 1, 1],
    [7, "Line_2_6", 2, 6, 1, 0.0700, 0.2000, 0.0500, 90, 1, 1, 1, 1],
    [8, "Line_3_5", 3, 5, 1, 0.1200, 0.2600, 0.0500, 70, 1, 1, 1, 1],
    [9, "Line_3_6", 3, 6, 1, 0.0200, 0.1000, 0.0200, 80, 1, 1, 1, 1],
    [10, "Line_4_5", 4, 5, 1, 0.2000, 0.4000, 0.0800, 20, 1, 1, 1, 1],
    [11, "Line_5_6", 5, 6, 1, 0.1000, 0.3000, 0.0600, 40, 1, 1, 1, 1],
]

# ========== TRANSFORMER DATA ==========
# Format: [xfmr_num, xfmr_name, from_bus, to_bus, I_R_pu, I_X_pu, I_tap_ratio, I_rateA_MVA, I_phase_shift_deg, I_min_tap, I_max_tap, I_step_size, I_status, area, zone, owner]
TRANSFORMER_DATA = []

# ========== CAPACITOR DATA ==========
# Format: [cap_num, name, bus, Q_cap, status, area, zone]
CAPACITOR_DATA = []

# ========== REACTOR DATA ==========
# Format: [reactor_num, name, bus, Q_react, status, area, zone]
REACTOR_DATA = []

# ========== SERIES COMPENSATION DATA ==========
# Format: [series_num, name, from_bus, to_bus, r, x, comp_pct, status, area, zone, owner]
SERIES_COMP_DATA = []

# ========== SERIES REACTOR DATA ==========
# Format: [series_reactor_num, name, from_bus, to_bus, r, x, status, area, zone, owner]
SERIES_REACTOR_DATA = []

# ========== SHUNT COMPENSATION DATA ==========
# Format: [shunt_num, name, bus, Q_shunt, status, area, zone]
SHUNT_DATA = []
