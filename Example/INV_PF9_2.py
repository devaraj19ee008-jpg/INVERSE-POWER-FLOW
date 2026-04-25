"""
ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)
INVERSE POWER FLOW TEMPLATE-THIS IS A EXAMPLE FILE TO SHOW THE FORMAT OF THE INPUT DATA
"""

BASE_MVA = 100.0

# ========== BUS DATA ==========
# Format: [bus_num, bus_name, bus_type, base_kV, V_init_pu, angle_init_deg, shunt_G, shunt_B, area, zone, owner]
BUS_DATA = [
    [1, "Bus_1", 3, 345, 1, 0, 0, 0, 1, 1, 1],
    [2, "Bus_2", 2, 345, 1, 0, 0, 0, 1, 1, 1],
    [3, "Bus_3", 2, 345, 1, 0, 0, 0, 1, 1, 1],
    [4, "Bus_4", 1, 345, 1, 0, 0, 0, 1, 1, 1],
    [5, "Bus_5", 1, 345, 1, 0, 0, 0, 1, 1, 1],
    [6, "Bus_6", 1, 345, 1, 0, 0, 0, 1, 1, 1],
    [7, "Bus_7", 1, 345, 1, 0, 0, 0, 1, 1, 1],
    [8, "Bus_8", 1, 345, 1, 0, 0, 0, 1, 1, 1],
    [9, "Bus_9", 1, 345, 1, 0, 0, 0, 1, 1, 1],
]

# ========== GENERATOR DATA ==========
# Format: [gen_num, gen_name, gen_type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone]
GENERATOR_DATA = [
    [1, "Gen_1_1", "SYNC", 1, 0, 0, 1, -300, 300, 1, 1, 1],
    [2, "Gen_2_2", "SYNC", 2, 248, 0, 1, -300, 300, 1, 1, 1],
    [3, "Gen_3_3", "SYNC", 3, 124.5900, 0, 1, -300, 300, 1, 1, 1],
]

# ========== LOAD DATA ==========
# Format: [load_num, load_name, bus, P_demand, Q_demand, model, area, zone]
LOAD_DATA = [
    [1, "Load_5", 5, 305.4000, 123.1500, "constant_PQ", 1, 1],
    [2, "Load_7", 7, 214.1100, 71.0900, "constant_PQ", 1, 1],
    [3, "Load_9", 9, 235.6100, 81.4400, "constant_PQ", 1, 1],
]

# ========== TRANSMISSION LINE DATA ==========
# Format: [line_num, line_name, from_bus, to_bus, I_length_km, I_R_per_km, I_X_per_km, I_B_per_km, I_rateA_MVA, I_status, area, zone, owner]
LINE_DATA = [
    [1, "Line_1_4", 1, 4, 1, 0, 0.0576, 0, 250, 1, 1, 1, 1],
    [2, "Line_4_5", 4, 5, 1, 0.0170, 0.0920, 0.1580, 250, 1, 1, 1, 1],
    [3, "Line_5_6", 5, 6, 1, 0.0390, 0.1700, 0.3580, 150, 1, 1, 1, 1],
    [4, "Line_3_6", 3, 6, 1, 0, 0.0586, 0, 300, 1, 1, 1, 1],
    [5, "Line_6_7", 6, 7, 1, 0.0119, 0.1008, 0.2090, 150, 1, 1, 1, 1],
    [6, "Line_7_8", 7, 8, 1, 0.0085, 0.0720, 0.1490, 250, 1, 1, 1, 1],
    [7, "Line_8_2", 8, 2, 1, 0, 0.0625, 0, 250, 1, 1, 1, 1],
    [8, "Line_8_9", 8, 9, 1, 0.0320, 0.1610, 0.3060, 250, 1, 1, 1, 1],
    [9, "Line_9_4", 9, 4, 1, 0.0100, 0.0850, 0.1760, 250, 1, 1, 1, 1],
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
