"""
ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)
INVERSE POWER FLOW TEMPLATE-THIS IS A EXAMPLE FILE TO SHOW THE FORMAT OF THE INPUT DATA
"""

BASE_MVA = 100.0

# ========== BUS DATA ==========
# Format: [bus_num, bus_name, bus_type, base_kV, V_init_pu, angle_init_deg, shunt_G, shunt_B, area, zone, owner]
BUS_DATA = [
    [0, "Bus_0", 3, 230, 1, 0, 0, 0, 1, 1, 1],
    [1, "Bus_1", 1, 230, 1, 0, 0, 0, 1, 1, 1],
    [2, "Bus_2", 1, 230, 1, 0, 0, 0, 1, 1, 1],
    [3, "Bus_3", 2, 230, 1, 0, 0, 0, 1, 1, 1],
]

# ========== GENERATOR DATA ==========
# Format: [gen_num, gen_name, gen_type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone]
GENERATOR_DATA = [
    [1, "Gen_3_1", "SYNC", 3, 318, 0, 1.0200, -100, 100, 1, 1, 1],
    [2, "Gen_0_2", "SYNC", 0, 0, 0, 1, -100, 100, 1, 1, 1],
]

# ========== LOAD DATA ==========
# Format: [load_num, load_name, bus, P_demand, Q_demand, model, area, zone]
LOAD_DATA = [
    [1, "Load_0", 0, 50, 30.9900, "constant_PQ", 1, 1],
    [2, "Load_1", 1, 170, 105.3500, "constant_PQ", 1, 1],
    [3, "Load_2", 2, 200, 123.9400, "constant_PQ", 1, 1],
    [4, "Load_3", 3, 80, 49.5800, "constant_PQ", 1, 1],
]

# ========== TRANSMISSION LINE DATA ==========
# Format: [line_num, line_name, from_bus, to_bus, I_length_km, I_R_per_km, I_X_per_km, I_B_per_km, I_rateA_MVA, I_status, area, zone, owner]
LINE_DATA = [
    [1, "Line_0_1", 0, 1, 1, 0.0101, 0.0504, 0.1025, 250, 1, 1, 1, 1],
    [2, "Line_0_2", 0, 2, 1, 0.0074, 0.0372, 0.0775, 250, 1, 1, 1, 1],
    [3, "Line_1_3", 1, 3, 1, 0.0074, 0.0372, 0.0775, 250, 1, 1, 1, 1],
    [4, "Line_2_3", 2, 3, 1, 0.0127, 0.0636, 0.1275, 250, 1, 1, 1, 1],
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
