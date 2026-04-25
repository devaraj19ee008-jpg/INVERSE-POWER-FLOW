"""
ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)
INVERSE POWER FLOW TEMPLATE-THIS IS A EXAMPLE FILE TO SHOW THE FORMAT OF THE INPUT DATA
"""

BASE_MVA = 100.0

# ========== BUS DATA ==========
# Format: [bus_num, bus_name, bus_type, base_kV, V_init_pu, angle_init_deg, shunt_G, shunt_B, area, zone, owner]
BUS_DATA = [
    [1, "Bus_1", 3, 0, 1.0600, 0, 0, 0, 1, 1, 1],
    [2, "Bus_2", 2, 0, 1.0450, -4.9800, 0, 0, 1, 1, 1],
    [3, "Bus_3", 2, 0, 1.0100, -12.7200, 0, 0, 1, 1, 1],
    [4, "Bus_4", 1, 0, 1.0190, -10.3300, 0, 0, 1, 1, 1],
    [5, "Bus_5", 1, 0, 1.0200, -8.7800, 0, 0, 1, 1, 1],
    [6, "Bus_6", 2, 0, 1.0700, -14.2200, 0, 0, 1, 1, 1],
    [7, "Bus_7", 1, 0, 1.0620, -13.3700, 0, 0, 1, 1, 1],
    [8, "Bus_8", 2, 0, 1.0900, -13.3600, 0, 0, 1, 1, 1],
    [9, "Bus_9", 1, 0, 1.0560, -14.9400, 0, 19, 1, 1, 1],
    [10, "Bus_10", 1, 0, 1.0510, -15.1000, 0, 0, 1, 1, 1],
    [11, "Bus_11", 1, 0, 1.0570, -14.7900, 0, 0, 1, 1, 1],
    [12, "Bus_12", 1, 0, 1.0550, -15.0700, 0, 0, 1, 1, 1],
    [13, "Bus_13", 1, 0, 1.0500, -15.1600, 0, 0, 1, 1, 1],
    [14, "Bus_14", 1, 0, 1.0360, -16.0400, 0, 0, 1, 1, 1],
]

# ========== GENERATOR DATA ==========
# Format: [gen_num, gen_name, gen_type, bus, P_out, Q_out, V_set, Qmin, Qmax, status, area, zone]
GENERATOR_DATA = [
    [1, "Gen_1_1", "SYNC", 1, 232.4000, -16.9000, 1.0600, 0, 10, 1, 1, 1],
    [2, "Gen_2_2", "SYNC", 2, 40, 42.4000, 1.0450, -40, 50, 1, 1, 1],
    [3, "Gen_3_3", "SYNC", 3, 0, 23.4000, 1.0100, 0, 40, 1, 1, 1],
    [4, "Gen_6_4", "SYNC", 6, 0, 12.2000, 1.0700, -6, 24, 1, 1, 1],
    [5, "Gen_8_5", "SYNC", 8, 0, 17.4000, 1.0900, -6, 24, 1, 1, 1],
]

# ========== LOAD DATA ==========
# Format: [load_num, load_name, bus, P_demand, Q_demand, model, area, zone]
LOAD_DATA = [
    [1, "Load_2", 2, 21.7000, 12.7000, "constant_PQ", 1, 1],
    [2, "Load_3", 3, 94.2000, 19, "constant_PQ", 1, 1],
    [3, "Load_4", 4, 47.8000, -3.9000, "constant_PQ", 1, 1],
    [4, "Load_5", 5, 7.6000, 1.6000, "constant_PQ", 1, 1],
    [5, "Load_6", 6, 11.2000, 7.5000, "constant_PQ", 1, 1],
    [6, "Load_9", 9, 29.5000, 16.6000, "constant_PQ", 1, 1],
    [7, "Load_10", 10, 9, 5.8000, "constant_PQ", 1, 1],
    [8, "Load_11", 11, 3.5000, 1.8000, "constant_PQ", 1, 1],
    [9, "Load_12", 12, 6.1000, 1.6000, "constant_PQ", 1, 1],
    [10, "Load_13", 13, 13.5000, 5.8000, "constant_PQ", 1, 1],
    [11, "Load_14", 14, 14.9000, 5, "constant_PQ", 1, 1],
]

# ========== TRANSMISSION LINE DATA ==========
# Format: [line_num, line_name, from_bus, to_bus, I_length_km, I_R_per_km, I_X_per_km, I_B_per_km, I_rateA_MVA, I_status, area, zone, owner]
LINE_DATA = [
    [1, "Line_1_2", 1, 2, 1, 0.0194, 0.0592, 0.0528, 9900, 1, 1, 1, 1],
    [2, "Line_1_5", 1, 5, 1, 0.0540, 0.2230, 0.0492, 9900, 1, 1, 1, 1],
    [3, "Line_2_3", 2, 3, 1, 0.0470, 0.1980, 0.0438, 9900, 1, 1, 1, 1],
    [4, "Line_2_4", 2, 4, 1, 0.0581, 0.1763, 0.0340, 9900, 1, 1, 1, 1],
    [5, "Line_2_5", 2, 5, 1, 0.0570, 0.1739, 0.0346, 9900, 1, 1, 1, 1],
    [6, "Line_3_4", 3, 4, 1, 0.0670, 0.1710, 0.0128, 9900, 1, 1, 1, 1],
    [7, "Line_4_5", 4, 5, 1, 0.0134, 0.0421, 0, 9900, 1, 1, 1, 1],
    [8, "Line_6_11", 6, 11, 1, 0.0950, 0.1989, 0, 9900, 1, 1, 1, 1],
    [9, "Line_6_12", 6, 12, 1, 0.1229, 0.2558, 0, 9900, 1, 1, 1, 1],
    [10, "Line_6_13", 6, 13, 1, 0.0662, 0.1303, 0, 9900, 1, 1, 1, 1],
    [11, "Line_7_8", 7, 8, 1, 0, 0.1762, 0, 9900, 1, 1, 1, 1],
    [12, "Line_7_9", 7, 9, 1, 0, 0.1100, 0, 9900, 1, 1, 1, 1],
    [13, "Line_9_10", 9, 10, 1, 0.0318, 0.0845, 0, 9900, 1, 1, 1, 1],
    [14, "Line_9_14", 9, 14, 1, 0.1271, 0.2704, 0, 9900, 1, 1, 1, 1],
    [15, "Line_10_11", 10, 11, 1, 0.0820, 0.1921, 0, 9900, 1, 1, 1, 1],
    [16, "Line_12_13", 12, 13, 1, 0.2209, 0.1999, 0, 9900, 1, 1, 1, 1],
    [17, "Line_13_14", 13, 14, 1, 0.1709, 0.3480, 0, 9900, 1, 1, 1, 1],
]

# ========== TRANSFORMER DATA ==========
# Format: [xfmr_num, xfmr_name, from_bus, to_bus, I_R_pu, I_X_pu, I_tap_ratio, I_rateA_MVA, I_phase_shift_deg, I_min_tap, I_max_tap, I_step_size, I_status, area, zone, owner]
TRANSFORMER_DATA = [
    [1, "Xfmr_4_7", 4, 7, 0, 0.2091, 0.9780, 100, 0, 0.9000, 1.1000, 0.0100, 1, 1, 1, 1],
    [2, "Xfmr_4_9", 4, 9, 0, 0.5562, 0.9690, 100, 0, 0.9000, 1.1000, 0.0100, 1, 1, 1, 1],
    [3, "Xfmr_5_6", 5, 6, 0, 0.2520, 0.9320, 100, 0, 0.9000, 1.1000, 0.0100, 1, 1, 1, 1],
]

# ========== CAPACITOR DATA ==========
# Format: [cap_num, name, bus, Q_cap, status, area, zone]
CAPACITOR_DATA = []

# ========== REACTOR DATA ==========
# Format: [reactor_num, name, bus, Q_react, status, area, zone]
REACTOR_DATA = [
    [1, "Reactor_9", 9, 1900, 1, 1, 1],
]

# ========== SERIES COMPENSATION DATA ==========
# Format: [series_num, name, from_bus, to_bus, r, x, comp_pct, status, area, zone, owner]
SERIES_COMP_DATA = []

# ========== SERIES REACTOR DATA ==========
# Format: [series_reactor_num, name, from_bus, to_bus, r, x, status, area, zone, owner]
SERIES_REACTOR_DATA = []

# ========== SHUNT COMPENSATION DATA ==========
# Format: [shunt_num, name, bus, Q_shunt, status, area, zone]
SHUNT_DATA = []
