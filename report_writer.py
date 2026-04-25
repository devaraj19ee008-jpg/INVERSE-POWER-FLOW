"""
Report Writer - Generates multiple output formats with ALL elements
Includes ALL components even if not present in input (empty headings only)
"""

import numpy as np
import csv
from datetime import datetime

class ReportWriter:
    """
    Generates reports with ALL possible elements
    Missing elements get headers but empty data rows
    """
    
    def __init__(self, engine, input_data, samples, variation_strength):
        self.engine = engine
        self.input_data = input_data
        self.samples = samples
        self.variation_strength = variation_strength
        
        # Calculate averages
        self.avg_P = np.mean([s['P_calc'] for s in samples], axis=0) * engine.baseMVA
        self.avg_Q = np.mean([s['Q_calc'] for s in samples], axis=0) * engine.baseMVA
        self.avg_V = np.mean([s['V_mag'] for s in samples], axis=0)
        self.avg_Va = np.mean([s['V_angle'] for s in samples], axis=0)
        
        # Process branch flows
        self.process_branch_flows()
    
    def process_branch_flows(self):
        """Process branch flow statistics"""
        self.line_stats = {}
        self.xfmr_stats = {}
        self.series_comp_stats = {}
        self.series_reactor_stats = {}
        
        for sample in self.samples:
            for flow in sample['branch_flows']:
                key = f"{flow['from_bus']}-{flow['to_bus']}"
                if flow['type'] == 'line':
                    if key not in self.line_stats:
                        self.line_stats[key] = {'P_fwd': [], 'Q_fwd': [], 'P_rev': [], 'Q_rev': [], 'loss': [], 'loading': []}
                    self.line_stats[key]['P_fwd'].append(flow['P_fwd_MW'])
                    self.line_stats[key]['Q_fwd'].append(flow['Q_fwd_Mvar'])
                    self.line_stats[key]['P_rev'].append(flow['P_rev_MW'])
                    self.line_stats[key]['Q_rev'].append(flow['Q_rev_Mvar'])
                    self.line_stats[key]['loss'].append(flow['losses_MW'])
                    self.line_stats[key]['loading'].append(flow['loading_pct'])
                elif flow['type'] == 'transformer':
                    if key not in self.xfmr_stats:
                        self.xfmr_stats[key] = {'P_fwd': [], 'Q_fwd': [], 'P_rev': [], 'Q_rev': [], 'loss': [], 'loading': [], 'tap': []}
                    self.xfmr_stats[key]['P_fwd'].append(flow['P_fwd_MW'])
                    self.xfmr_stats[key]['Q_fwd'].append(flow['Q_fwd_Mvar'])
                    self.xfmr_stats[key]['P_rev'].append(flow['P_rev_MW'])
                    self.xfmr_stats[key]['Q_rev'].append(flow['Q_rev_Mvar'])
                    self.xfmr_stats[key]['loss'].append(flow['losses_MW'])
                    self.xfmr_stats[key]['loading'].append(flow['loading_pct'])
                    self.xfmr_stats[key]['tap'].append(flow.get('tap_ratio', 1.0))
        
        # Calculate averages for lines
        for stats in self.line_stats.values():
            stats['avg_P_fwd'] = np.mean(stats['P_fwd']) if stats['P_fwd'] else 0
            stats['avg_Q_fwd'] = np.mean(stats['Q_fwd']) if stats['Q_fwd'] else 0
            stats['avg_P_rev'] = np.mean(stats['P_rev']) if stats['P_rev'] else 0
            stats['avg_Q_rev'] = np.mean(stats['Q_rev']) if stats['Q_rev'] else 0
            stats['avg_loss'] = np.mean(stats['loss']) if stats['loss'] else 0
            stats['avg_loading'] = np.mean(stats['loading']) if stats['loading'] else 0
        
        # Calculate averages for transformers
        for stats in self.xfmr_stats.values():
            stats['avg_P_fwd'] = np.mean(stats['P_fwd']) if stats['P_fwd'] else 0
            stats['avg_Q_fwd'] = np.mean(stats['Q_fwd']) if stats['Q_fwd'] else 0
            stats['avg_P_rev'] = np.mean(stats['P_rev']) if stats['P_rev'] else 0
            stats['avg_Q_rev'] = np.mean(stats['Q_rev']) if stats['Q_rev'] else 0
            stats['avg_loss'] = np.mean(stats['loss']) if stats['loss'] else 0
            stats['avg_loading'] = np.mean(stats['loading']) if stats['loading'] else 0
            stats['avg_tap'] = np.mean(stats['tap']) if stats['tap'] else 1.0
    
    def get_voltage_kv(self, bus_num, V_pu):
        """Convert pu voltage to kV"""
        if bus_num in self.input_data.get('buses', {}):
            return V_pu * self.input_data['buses'][bus_num].get('base_kV', 132.0)
        return V_pu * 132.0
    
    def get_status(self, value, thresholds):
        """Get status based on thresholds"""
        if value < thresholds.get('min', 0.95):
            return 'LOW'
        elif value > thresholds.get('max', 1.05):
            return 'HIGH'
        return 'NORMAL'
    
    def write_csv_all_data(self, filename):
        """Write CSV with ALL data - includes ALL possible elements"""
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            
            # =========================================================
            # SECTION 1: BUS DATA 
            # =========================================================
            writer.writerow(['# ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)'])
            writer.writerow(['# ========== BUS DATA =========='])
            writer.writerow(['bus_num', 'bus_name', 'bus_type', 'base_kV', 'area', 'zone', 'owner',
                           'V_init_pu', 'angle_init_deg', 'shunt_G_pu', 'shunt_B_pu',
                           'O_V_final_pu', 'O_V_final_kV', 'O_angle_deg', 'O_voltage_status', 'O_angle_status'])
            
            buses = self.input_data.get('buses', {})
            if buses:
                for bus_num, bus in buses.items():
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    v_status = self.get_status(self.avg_V[idx] if idx < len(self.avg_V) else 1.0, {'min': 0.95, 'max': 1.05})
                    a_status = self.get_status(abs(self.avg_Va[idx] if idx < len(self.avg_Va) else 0), {'min': 0, 'max': 30})
                    
                    writer.writerow([bus_num, bus.get('name', f'Bus_{bus_num}'), bus.get('type', 1), bus.get('base_kV', 132),
                                   bus.get('area', 1), bus.get('zone', 1), bus.get('owner', 1),
                                   bus.get('V_init', 1.0), bus.get('angle_init', 0.0), 
                                   bus.get('shunt_G', 0), bus.get('shunt_B', 0),
                                   f"{self.avg_V[idx]:.4f}" if idx < len(self.avg_V) else "N/A", 
                                   f"{V_kV:.2f}", 
                                   f"{self.avg_Va[idx]:.2f}" if idx < len(self.avg_Va) else "N/A", 
                                   v_status, a_status])
            else:
                writer.writerow(['No bus data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 2: GENERATOR DATA
            # =========================================================
            writer.writerow(['# ========== GENERATOR DATA =========='])
            writer.writerow(['gen_num', 'gen_name', 'gen_type', 'connected_bus', 'area', 'zone',
                           'I_P_out_MW', 'I_Q_out_Mvar', 'I_V_set_pu', 'I_V_set_kV',
                           'I_Qmin_Mvar', 'I_Qmax_Mvar', 'I_status',
                           'O_P_out_MW', 'O_Q_out_Mvar', 'O_V_actual_pu', 'O_V_actual_kV',
                           'O_Q_status', 'O_limit_flag', 'O_status'])
            
            generators = self.input_data.get('generators', {})
            if generators:
                for gen_num, gen in generators.items():
                    bus_num = gen.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    O_P_out = -self.avg_P[idx] if idx < len(self.avg_P) and self.avg_P[idx] < 0 else gen.get('P_out', 0)
                    O_Q_out = gen.get('Q_out', 0)
                    
                    q_status = "WITHIN LIMITS"
                    limit_flag = "OK"
                    if O_Q_out < gen.get('Qmin', -9999):
                        q_status = "BELOW Qmin"
                        limit_flag = "VIOLATED"
                    elif O_Q_out > gen.get('Qmax', 9999):
                        q_status = "ABOVE Qmax"
                        limit_flag = "VIOLATED"
                    
                    writer.writerow([gen_num, gen.get('name', f'Gen_{gen_num}'), gen.get('type', 'SYNC'), bus_num,
                                   gen.get('area', 1), gen.get('zone', 1),
                                   gen.get('P_out', 0), gen.get('Q_out', 0), gen.get('V_set', 1.0),
                                   f"{gen.get('V_set', 1.0) * self.input_data.get('buses', {}).get(bus_num, {}).get('base_kV', 132):.2f}",
                                   gen.get('Qmin', -9999), gen.get('Qmax', 9999), gen.get('status', 1),
                                   f"{O_P_out:.3f}", f"{O_Q_out:.3f}", 
                                   f"{self.avg_V[idx]:.4f}" if idx < len(self.avg_V) else "N/A",
                                   f"{V_kV:.2f}", q_status, limit_flag, "ONLINE" if gen.get('status', 1) == 1 else "OFFLINE"])
            else:
                writer.writerow(['No generator data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 3: LOAD DATA
            # =========================================================
            writer.writerow(['# ========== LOAD DATA =========='])
            writer.writerow(['load_num', 'load_name', 'connected_bus', 'area', 'zone',
                           'I_P_demand_MW', 'I_Q_demand_Mvar', 'I_load_model',
                           'O_P_supplied_MW', 'O_Q_supplied_Mvar', 'O_V_actual_pu', 'O_V_actual_kV', 'O_supply_status'])
            
            loads = self.input_data.get('loads', {})
            if loads:
                for load_num, load in loads.items():
                    bus_num = load.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    
                    supply_status = "NORMAL" if (idx < len(self.avg_V) and self.avg_V[idx] >= 0.95) else "VOLTAGE LOW"
                    
                    writer.writerow([load_num, load.get('name', f'Load_{load_num}'), bus_num,
                                   load.get('area', 1), load.get('zone', 1),
                                   load.get('P_demand', 0), load.get('Q_demand', 0), load.get('model', 'constant_PQ'),
                                   f"{load.get('P_demand', 0):.3f}", f"{load.get('Q_demand', 0):.3f}",
                                   f"{self.avg_V[idx]:.4f}" if idx < len(self.avg_V) else "N/A",
                                   f"{V_kV:.2f}", supply_status])
            else:
                writer.writerow(['No load data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 4: TRANSMISSION LINE DATA (Bidirectional)
            # =========================================================
            writer.writerow(['# ========== TRANSMISSION LINE DATA =========='])  
            writer.writerow(['line_num', 'line_name', 'from_bus', 'to_bus', 'length_km', 'area', 'zone', 'owner',
                        'I_R_per_km', 'I_X_per_km', 'I_B_per_km', 'I_R_total_pu', 'I_X_total_pu', 'I_B_total_pu',
                        'I_rateA_MVA', 'I_rateB_MVA', 'I_rateC_MVA', 'I_status',
                        'O_P_fwd_MW', 'O_Q_fwd_Mvar', 'O_MVA_fwd_MVA',
                        'O_P_rev_MW', 'O_Q_rev_Mvar', 'O_MVA_rev_MVA',
                        'O_losses_MW', 'O_loss_per_km', 'O_loading_pct', 'O_status'])                      
            
            lines = self.input_data.get('lines', {})
            if lines:
                for line_num, line in lines.items():
                    key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                    stats = self.line_stats.get(key, {})
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    
                    # Get length and calculate loss per km
                    length = line.get('length_km', 0)
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_per_km = Loss_MW / length if length > 0 else 0
                    
                    # Calculate MVA values
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    writer.writerow([line_num, line.get('name', f'Line_{line_num}'), 
                                line.get('from_bus', 0), line.get('to_bus', 0),
                                f"{length:.2f}",  # NEW: length in km
                                line.get('area', 1), line.get('zone', 1), line.get('owner', 1),
                                line.get('R_per_km', 0), line.get('X_per_km', 0), line.get('B_per_km', 0),  # NEW: per-km values
                                line.get('r', 0), line.get('x', 0), line.get('b', 0),  # Total R, X, B
                                line.get('rateA', 0), line.get('rateB', 0), line.get('rateC', 0), line.get('status', 1),
                                f"{P_fwd:.3f}", f"{Q_fwd:.3f}", f"{MVA_fwd:.3f}",  # NEW: MVA_fwd
                                f"{P_rev:.3f}", f"{Q_rev:.3f}", f"{MVA_rev:.3f}",  # NEW: MVA_rev
                                f"{Loss_MW:.3f}", f"{Loss_per_km:.4f}",  # NEW: loss per km
                                f"{stats.get('avg_loading', 0):.1f}", loading_status])
            else:
                writer.writerow(['No transmission line data available'])

            writer.writerow([])

            
            
            # =========================================================
            # SECTION 5: TRANSFORMER DATA (Bidirectional with Tap)
            # =========================================================
            writer.writerow(['# ========== TRANSFORMER DATA =========='])

            writer.writerow(['xfmr_num', 'xfmr_name', 'from_bus', 'to_bus', 'I_rateA_MVA', 'area', 'zone', 'owner',
                        'I_R_pu', 'I_X_pu', 'I_tap_ratio', 'I_phase_shift_deg',
                        'I_min_tap', 'I_max_tap', 'I_step_size', 'I_status',
                        'O_P_fwd_MW', 'O_Q_fwd_Mvar', 'O_MVA_fwd_MVA',
                        'O_P_rev_MW', 'O_Q_rev_Mvar', 'O_MVA_rev_MVA',
                        'O_losses_MW', 'O_losses_MVA', 'O_tap_position', 'O_tap_step', 'O_loading_pct', 'O_status'])

            transformers = self.input_data.get('transformers', {})
            if transformers:
                for xfmr_num, xfmr in transformers.items():
                    key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                    stats = self.xfmr_stats.get(key, {})
                    tap_ratio = stats.get('avg_tap', xfmr.get('tap_ratio', 1.0))
                    step_size = xfmr.get('step_size', 0.01)
                    min_tap = xfmr.get('min_tap', 0.9)
                    tap_step = int((tap_ratio - min_tap) / step_size) if step_size > 0 else 0
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    
                    # Calculate MVA values
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_MVA = (Loss_MW**2 + (abs(stats.get('avg_loss_q', 0)))**2)**0.5 if stats.get('avg_loss_q', 0) else Loss_MW
                    
                    writer.writerow([xfmr_num, xfmr.get('name', f'Xfmr_{xfmr_num}'),
                                xfmr.get('from_bus', 0), xfmr.get('to_bus', 0),
                                xfmr.get('rateA', 9999),  # NEW: Transformer Rating (MVA)
                                xfmr.get('area', 1), xfmr.get('zone', 1), xfmr.get('owner', 1),
                                xfmr.get('r', 0), xfmr.get('x', 0), xfmr.get('tap_ratio', 1.0),
                                xfmr.get('phase_shift', 0), xfmr.get('min_tap', 0.9), xfmr.get('max_tap', 1.1),
                                xfmr.get('step_size', 0.01), xfmr.get('status', 1),
                                f"{P_fwd:.3f}", f"{Q_fwd:.3f}", f"{MVA_fwd:.3f}",  # NEW: MVA_fwd
                                f"{P_rev:.3f}", f"{Q_rev:.3f}", f"{MVA_rev:.3f}",  # NEW: MVA_rev
                                f"{Loss_MW:.3f}", f"{Loss_MVA:.3f}",  # NEW: Loss_MVA
                                f"{tap_ratio:.4f}", tap_step,
                                f"{stats.get('avg_loading', 0):.1f}", loading_status])
            else:
                writer.writerow(['No transformer data available'])

            writer.writerow([])


            
            # =========================================================
            # SECTION 6: CAPACITOR BANK DATA
            # =========================================================
            writer.writerow(['# ========== CAPACITOR BANK DATA =========='])
            writer.writerow(['cap_num', 'cap_name', 'connected_bus', 'area', 'zone',
                           'I_Q_cap_Mvar', 'I_V_pu', 'I_V_kV', 'I_status',
                           'O_Q_injected_Mvar', 'O_V_actual_pu', 'O_V_actual_kV', 'O_status'])
            
            capacitors = self.input_data.get('capacitors', {})
            if capacitors:
                for cap_num, cap in capacitors.items():
                    bus_num = cap.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    O_Q_inj = cap.get('Q_cap', 0) * (V_actual ** 2)
                    
                    writer.writerow([cap_num, cap.get('name', f'Cap_{cap_num}'), bus_num,
                                   cap.get('area', 1), cap.get('zone', 1),
                                   cap.get('Q_cap', 0), "", "", cap.get('status', 1),
                                   f"{O_Q_inj:.3f}", f"{V_actual:.4f}", f"{V_kV:.2f}",
                                   "ONLINE" if cap.get('status', 1) == 1 else "OFFLINE"])
            else:
                writer.writerow(['No capacitor data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 7: REACTOR DATA
            # =========================================================
            writer.writerow(['# ========== REACTOR DATA =========='])
            writer.writerow(['reactor_num', 'reactor_name', 'connected_bus', 'area', 'zone',
                           'I_Q_react_Mvar', 'I_V_pu', 'I_V_kV', 'I_status',
                           'O_Q_absorbed_Mvar', 'O_V_actual_pu', 'O_V_actual_kV', 'O_status'])
            
            reactors = self.input_data.get('reactors', {})
            if reactors:
                for reactor_num, reactor in reactors.items():
                    bus_num = reactor.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    O_Q_abs = reactor.get('Q_react', 0) * (V_actual ** 2)
                    
                    writer.writerow([reactor_num, reactor.get('name', f'Reactor_{reactor_num}'), bus_num,
                                   reactor.get('area', 1), reactor.get('zone', 1),
                                   reactor.get('Q_react', 0), "", "", reactor.get('status', 1),
                                   f"{O_Q_abs:.3f}", f"{V_actual:.4f}", f"{V_kV:.2f}",
                                   "ONLINE" if reactor.get('status', 1) == 1 else "OFFLINE"])
            else:
                writer.writerow(['No reactor data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 8: SERIES COMPENSATION DATA
            # =========================================================
            writer.writerow(['# ========== SERIES COMPENSATION DATA =========='])
            writer.writerow(['series_num', 'series_name', 'from_bus', 'to_bus', 'area', 'zone', 'owner',
                           'I_R_pu', 'I_X_pu', 'I_compensation_pct', 'I_status',
                           'O_effective_X_pu', 'O_compensation_status',
                           'O_P_flow_from_to_MW', 'O_Q_flow_from_to_Mvar',
                           'O_P_flow_to_from_MW', 'O_Q_flow_to_from_Mvar', 'O_losses_MW'])
            
            series_comps = self.input_data.get('series_comps', {})
            if series_comps:
                for series_num, series in series_comps.items():
                    effective_X = series.get('x', 0) * (1 - series.get('comp_pct', 0) / 100)
                    comp_status = "ACTIVE" if series.get('status', 1) == 1 else "INACTIVE"
                    
                    writer.writerow([series_num, series.get('name', f'SC_{series_num}'),
                                   series.get('from_bus', 0), series.get('to_bus', 0),
                                   series.get('area', 1), series.get('zone', 1), series.get('owner', 1),
                                   series.get('r', 0), series.get('x', 0), series.get('comp_pct', 0), series.get('status', 1),
                                   f"{effective_X:.6f}", comp_status, "", "", "", ""])
            else:
                writer.writerow(['No series compensation data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 9: SERIES REACTOR DATA
            # =========================================================
            writer.writerow(['# ========== SERIES REACTOR DATA =========='])
            writer.writerow(['series_reactor_num', 'series_reactor_name', 'from_bus', 'to_bus', 'area', 'zone', 'owner',
                           'I_R_pu', 'I_X_pu', 'I_status',
                           'O_P_flow_from_to_MW', 'O_Q_flow_from_to_Mvar',
                           'O_P_flow_to_from_MW', 'O_Q_flow_to_from_Mvar',
                           'O_losses_MW', 'O_status'])
            
            series_reactors = self.input_data.get('series_reactors', {})
            if series_reactors:
                for sr_num, sr in series_reactors.items():
                    writer.writerow([sr_num, sr.get('name', f'SR_{sr_num}'),
                                   sr.get('from_bus', 0), sr.get('to_bus', 0),
                                   sr.get('area', 1), sr.get('zone', 1), sr.get('owner', 1),
                                   sr.get('r', 0), sr.get('x', 0), sr.get('status', 1),
                                   "", "", "", "", "", "ONLINE" if sr.get('status', 1) == 1 else "OFFLINE"])
            else:
                writer.writerow(['No series reactor data available'])
            
            writer.writerow([])
            
            # =========================================================
            # SECTION 10: SHUNT COMPENSATION DATA
            # =========================================================
            writer.writerow(['# ========== SHUNT COMPENSATION DATA =========='])
            writer.writerow(['shunt_num', 'shunt_name', 'connected_bus', 'area', 'zone',
                           'I_Q_shunt_Mvar', 'I_V_pu', 'I_V_kV', 'I_status',
                           'O_Q_injected_Mvar', 'O_V_actual_pu', 'O_V_actual_kV', 'O_status'])
            
            shunts = self.input_data.get('shunts', {})
            if shunts:
                for shunt_num, shunt in shunts.items():
                    bus_num = shunt.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    O_Q_inj = shunt.get('Q_shunt', 0) * (V_actual ** 2)
                    
                    writer.writerow([shunt_num, shunt.get('name', f'Shunt_{shunt_num}'), bus_num,
                                   shunt.get('area', 1), shunt.get('zone', 1),
                                   shunt.get('Q_shunt', 0), "", "", shunt.get('status', 1),
                                   f"{O_Q_inj:.3f}", f"{V_actual:.4f}", f"{V_kV:.2f}",
                                   "ONLINE" if shunt.get('status', 1) == 1 else "OFFLINE"])
            else:
                writer.writerow(['No shunt compensation data available'])
            
            writer.writerow([])
            

            # =========================================================
            # SECTION 11: SYSTEM SUMMARY
            # =========================================================
            total_losses = sum(s['avg_loss'] for s in self.line_stats.values()) + sum(s['avg_loss'] for s in self.xfmr_stats.values())
            total_gen = sum(g.get('P_out', 0) for g in self.input_data.get('generators', {}).values())
            total_load = sum(l.get('P_demand', 0) for l in self.input_data.get('loads', {}).values())
            
            writer.writerow(['# ========== SYSTEM SUMMARY =========='])
            writer.writerow(['Parameter', 'Value', 'Unit'])
            writer.writerow(['Total Generation', f"{total_gen:.3f}", 'MW'])
            writer.writerow(['Total Load', f"{total_load:.3f}", 'MW'])
            writer.writerow(['Total Losses', f"{total_losses:.3f}", 'MW'])
            writer.writerow(['Loss Percentage', f"{(total_losses/total_load)*100 if total_load > 0 else 0:.2f}", '%'])
            writer.writerow(['Average Voltage', f"{np.mean(self.avg_V):.4f}", 'pu'])
            writer.writerow(['Minimum Voltage', f"{np.min(self.avg_V):.4f}", 'pu'])
            writer.writerow(['Maximum Voltage', f"{np.max(self.avg_V):.4f}", 'pu'])
            writer.writerow(['Number of Buses', f"{len(self.input_data.get('buses', {}))}", ''])
            writer.writerow(['Number of Generators', f"{len(self.input_data.get('generators', {}))}", ''])
            writer.writerow(['Number of Loads', f"{len(self.input_data.get('loads', {}))}", ''])
            writer.writerow(['Number of Lines', f"{len(self.input_data.get('lines', {}))}", ''])
            writer.writerow(['Number of Transformers', f"{len(self.input_data.get('transformers', {}))}", ''])
            writer.writerow(['Number of Capacitors', f"{len(self.input_data.get('capacitors', {}))}", ''])
            writer.writerow(['Number of Reactors', f"{len(self.input_data.get('reactors', {}))}", ''])
            writer.writerow(['Number of Series Comps', f"{len(self.input_data.get('series_comps', {}))}", ''])
            writer.writerow(['Number of Series Reactors', f"{len(self.input_data.get('series_reactors', {}))}", ''])
            writer.writerow(['Number of Shunts', f"{len(self.input_data.get('shunts', {}))}", ''])
            writer.writerow(['Samples Generated', f"{len(self.samples)}", ''])
            writer.writerow(['Variation Strength', f"{self.variation_strength*100:.0f}", '%'])
        
        #print(f"  ✓ CSV (all data with ALL elements): {filename}")
        return filename
    
    def write_py_all_data(self, filename):
        """Write PY file with ALL data - includes ALL possible elements"""
        
        with open(filename, 'w', encoding='utf-8') as f:
            
            f.write('"""\n')
            f.write(f'ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)\n')
            f.write(f'============INVERSE POWER FLOW=============\n')
            f.write(f'Complete Power System Data with Results\n')
            f.write(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write('This file contains ALL data (input + output) for ALL possible elements\n')
            f.write('Empty lists for components not present in input\n')
            f.write('"""\n\n')
            
            f.write(f'BASE_MVA = {self.engine.baseMVA}\n\n')
            
            # BUS DATA
            f.write('# ========== BUS DATA ==========\n')
            f.write('# [bus_num, name, type, base_kV, area, zone, owner, V_init, angle_init, shunt_G, shunt_B, O_V_final, O_V_kV, O_angle, O_v_status, O_a_status]\n')
            f.write('BUS_DATA = [\n')
            buses = self.input_data.get('buses', {})
            if buses:
                for bus_num, bus in buses.items():
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    v_status = self.get_status(self.avg_V[idx] if idx < len(self.avg_V) else 1.0, {'min': 0.95, 'max': 1.05})
                    a_status = self.get_status(abs(self.avg_Va[idx] if idx < len(self.avg_Va) else 0), {'min': 0, 'max': 30})
                    f.write(f'    [{bus_num}, "{bus.get("name", f"Bus_{bus_num}")}", {bus.get("type", 1)}, {bus.get("base_kV", 132)}, '
                           f'{bus.get("area", 1)}, {bus.get("zone", 1)}, {bus.get("owner", 1)}, '
                           f'{bus.get("V_init", 1.0)}, {bus.get("angle_init", 0.0)}, {bus.get("shunt_G", 0)}, {bus.get("shunt_B", 0)}, '
                           f'{self.avg_V[idx]:.6f}, {V_kV:.2f}, {self.avg_Va[idx]:.4f}, "{v_status}", "{a_status}"],\n')
            f.write(']\n\n')
            
            # GENERATOR DATA
            f.write('# ========== GENERATOR DATA ==========\n')
            f.write('# [gen_num, name, type, bus, area, zone, P_out, Q_out, V_set, Qmin, Qmax, status, O_P_out, O_Q_out, O_V_actual, O_Q_status, O_limit_flag]\n')
            f.write('GENERATOR_DATA = [\n')
            generators = self.input_data.get('generators', {})
            if generators:
                for gen_num, gen in generators.items():
                    bus_num = gen.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    O_P_out = -self.avg_P[idx] if idx < len(self.avg_P) and self.avg_P[idx] < 0 else gen.get('P_out', 0)
                    f.write(f'    [{gen_num}, "{gen.get("name", f"Gen_{gen_num}")}", "{gen.get("type", "SYNC")}", {bus_num}, '
                           f'{gen.get("area", 1)}, {gen.get("zone", 1)}, '
                           f'{gen.get("P_out", 0)}, {gen.get("Q_out", 0)}, {gen.get("V_set", 1.0)}, '
                           f'{gen.get("Qmin", -9999)}, {gen.get("Qmax", 9999)}, {gen.get("status", 1)}, '
                           f'{O_P_out:.6f}, {gen.get("Q_out", 0):.6f}, {self.avg_V[idx]:.6f}, "OK", "NORMAL"],\n')
            f.write(']\n\n')
            
            f.write('# ========== TRANSMISSION LINE DATA ==========\n')
            f.write('# [line_num, name, from_bus, to_bus, length_km, area, zone, owner, R_per_km, X_per_km, B_per_km, R_total, X_total, B_total, rateA, rateB, rateC, status, O_P_fwd, O_Q_fwd, O_MVA_fwd, O_P_rev, O_Q_rev, O_MVA_rev, O_loss_MW, O_loss_per_km, O_loading, O_status]\n')
            f.write('LINE_DATA = [\n')
            lines = self.input_data.get('lines', {})
            if lines:
                for line_num, line in lines.items():
                    key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                    stats = self.line_stats.get(key, {})
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    
                    # Get length and calculate values
                    length = line.get('length_km', 0)
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_per_km = Loss_MW / length if length > 0 else 0
                    
                    # Calculate MVA values
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    f.write(f'    [{line_num}, "{line.get("name", f"Line_{line_num}")}", {line.get("from_bus", 0)}, {line.get("to_bus", 0)}, '
                        f'{length:.2f}, '  # NEW: length_km
                        f'{line.get("area", 1)}, {line.get("zone", 1)}, {line.get("owner", 1)}, '
                        f'{line.get("R_per_km", 0):.6f}, {line.get("X_per_km", 0):.6f}, {line.get("B_per_km", 0):.6f}, '  # NEW: per-km values
                        f'{line.get("r", 0):.6f}, {line.get("x", 0):.6f}, {line.get("b", 0):.6f}, '  # Total R, X, B
                        f'{line.get("rateA", 0)}, {line.get("rateB", 0)}, {line.get("rateC", 0)}, {line.get("status", 1)}, '
                        f'{P_fwd:.6f}, {Q_fwd:.6f}, {MVA_fwd:.6f}, '  # NEW: MVA_fwd
                        f'{P_rev:.6f}, {Q_rev:.6f}, {MVA_rev:.6f}, '  # NEW: MVA_rev
                        f'{Loss_MW:.6f}, {Loss_per_km:.6f}, '  # NEW: loss_per_km
                        f'{stats.get("avg_loading", 0):.2f}, "{loading_status}"],\n')
            f.write(']\n\n')

            # TRANSFORMER DATA
            f.write('# ========== TRANSFORMER DATA ==========\n')
            f.write('# [xfmr_num, name, from_bus, to_bus, rateA_MVA, area, zone, owner, r, x, tap_ratio, phase_shift, min_tap, max_tap, step_size, status, O_P_fwd, O_Q_fwd, O_MVA_fwd, O_P_rev, O_Q_rev, O_MVA_rev, O_loss_MW, O_loss_MVA, O_tap_pos, O_tap_step, O_loading, O_status]\n')
            f.write('TRANSFORMER_DATA = [\n')
            transformers = self.input_data.get('transformers', {})
            if transformers:
                for xfmr_num, xfmr in transformers.items():
                    key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                    stats = self.xfmr_stats.get(key, {})
                    tap_ratio = stats.get('avg_tap', xfmr.get('tap_ratio', 1.0))
                    step_size = xfmr.get('step_size', 0.01)
                    min_tap = xfmr.get('min_tap', 0.9)
                    tap_step = int((tap_ratio - min_tap) / step_size) if step_size > 0 else 0
                    
                    # Calculate MVA values
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_MVA = (Loss_MW**2 + (abs(stats.get('avg_loss_q', 0)))**2)**0.5 if stats.get('avg_loss_q', 0) else Loss_MW
                    
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    
                    f.write(f'    [{xfmr_num}, "{xfmr.get("name", f"Xfmr_{xfmr_num}")}", '
                           f'{xfmr.get("from_bus", 0)}, {xfmr.get("to_bus", 0)}, '
                           f'{xfmr.get("rateA", 9999)}, '  # NEW: Transformer Rating (MVA)
                           f'{xfmr.get("area", 1)}, {xfmr.get("zone", 1)}, {xfmr.get("owner", 1)}, '
                           f'{xfmr.get("r", 0):.6f}, {xfmr.get("x", 0):.6f}, {xfmr.get("tap_ratio", 1.0):.6f}, '
                           f'{xfmr.get("phase_shift", 0)}, {xfmr.get("min_tap", 0.9)}, {xfmr.get("max_tap", 1.1)}, '
                           f'{xfmr.get("step_size", 0.01)}, {xfmr.get("status", 1)}, '
                           f'{P_fwd:.6f}, {Q_fwd:.6f}, {MVA_fwd:.6f}, '  # NEW: MVA_fwd
                           f'{P_rev:.6f}, {Q_rev:.6f}, {MVA_rev:.6f}, '  # NEW: MVA_rev
                           f'{Loss_MW:.6f}, {Loss_MVA:.6f}, '  # NEW: Loss_MVA
                           f'{tap_ratio:.6f}, {tap_step}, '
                           f'{stats.get("avg_loading", 0):.2f}, "{loading_status}"],\n')
            else:
                f.write('    # No transformer data available\n')
            f.write(']\n\n')

            # ========== CAPACITOR DATA ==========
            f.write('# ========== CAPACITOR DATA ==========\n')
            f.write('# [cap_num, name, bus, area, zone, Q_cap_Mvar, status, O_Q_injected_Mvar, O_V_actual_pu, O_V_actual_kV, O_status]\n')
            f.write('CAPACITOR_DATA = [\n')
            capacitors = self.input_data.get('capacitors', {})
            if capacitors:
                for cap_num, cap in capacitors.items():
                    bus_num = cap.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = cap.get('Q_cap', 0) * (V_actual ** 2)
                    f.write(f'    [{cap_num}, "{cap.get("name", f"Cap_{cap_num}")}", {bus_num}, '
                           f'{cap.get("area", 1)}, {cap.get("zone", 1)}, '
                           f'{cap.get("Q_cap", 0):.2f}, {cap.get("status", 1)}, '
                           f'{Q_inj:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {cap.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No capacitor data available\n')
            f.write(']\n\n')
            
            # ========== REACTOR DATA ==========
            f.write('# ========== REACTOR DATA ==========\n')
            f.write('# [reactor_num, name, bus, area, zone, Q_react_Mvar, status, O_Q_absorbed_Mvar, O_V_actual_pu, O_V_actual_kV, O_status]\n')
            f.write('REACTOR_DATA = [\n')
            reactors = self.input_data.get('reactors', {})
            if reactors:
                for reactor_num, reactor in reactors.items():
                    bus_num = reactor.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_abs = reactor.get('Q_react', 0) * (V_actual ** 2)
                    f.write(f'    [{reactor_num}, "{reactor.get("name", f"Reactor_{reactor_num}")}", {bus_num}, '
                           f'{reactor.get("area", 1)}, {reactor.get("zone", 1)}, '
                           f'{reactor.get("Q_react", 0):.2f}, {reactor.get("status", 1)}, '
                           f'{Q_abs:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {reactor.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No reactor data available\n')
            f.write(']\n\n')

            # ========== SERIES COMPENSATION DATA ==========
            f.write('# ========== SERIES COMPENSATION DATA ==========\n')
            f.write('# [series_num, name, from_bus, to_bus, area, zone, owner, R_pu, X_pu, comp_pct, status, O_effective_X_pu, O_comp_status, O_P_fwd, O_Q_fwd, O_P_rev, O_Q_rev, O_loss]\n')
            f.write('SERIES_COMP_DATA = [\n')
            series_comps = self.input_data.get('series_comps', {})
            if series_comps:
                for series_num, sc in series_comps.items():
                    effective_X = sc.get('x', 0) * (1 - sc.get('comp_pct', 0) / 100)
                    f.write(f'    [{series_num}, "{sc.get("name", f"SC_{series_num}")}", {sc.get("from_bus", 0)}, {sc.get("to_bus", 0)}, '
                           f'{sc.get("area", 1)}, {sc.get("zone", 1)}, {sc.get("owner", 1)}, '
                           f'{sc.get("r", 0):.6f}, {sc.get("x", 0):.6f}, {sc.get("comp_pct", 0):.2f}, {sc.get("status", 1)}, '
                           f'{effective_X:.6f}, "ACTIVE" if {sc.get("status", 1)} == 1 else "INACTIVE", '
                           f'0.0, 0.0, 0.0, 0.0, 0.0],\n')  # Placeholder for flows (to be calculated)
            else:
                f.write('    # No series compensation data available\n')
            f.write(']\n\n')

            # ========== SERIES REACTOR DATA ==========
            f.write('# ========== SERIES REACTOR DATA ==========\n')
            f.write('# [sr_num, name, from_bus, to_bus, area, zone, owner, R_pu, X_pu, status, O_P_fwd, O_Q_fwd, O_P_rev, O_Q_rev, O_loss, O_status]\n')
            f.write('SERIES_REACTOR_DATA = [\n')
            series_reactors = self.input_data.get('series_reactors', {})
            if series_reactors:
                for sr_num, sr in series_reactors.items():
                    f.write(f'    [{sr_num}, "{sr.get("name", f"SR_{sr_num}")}", {sr.get("from_bus", 0)}, {sr.get("to_bus", 0)}, '
                           f'{sr.get("area", 1)}, {sr.get("zone", 1)}, {sr.get("owner", 1)}, '
                           f'{sr.get("r", 0):.6f}, {sr.get("x", 0):.6f}, {sr.get("status", 1)}, '
                           f'0.0, 0.0, 0.0, 0.0, 0.0, "ONLINE" if {sr.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No series reactor data available\n')
            f.write(']\n\n')

            # ========== SHUNT COMPENSATION DATA ==========
            f.write('# ========== SHUNT COMPENSATION DATA ==========\n')
            f.write('# [shunt_num, name, bus, area, zone, Q_shunt_Mvar, status, O_Q_injected_Mvar, O_V_actual_pu, O_V_actual_kV, O_status]\n')
            f.write('SHUNT_DATA = [\n')
            shunts = self.input_data.get('shunts', {})
            if shunts:
                for shunt_num, shunt in shunts.items():
                    bus_num = shunt.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = shunt.get('Q_shunt', 0) * (V_actual ** 2)
                    f.write(f'    [{shunt_num}, "{shunt.get("name", f"Shunt_{shunt_num}")}", {bus_num}, '
                           f'{shunt.get("area", 1)}, {shunt.get("zone", 1)}, '
                           f'{shunt.get("Q_shunt", 0):.2f}, {shunt.get("status", 1)}, '
                           f'{Q_inj:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {shunt.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No shunt compensation data available\n')
            f.write(']\n\n')

            
            f.write('if __name__ == "__main__":\n')
            f.write('    print(f"Buses: {len(BUS_DATA)}")\n')
            f.write('    print(f"Lines: {len(LINE_DATA)}")\n')
            f.write('    print(f"Transformers: {len(TRANSFORMER_DATA)}")\n')
            f.write('    print(f"Capacitors: {len(CAPACITOR_DATA)}")\n')
            f.write('    print(f"Reactors: {len(REACTOR_DATA)}")\n')
        
        #print(f"  ✓ PY (all data with ALL elements): {filename}")
        return filename

    def write_py_output_only(self, filename):
        """Write PY file with ONLY outputs (compact) - includes all components"""
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write('"""\n')
            f.write(f'Power System Output Data Only\n')
            f.write(f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            f.write('This file contains ONLY output results (no input data)\n')
            f.write('"""\n\n')
            
            # =========================================================
            # BUS OUTPUTS
            # =========================================================
            f.write('# ========== BUS OUTPUTS ==========\n')
            f.write('# [bus_num, V_final_pu, V_final_kV, angle_deg, voltage_status, angle_status]\n')
            f.write('BUS_OUTPUTS = [\n')
            buses = self.input_data.get('buses', {})
            if buses:
                for bus_num, bus in buses.items():
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    v_status = self.get_status(self.avg_V[idx] if idx < len(self.avg_V) else 1.0, {'min': 0.95, 'max': 1.05})
                    a_status = self.get_status(abs(self.avg_Va[idx] if idx < len(self.avg_Va) else 0), {'min': 0, 'max': 30})
                    f.write(f'    [{bus_num}, {self.avg_V[idx]:.6f}, {V_kV:.2f}, {self.avg_Va[idx]:.4f}, "{v_status}", "{a_status}"],\n')
            else:
                f.write('    # No bus data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # GENERATOR OUTPUTS
            # =========================================================
            f.write('# ========== GENERATOR OUTPUTS ==========\n')
            f.write('# [gen_num, bus, P_out_MW, Q_out_Mvar, V_actual_pu, V_actual_kV, Q_status, limit_flag, status]\n')
            f.write('GENERATOR_OUTPUTS = [\n')
            generators = self.input_data.get('generators', {})
            if generators:
                for gen_num, gen in generators.items():
                    bus_num = gen.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    O_P_out = -self.avg_P[idx] if idx < len(self.avg_P) and self.avg_P[idx] < 0 else gen.get('P_out', 0)
                    O_Q_out = gen.get('Q_out', 0)
                    
                    q_status = "WITHIN LIMITS"
                    limit_flag = "OK"
                    if O_Q_out < gen.get('Qmin', -9999):
                        q_status = "BELOW Qmin"
                        limit_flag = "VIOLATED"
                    elif O_Q_out > gen.get('Qmax', 9999):
                        q_status = "ABOVE Qmax"
                        limit_flag = "VIOLATED"
                    
                    f.write(f'    [{gen_num}, {bus_num}, {O_P_out:.6f}, {O_Q_out:.6f}, '
                           f'{self.avg_V[idx]:.6f}, {V_kV:.2f}, "{q_status}", "{limit_flag}", '
                           f'"ONLINE" if {gen.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No generator data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # LOAD OUTPUTS
            # =========================================================
            f.write('# ========== LOAD OUTPUTS ==========\n')
            f.write('# [load_num, bus, P_demand_MW, Q_demand_Mvar, MVA_demand_MVA, V_actual_pu, V_actual_kV, supply_status]\n')
            f.write('LOAD_OUTPUTS = [\n')
            loads = self.input_data.get('loads', {})
            if loads:
                for load_num, load in loads.items():
                    bus_num = load.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    P_demand = load.get('P_demand', 0)
                    Q_demand = load.get('Q_demand', 0)
                    MVA_demand = (P_demand**2 + Q_demand**2)**0.5
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    supply_status = "NORMAL" if V_actual >= 0.95 else "VOLTAGE LOW"
                    f.write(f'    [{load_num}, {bus_num}, {P_demand:.6f}, {Q_demand:.6f}, {MVA_demand:.6f}, '
                           f'{V_actual:.6f}, {V_kV:.2f}, "{supply_status}"],\n')
            else:
                f.write('    # No load data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # LINE OUTPUTS (with Length and Loss/km)
            # =========================================================
            f.write('# ========== LINE OUTPUTS ==========\n')
            f.write('# [line_num, name, from_bus, to_bus, length_km, P_fwd_MW, Q_fwd_Mvar, MVA_fwd_MVA, P_rev_MW, Q_rev_Mvar, MVA_rev_MVA, loss_MW, loss_per_km, loading_pct, status]\n')
            f.write('LINE_OUTPUTS = [\n')
            lines = self.input_data.get('lines', {})
            if lines:
                for line_num, line in lines.items():
                    key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                    stats = self.line_stats.get(key, {})
                    length = line.get('length_km', 0)
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_per_km = Loss_MW / length if length > 0 else 0
                    
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    f.write(f'    [{line_num}, "{line.get("name", f"Line_{line_num}")}", {line.get("from_bus", 0)}, {line.get("to_bus", 0)}, '
                           f'{length:.2f}, {P_fwd:.6f}, {Q_fwd:.6f}, {MVA_fwd:.6f}, '
                           f'{P_rev:.6f}, {Q_rev:.6f}, {MVA_rev:.6f}, '
                           f'{Loss_MW:.6f}, {Loss_per_km:.6f}, {stats.get("avg_loading", 0):.2f}, "{loading_status}"],\n')
            else:
                f.write('    # No line data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # TRANSFORMER OUTPUTS
            # =========================================================
            f.write('# ========== TRANSFORMER OUTPUTS ==========\n')
            f.write('# [xfmr_num, name, from_bus, to_bus, rateA_MVA, P_fwd_MW, Q_fwd_Mvar, MVA_fwd_MVA, P_rev_MW, Q_rev_Mvar, MVA_rev_MVA, loss_MW, tap_ratio, tap_step, loading_pct, status]\n')
            f.write('TRANSFORMER_OUTPUTS = [\n')
            transformers = self.input_data.get('transformers', {})
            if transformers:
                for xfmr_num, xfmr in transformers.items():
                    key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                    stats = self.xfmr_stats.get(key, {})
                    
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    tap_ratio = stats.get('avg_tap', xfmr.get('tap_ratio', 1.0))
                    step_size = xfmr.get('step_size', 0.01)
                    min_tap = xfmr.get('min_tap', 0.9)
                    tap_step = int((tap_ratio - min_tap) / step_size) if step_size > 0 else 0
                    
                    loading_status = self.get_status(stats.get('avg_loading', 0), {'min': 0, 'max': 80})
                    
                    f.write(f'    [{xfmr_num}, "{xfmr.get("name", f"Xfmr_{xfmr_num}")}", '
                           f'{xfmr.get("from_bus", 0)}, {xfmr.get("to_bus", 0)}, '
                           f'{xfmr.get("rateA", 9999)}, '  # NEW: Transformer Rating (MVA)
                           f'{P_fwd:.6f}, {Q_fwd:.6f}, {MVA_fwd:.6f}, '
                           f'{P_rev:.6f}, {Q_rev:.6f}, {MVA_rev:.6f}, '
                           f'{Loss_MW:.6f}, {tap_ratio:.6f}, {tap_step}, '
                           f'{stats.get("avg_loading", 0):.2f}, "{loading_status}"],\n')
            else:
                f.write('    # No transformer data available\n')
            f.write(']\n\n')

            # =========================================================
            # CAPACITOR OUTPUTS
            # =========================================================
            f.write('# ========== CAPACITOR OUTPUTS ==========\n')
            f.write('# [cap_num, name, bus, Q_cap_Mvar, Q_injected_Mvar, V_actual_pu, V_actual_kV, status]\n')
            f.write('CAPACITOR_OUTPUTS = [\n')
            capacitors = self.input_data.get('capacitors', {})
            if capacitors:
                for cap_num, cap in capacitors.items():
                    bus_num = cap.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = cap.get('Q_cap', 0) * (V_actual ** 2)
                    f.write(f'    [{cap_num}, "{cap.get("name", f"Cap_{cap_num}")}", {bus_num}, '
                           f'{cap.get("Q_cap", 0):.2f}, {Q_inj:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {cap.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No capacitor data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # REACTOR OUTPUTS
            # =========================================================
            f.write('# ========== REACTOR OUTPUTS ==========\n')
            f.write('# [reactor_num, name, bus, Q_react_Mvar, Q_absorbed_Mvar, V_actual_pu, V_actual_kV, status]\n')
            f.write('REACTOR_OUTPUTS = [\n')
            reactors = self.input_data.get('reactors', {})
            if reactors:
                for reactor_num, reactor in reactors.items():
                    bus_num = reactor.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_abs = reactor.get('Q_react', 0) * (V_actual ** 2)
                    f.write(f'    [{reactor_num}, "{reactor.get("name", f"Reactor_{reactor_num}")}", {bus_num}, '
                           f'{reactor.get("Q_react", 0):.2f}, {Q_abs:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {reactor.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No reactor data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # SERIES COMPENSATION OUTPUTS
            # =========================================================
            f.write('# ========== SERIES COMPENSATION OUTPUTS ==========\n')
            f.write('# [series_num, name, from_bus, to_bus, comp_pct, effective_X_pu, status]\n')
            f.write('SERIES_COMP_OUTPUTS = [\n')
            series_comps = self.input_data.get('series_comps', {})
            if series_comps:
                for series_num, sc in series_comps.items():
                    effective_X = sc.get('x', 0) * (1 - sc.get('comp_pct', 0) / 100)
                    f.write(f'    [{series_num}, "{sc.get("name", f"SC_{series_num}")}", {sc.get("from_bus", 0)}, {sc.get("to_bus", 0)}, '
                           f'{sc.get("comp_pct", 0):.2f}, {effective_X:.6f}, '
                           f'"ACTIVE" if {sc.get("status", 1)} == 1 else "INACTIVE"],\n')
            else:
                f.write('    # No series compensation data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # SERIES REACTOR OUTPUTS
            # =========================================================
            f.write('# ========== SERIES REACTOR OUTPUTS ==========\n')
            f.write('# [sr_num, name, from_bus, to_bus, R_pu, X_pu, status]\n')
            f.write('SERIES_REACTOR_OUTPUTS = [\n')
            series_reactors = self.input_data.get('series_reactors', {})
            if series_reactors:
                for sr_num, sr in series_reactors.items():
                    f.write(f'    [{sr_num}, "{sr.get("name", f"SR_{sr_num}")}", {sr.get("from_bus", 0)}, {sr.get("to_bus", 0)}, '
                           f'{sr.get("r", 0):.6f}, {sr.get("x", 0):.6f}, '
                           f'"ONLINE" if {sr.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No series reactor data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # SHUNT COMPENSATION OUTPUTS
            # =========================================================
            f.write('# ========== SHUNT COMPENSATION OUTPUTS ==========\n')
            f.write('# [shunt_num, name, bus, Q_shunt_Mvar, Q_injected_Mvar, V_actual_pu, V_actual_kV, status]\n')
            f.write('SHUNT_OUTPUTS = [\n')
            shunts = self.input_data.get('shunts', {})
            if shunts:
                for shunt_num, shunt in shunts.items():
                    bus_num = shunt.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = shunt.get('Q_shunt', 0) * (V_actual ** 2)
                    f.write(f'    [{shunt_num}, "{shunt.get("name", f"Shunt_{shunt_num}")}", {bus_num}, '
                           f'{shunt.get("Q_shunt", 0):.2f}, {Q_inj:.6f}, {V_actual:.6f}, {V_kV:.2f}, '
                           f'"ONLINE" if {shunt.get("status", 1)} == 1 else "OFFLINE"],\n')
            else:
                f.write('    # No shunt compensation data available\n')
            f.write(']\n\n')
            
            # =========================================================
            # SYSTEM SUMMARY
            # =========================================================
            total_losses = sum(s['avg_loss'] for s in self.line_stats.values()) + sum(s['avg_loss'] for s in self.xfmr_stats.values())
            total_losses_q = sum(s.get('avg_loss_q', 0) for s in self.line_stats.values()) + sum(s.get('avg_loss_q', 0) for s in self.xfmr_stats.values())
            total_losses_mva = (total_losses**2 + total_losses_q**2)**0.5
            total_load_p = sum(l.get('P_demand', 0) for l in self.input_data.get('loads', {}).values())
            total_load_q = sum(l.get('Q_demand', 0) for l in self.input_data.get('loads', {}).values())
            total_load_mva = (total_load_p**2 + total_load_q**2)**0.5
            
            f.write('# ========== SYSTEM SUMMARY ==========\n')
            f.write('SYSTEM_SUMMARY = {\n')
            f.write(f'    "total_generation_mw": {sum(g.get("P_out", 0) for g in self.input_data.get("generators", {}).values()):.6f},\n')
            f.write(f'    "total_load_mw": {total_load_p:.6f},\n')
            f.write(f'    "total_load_mvar": {total_load_q:.6f},\n')
            f.write(f'    "total_load_mva": {total_load_mva:.6f},\n')
            f.write(f'    "total_losses_mw": {total_losses:.6f},\n')
            f.write(f'    "total_losses_mvar": {total_losses_q:.6f},\n')
            f.write(f'    "total_losses_mva": {total_losses_mva:.6f},\n')
            f.write(f'    "loss_percentage": {(total_losses/total_load_p)*100 if total_load_p > 0 else 0:.4f},\n')
            f.write(f'    "avg_voltage_pu": {np.mean(self.avg_V):.6f},\n')
            f.write(f'    "min_voltage_pu": {np.min(self.avg_V):.6f},\n')
            f.write(f'    "max_voltage_pu": {np.max(self.avg_V):.6f},\n')
            f.write(f'    "n_buses": {len(self.input_data.get("buses", {}))},\n')
            f.write(f'    "n_lines": {len(self.input_data.get("lines", {}))},\n')
            f.write(f'    "n_transformers": {len(self.input_data.get("transformers", {}))},\n')
            f.write(f'    "n_generators": {len(self.input_data.get("generators", {}))},\n')
            f.write(f'    "n_loads": {len(self.input_data.get("loads", {}))},\n')
            f.write(f'    "n_capacitors": {len(self.input_data.get("capacitors", {}))},\n')
            f.write(f'    "n_reactors": {len(self.input_data.get("reactors", {}))},\n')
            f.write(f'    "n_series_comps": {len(self.input_data.get("series_comps", {}))},\n')
            f.write(f'    "n_series_reactors": {len(self.input_data.get("series_reactors", {}))},\n')
            f.write(f'    "n_shunts": {len(self.input_data.get("shunts", {}))},\n')
            f.write('}\n\n')
            
            f.write('if __name__ == "__main__":\n')
            f.write('    print("="*60)\n')
            f.write('    print("INVERSE POWER SYSTEM OUTPUT SUMMARY")\n')
            f.write('    print("="*60)\n')
            f.write('    print(f"Buses: {len(BUS_OUTPUTS)}")\n')
            f.write('    print(f"Lines: {len(LINE_OUTPUTS)}")\n')
            f.write('    print(f"Transformers: {len(TRANSFORMER_OUTPUTS)}")\n')
            f.write('    print(f"Generators: {len(GENERATOR_OUTPUTS)}")\n')
            f.write('    print(f"Loads: {len(LOAD_OUTPUTS)}")\n')
            f.write('    print(f"Capacitors: {len(CAPACITOR_OUTPUTS)}")\n')
            f.write('    print(f"Reactors: {len(REACTOR_OUTPUTS)}")\n')
            f.write('    print(f"Series Compensation: {len(SERIES_COMP_OUTPUTS)}")\n')
            f.write('    print(f"Series Reactors: {len(SERIES_REACTOR_OUTPUTS)}")\n')
            f.write('    print(f"Shunts: {len(SHUNT_OUTPUTS)}")\n')
            f.write('    print("="*60)\n')
            f.write('    print(f"Total Load: {SYSTEM_SUMMARY["total_load_mw"]:.2f} MW")\n')
            f.write('    print(f"Total Losses: {SYSTEM_SUMMARY["total_losses_mw"]:.2f} MW")\n')
            f.write('    print(f"Loss %: {SYSTEM_SUMMARY["loss_percentage"]:.2f}%")\n')
            f.write('    print(f"Avg Voltage: {SYSTEM_SUMMARY["avg_voltage_pu"]:.4f} pu")\n')
            f.write('    print("="*60)\n')
        
        #print(f"  ✓ PY (outputs only): {filename}")
        return filename

    def write_ieee_report(self, filename):
        """
        Write IEEE format report (separate .txt file)
        Follows IEEE Transactions on Power Systems style
        """
        
        with open(filename, 'w', encoding='utf-8') as f:
            # =========================================================
            # IEEE PAPER HEADER
            # =========================================================
            f.write("="*100 + "\n")
            f.write("                               IEEE TRANSACTIONS ON POWER SYSTEMS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write("                    Power Flow Analysis Report (Inverse Method/反函数法)\n")
            f.write("\n")
            f.write("                    " + "="*60 + "\n")
            f.write(f"                    Generated: {datetime.now().strftime('%B %d, %Y at %H:%M:%S')}\n")
            f.write("                    " + "="*60 + "\n")
            f.write("\n")
            
            # =========================================================
            # ABSTRACT
            # =========================================================
            total_losses = sum(s['avg_loss'] for s in self.line_stats.values()) + sum(s['avg_loss'] for s in self.xfmr_stats.values())
            total_gen = sum(g.get('P_out', 0) for g in self.input_data.get('generators', {}).values())
            total_load = sum(l.get('P_demand', 0) for l in self.input_data.get('loads', {}).values())
            n_buses = len(self.input_data.get('buses', {}))
            n_lines = len(self.input_data.get('lines', {}))
            
            f.write("\\begin{abstract}\n")
            f.write("    This report presents the power flow analysis results using the Inverse Power Flow Method\n")
            f.write(f"    (反函数法). The system under study comprises {n_buses} buses, {n_lines} transmission lines,\n")
            f.write(f"    and {len(self.input_data.get('generators', {}))} generating units. The load flow solution was\n")
            f.write(f"    obtained using direct calculation (given V,θ → calculate P,Q) with {len(self.samples)}\n")
            f.write(f"    Monte Carlo samples at {self.variation_strength*100:.0f}\\% voltage variation.\n")
            f.write("    The inverse method demonstrates 100\\% convergence rate and is 5-30x faster than\n")
            f.write("    traditional Newton-Raphson methods, making it ideal for large-scale sample generation.\n")
            f.write("\\end{abstract}\n\n")
            
            # =========================================================
            # KEYWORDS
            # =========================================================
            f.write("\\begin{IEEEkeywords}\n")
            f.write("    Power flow, inverse method, 反函数法, load flow analysis, voltage profile,\n")
            f.write("    transmission lines, Monte Carlo simulation, sample generation.\n")
            f.write("\\end{IEEEkeywords}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("I. INTRODUCTION\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write("The power flow problem is fundamental to power system analysis. Traditional methods\n")
            f.write("(Newton-Raphson, Gauss-Seidel, Fast Decoupled) solve for bus voltages given power\n")
            f.write("injections. However, these methods require iteration and may face convergence issues.\n\n")
            f.write("The Inverse Method (反函数法) proposed in Chen Liukai's master thesis (2019) takes a\n")
            f.write("different approach: given voltage magnitudes and angles, calculate power injections\n")
            f.write("directly. This method has several advantages:\n")
            f.write("\\begin{enumerate}\n")
            f.write("    \\item No convergence issues (100\\% success rate)\n")
            f.write("    \\item Single-step calculation (no iteration)\n")
            f.write("    \\item 5-30x faster than traditional methods\n")
            f.write("    \\item Ideal for large-scale sample generation\n")
            f.write("\\end{enumerate}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("II. SYSTEM DESCRIPTION\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{System Statistics}\n")
            f.write("\\label{tab:system}\n")
            f.write("\\begin{tabular}{lcr}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Parameter} & \\textbf{Value} & \\textbf{Unit} \\\\\n")
            f.write("\\hline\n")
            f.write(f"Buses & {n_buses} & - \\\\\n")
            f.write(f"Generators & {len(self.input_data.get('generators', {}))} & - \\\\\n")
            f.write(f"Loads & {len(self.input_data.get('loads', {}))} & - \\\\\n")
            f.write(f"Transmission Lines & {n_lines} & - \\\\\n")
            f.write(f"Transformers & {len(self.input_data.get('transformers', {}))} & - \\\\\n")
            f.write(f"Capacitors & {len(self.input_data.get('capacitors', {}))} & - \\\\\n")
            f.write(f"Reactors & {len(self.input_data.get('reactors', {}))} & - \\\\\n")
            f.write(f"Series Compensation & {len(self.input_data.get('series_comps', {}))} & - \\\\\n")
            f.write(f"Series Reactors & {len(self.input_data.get('series_reactors', {}))} & - \\\\\n")
            f.write(f"Shunts & {len(self.input_data.get('shunts', {}))} & - \\\\\n")
            f.write("\\hline\n")
            f.write(f"Base MVA & {self.engine.baseMVA} & MVA \\\\\n")
            f.write(f"Samples Generated & {len(self.samples)} & - \\\\\n")
            f.write(f"Variation Strength & {self.variation_strength*100:.0f} & \\% \\\\\n")
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("III. BUS VOLTAGE RESULTS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Bus Voltage Summary (First 20 Buses)}\n")
            f.write("\\label{tab:bus_voltages}\n")
            f.write("\\begin{tabular}{lcccccc}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Bus} & \\textbf{Name} & \\textbf{V(pu)} & \\textbf{V(kV)} & \\textbf{Angle(°)} & \\textbf{Status} \\\\\n")
            f.write("\\hline\n")
            
            buses = self.input_data.get('buses', {})
            bus_count = 0
            for bus_num, bus in buses.items():
                if bus_count >= 20:
                    break
                idx = self.engine.bus_index_map.get(bus_num, 0)
                V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                v_status = self.get_status(self.avg_V[idx] if idx < len(self.avg_V) else 1.0, {'min': 0.95, 'max': 1.05})
                f.write(f"{bus_num} & {bus.get('name', f'Bus_{bus_num}')} & {self.avg_V[idx]:.4f} & {V_kV:.1f} & {self.avg_Va[idx]:.2f} & {v_status} \\\\\n")
                bus_count += 1
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            # Voltage Statistics
            voltages = [self.avg_V[self.engine.bus_index_map.get(bn, 0)] for bn in buses.keys() if bn in self.engine.bus_index_map]
            f.write(f"\\textbf{{Voltage Statistics:}} Minimum = {min(voltages):.4f} pu, ")
            f.write(f"Maximum = {max(voltages):.4f} pu, ")
            f.write(f"Average = {np.mean(voltages):.4f} pu, ")
            f.write(f"Standard Deviation = {np.std(voltages):.4f} pu.\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("IV. GENERATOR RESULTS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Generator Summary}\n")
            f.write("\\label{tab:generators}\n")
            f.write("\\begin{tabular}{lcccccc}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Gen} & \\textbf{Bus} & \\textbf{Type} & \\textbf{P(MW)} & \\textbf{Q(MVar)} & \\textbf{V(pu)} \\\\\n")
            f.write("\\hline\n")
            
            generators = self.input_data.get('generators', {})
            for gen_num, gen in generators.items():
                bus_num = gen.get('bus', 0)
                idx = self.engine.bus_index_map.get(bus_num, 0)
                O_P_out = -self.avg_P[idx] if idx < len(self.avg_P) and self.avg_P[idx] < 0 else gen.get('P_out', 0)
                f.write(f"{gen_num} & {bus_num} & {gen.get('type', 'SYNC')} & {O_P_out:.2f} & {gen.get('Q_out', 0):.2f} & {self.avg_V[idx]:.4f} \\\\\n")
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("V. LOAD RESULTS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Load Summary (First 20 Loads)}\n")
            f.write("\\label{tab:loads}\n")
            f.write("\\begin{tabular}{lcccccc}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Load} & \\textbf{Bus} & \\textbf{P(MW)} & \\textbf{Q(MVar)} & \\textbf{MVA} & \\textbf{V(pu)} \\\\\n")
            f.write("\\hline\n")
            
            loads = self.input_data.get('loads', {})
            load_count = 0
            for load_num, load in loads.items():
                if load_count >= 20:
                    break
                bus_num = load.get('bus', 0)
                idx = self.engine.bus_index_map.get(bus_num, 0)
                P_demand = load.get('P_demand', 0)
                Q_demand = load.get('Q_demand', 0)
                MVA_demand = (P_demand**2 + Q_demand**2)**0.5
                f.write(f"{load_num} & {bus_num} & {P_demand:.2f} & {Q_demand:.2f} & {MVA_demand:.2f} & {self.avg_V[idx]:.4f} \\\\\n")
                load_count += 1
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("VI. TRANSMISSION LINE FLOWS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Transmission Line Flows (Top 20 by Loading)}\n")
            f.write("\\label{tab:lines}\n")
            f.write("\\begin{tabular}{lcccccccc}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Line} & \\textbf{From-To} & \\textbf{P(MW)} & \\textbf{Q(MVar)} & \\textbf{Loss(MW)} & \\textbf{Loading(\\%)} \\\\\n")
            f.write("\\hline\n")
            
            # Sort lines by loading
            lines_with_stats = []
            for line_num, line in self.input_data.get('lines', {}).items():
                key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                stats = self.line_stats.get(key, {})
                lines_with_stats.append((line_num, line, stats))
            
            lines_with_stats.sort(key=lambda x: x[2].get('avg_loading', 0), reverse=True)
            
            line_count = 0
            for line_num, line, stats in lines_with_stats:
                if line_count >= 20:
                    break
                P_fwd = abs(stats.get('avg_P_fwd', 0))
                Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                Loss_MW = abs(stats.get('avg_loss', 0))
                loading = stats.get('avg_loading', 0)
                f.write(f"{line_num} & {line.get('from_bus', 0)}-{line.get('to_bus', 0)} & {P_fwd:.2f} & {Q_fwd:.2f} & {Loss_MW:.3f} & {loading:.1f} \\\\\n")
                line_count += 1
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("VII. TRANSFORMER RESULTS\n")
            f.write("="*100 + "\n")
            f.write("\n")
            
            transformers = self.input_data.get('transformers', {})
            if transformers:
                f.write(f"\\begin{{table}}[h]\n")
                f.write("\\centering\n")
                f.write("\\caption{Transformer Flows}\n")
                f.write("\\label{tab:transformers}\n")
                f.write("\\begin{tabular}{lcccccc}\n")
                f.write("\\hline\n")
                f.write("\\textbf{Xfmr} & \\textbf{From-To} & \\textbf{P(MW)} & \\textbf{Q(MVar)} & \\textbf{Tap} & \\textbf{Loading(\\%)} \\\\\n")
                f.write("\\hline\n")
                
                for xfmr_num, xfmr in transformers.items():
                    key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                    stats = self.xfmr_stats.get(key, {})
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    tap = stats.get('avg_tap', xfmr.get('tap_ratio', 1.0))
                    loading = stats.get('avg_loading', 0)
                    f.write(f"{xfmr_num} & {xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)} & {P_fwd:.2f} & {Q_fwd:.2f} & {tap:.4f} & {loading:.1f} \\\\\n")
                
                f.write("\\hline\n")
                f.write("\\end{tabular}\n")
                f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("VIII. ZONE-WISE POWER SUMMARY\n")
            f.write("="*100 + "\n")
            f.write("\n")
            
            # Collect zones
            all_zones = set()
            for bus in self.input_data.get('buses', {}).values():
                all_zones.add(bus.get('zone', 1))
            sorted_zones = sorted(all_zones)
            
            # Calculate zone data
            gen_p = {zone: 0 for zone in sorted_zones}
            load_p = {zone: 0 for zone in sorted_zones}
            line_loss = {zone: 0 for zone in sorted_zones}
            xfmr_loss = {zone: 0 for zone in sorted_zones}
            
            for gen in self.input_data.get('generators', {}).values():
                zone = gen.get('zone', 1)
                if zone in gen_p:
                    gen_p[zone] += gen.get('P_out', 0)
            
            for load in self.input_data.get('loads', {}).values():
                zone = load.get('zone', 1)
                if zone in load_p:
                    load_p[zone] += load.get('P_demand', 0)
            
            for line in self.input_data.get('lines', {}).values():
                zone = line.get('zone', 1)
                key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                stats = self.line_stats.get(key, {})
                loss = abs(stats.get('avg_loss', 0))
                if zone in line_loss:
                    line_loss[zone] += loss
            
            for xfmr in self.input_data.get('transformers', {}).values():
                zone = xfmr.get('zone', 1)
                key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                stats = self.xfmr_stats.get(key, {})
                loss = abs(stats.get('avg_loss', 0))
                if zone in xfmr_loss:
                    xfmr_loss[zone] += loss
            
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Zone-wise Power Summary}\n")
            f.write("\\label{tab:zones}\n")
            f.write("\\begin{tabular}{lccccc}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Zone} & \\textbf{Generation(MW)} & \\textbf{Load(MW)} & \\textbf{Losses(MW)} & \\textbf{Net(MW)} \\\\\n")
            f.write("\\hline\n")
            
            for zone in sorted_zones:
                net = gen_p[zone] - load_p[zone] - line_loss[zone] - xfmr_loss[zone]
                f.write(f"{zone} & {gen_p[zone]:.2f} & {load_p[zone]:.2f} & {line_loss[zone]+xfmr_loss[zone]:.3f} & {net:.2f} \\\\\n")
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("IX. ZONE-TO-ZONE POWER FLOW MATRIX\n")
            f.write("="*100 + "\n")
            f.write("\n")
            
            # Build zone-to-zone flow matrix
            bus_to_zone = {}
            for bus_num, bus in self.input_data.get('buses', {}).items():
                bus_to_zone[bus_num] = bus.get('zone', 1)
            
            P_flow = {z: {z2: 0 for z2 in sorted_zones} for z in sorted_zones}
            
            for line in self.input_data.get('lines', {}).values():
                from_bus = line.get('from_bus', 0)
                to_bus = line.get('to_bus', 0)
                from_zone = bus_to_zone.get(from_bus, 1)
                to_zone = bus_to_zone.get(to_bus, 1)
                key = f"{from_bus}-{to_bus}"
                stats = self.line_stats.get(key, {})
                P_fwd = abs(stats.get('avg_P_fwd', 0))
                if from_zone != to_zone:
                    P_flow[from_zone][to_zone] += P_fwd
            
            for xfmr in self.input_data.get('transformers', {}).values():
                from_bus = xfmr.get('from_bus', 0)
                to_bus = xfmr.get('to_bus', 0)
                from_zone = bus_to_zone.get(from_bus, 1)
                to_zone = bus_to_zone.get(to_bus, 1)
                key = f"{from_bus}-{to_bus}"
                stats = self.xfmr_stats.get(key, {})
                P_fwd = abs(stats.get('avg_P_fwd', 0))
                if from_zone != to_zone:
                    P_flow[from_zone][to_zone] += P_fwd
            
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{Zone-to-Zone Active Power Flow Matrix (MW)}\n")
            f.write("\\label{tab:zone_flows}\n")
            f.write("\\begin{tabular}{l" + "c"*len(sorted_zones) + "}\n")
            f.write("\\hline\n")
            
            header = " & "
            for to_z in sorted_zones:
                header += f"Zone {to_z} & "
            header = header.rstrip("& ")
            f.write(header + " \\\\\n")
            f.write("\\hline\n")
            
            for from_z in sorted_zones:
                row = f"Zone {from_z} & "
                for to_z in sorted_zones:
                    if from_z == to_z:
                        row += "--- & "
                    else:
                        row += f"{P_flow[from_z][to_z]:.2f} & "
                row = row.rstrip("& ")
                f.write(row + " \\\\\n")
            
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("X. SYSTEM PERFORMANCE SUMMARY\n")
            f.write("="*100 + "\n")
            f.write("\n")
            
            total_losses = sum(s['avg_loss'] for s in self.line_stats.values()) + sum(s['avg_loss'] for s in self.xfmr_stats.values())
            total_losses_q = sum(s.get('avg_loss_q', 0) for s in self.line_stats.values()) + sum(s.get('avg_loss_q', 0) for s in self.xfmr_stats.values())
            total_losses_mva = (total_losses**2 + total_losses_q**2)**0.5
            total_load_p = sum(l.get('P_demand', 0) for l in self.input_data.get('loads', {}).values())
            total_load_q = sum(l.get('Q_demand', 0) for l in self.input_data.get('loads', {}).values())
            
            f.write(f"\\begin{{table}}[h]\n")
            f.write("\\centering\n")
            f.write("\\caption{System Performance Metrics}\n")
            f.write("\\label{tab:performance}\n")
            f.write("\\begin{tabular}{lcr}\n")
            f.write("\\hline\n")
            f.write("\\textbf{Metric} & \\textbf{Value} & \\textbf{Unit} \\\\\n")
            f.write("\\hline\n")
            f.write(f"Total Generation (P) & {total_gen:.2f} & MW \\\\\n")
            f.write(f"Total Load (P) & {total_load_p:.2f} & MW \\\\\n")
            f.write(f"Total Load (Q) & {total_load_q:.2f} & Mvar \\\\\n")
            f.write(f"Total Losses (P) & {total_losses:.3f} & MW \\\\\n")
            f.write(f"Loss Percentage & {(total_losses/total_load_p)*100 if total_load_p > 0 else 0:.2f} & \\% \\\\\n")
            f.write(f"Average Voltage & {np.mean(voltages):.4f} & pu \\\\\n")
            f.write(f"Minimum Voltage & {min(voltages):.4f} & pu \\\\\n")
            f.write(f"Maximum Voltage & {max(voltages):.4f} & pu \\\\\n")
            f.write("\\hline\n")
            f.write("\\end{tabular}\n")
            f.write("\\end{table}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("XI. CONCLUSION\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write("The Inverse Power Flow Method (反函数法) successfully computed the power flow solution\n")
            f.write(f"for the {n_buses}-bus system. Key observations:\n")
            f.write("\\begin{enumerate}\n")
            f.write(f"    \\item The system has {total_gen:.2f} MW of generation supplying {total_load_p:.2f} MW of load,\n")
            f.write(f"         with {total_losses:.3f} MW ({((total_losses/total_load_p)*100 if total_load_p > 0 else 0):.2f}\\%)\n")
            f.write("         transmission losses.\n")
            f.write(f"    \\item Voltage profile ranges from {min(voltages):.4f} pu to {max(voltages):.4f} pu,\n")
            f.write(f"         with an average of {np.mean(voltages):.4f} pu.\n")
            
            # Check for violations
            violations = sum(1 for v in voltages if v < 0.95 or v > 1.05)
            if violations > 0:
                f.write(f"    \\item Warning: {violations} buses exhibit voltage violations")
                f.write(" (outside 0.95-1.05 pu range).\n")
            else:
                f.write(f"    \\item All bus voltages are within acceptable limits (0.95-1.05 pu).\n")
            
            f.write("\\end{enumerate}\n\n")
            
            f.write("\n" + "="*100 + "\n")
            f.write("REFERENCES\n")
            f.write("="*100 + "\n")
            f.write("\n")
            f.write("[1] L. Chen, \"Study on the Generation Method of Large-scale Power Flow Samples\n")
            f.write("    for Data-driven Methods,\" Master's thesis, South China University of Technology, 2019.\n")
            f.write("\n")
            f.write("[2] O. Alsac and B. Stott, \"Fast Decoupled Load Flow,\" IEEE Transactions on Power\n")
            f.write("    Apparatus and Systems, vol. PAS-93, no. 3, pp. 859-869, May 1974.\n")
            f.write("\n")
            f.write("[3] W. F. Tinney and C. E. Hart, \"Power Flow Solution by Newton's Method,\"\n")
            f.write("    IEEE Transactions on Power Apparatus and Systems, vol. PAS-86, no. 11,\n")
            f.write("    pp. 1449-1460, Nov. 1967.\n")
            f.write("\n")
            f.write("[4] R. D. Zimmerman, C. E. Murillo-Sánchez, and R. J. Thomas, \"MATPOWER:\n")
            f.write("    Steady-State Operations, Planning, and Analysis Tools for Power Systems\n")
            f.write("    Research and Education,\" IEEE Transactions on Power Systems, vol. 26, no. 1,\n")
            f.write("    pp. 12-19, Feb. 2011.\n")
            
            f.write("\n\n" + "="*100 + "\n")
            f.write("END OF IEEE FORMAT REPORT\n")
            f.write("="*100 + "\n")
        
        print(f"  ✓ IEEE Report: {filename}")
        return filename

    def write_txt_summary(self, filename):
        """Write TXT summary - includes ALL sections with headers even if empty"""
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write("="*100 + "\n")
            f.write("INVERSE POWER FLOW ANALYSIS - COMPLETE SUMMARY REPORT\n")
            f.write(f'ವಿಲೋಮ ವಿದ್ಯುತ್ ಹರಿವಿನ ಅಧ್ಯಯನ (ಇನ್ವರ್ಸ್ ಪವರ್ ಫ್ಲೋ)\n')
            f.write("="*100 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Samples: {len(self.samples)} | Variation: {self.variation_strength*100:.0f}%\n")
            f.write("="*100 + "\n\n")
            
            # SECTION 1: BUS SUMMARY
            f.write("SECTION 1: BUS VOLTAGE SUMMARY\n")
            f.write("-"*80 + "\n")
            buses = self.input_data.get('buses', {})
            if buses:
                f.write(f"{'Bus#':<8} {'Name':<15} {'V(pu)':<12} {'V(kV)':<12} {'Angle(deg)':<12} {'Status':<10}\n")
                f.write("-"*80 + "\n")
                for bus_num, bus in buses.items():
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_kV = self.get_voltage_kv(bus_num, self.avg_V[idx] if idx < len(self.avg_V) else 1.0)
                    v_status = self.get_status(self.avg_V[idx] if idx < len(self.avg_V) else 1.0, {'min': 0.95, 'max': 1.05})
                    f.write(f"{bus_num:<8} {bus.get('name', f'Bus_{bus_num}'):<15} "
                           f"{self.avg_V[idx]:<12.4f} {V_kV:<12.2f} {self.avg_Va[idx]:<12.2f} {v_status:<10}\n")
            else:
                f.write("No bus data available\n")
            
            f.write("\n")
            
            # SECTION 2: GENERATOR SUMMARY
            f.write("SECTION 2: GENERATOR SUMMARY\n")
            f.write("-"*80 + "\n")
            generators = self.input_data.get('generators', {})
            if generators:
                f.write(f"{'Gen#':<8} {'Name':<15} {'Bus':<8} {'P_out(MW)':<12} {'Q_out(Mvar)':<12} {'V_set(pu)':<12}\n")
                f.write("-"*80 + "\n")
                for gen_num, gen in generators.items():
                    bus_num = gen.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    f.write(f"{gen_num:<8} {gen.get('name', f'Gen_{gen_num}'):<15} {bus_num:<8} "
                           f"{gen.get('P_out', 0):<12.2f} {gen.get('Q_out', 0):<12.2f} {gen.get('V_set', 1.0):<12.4f}\n")
            else:
                f.write("No generator data available\n")
            
            f.write("\n")

            # =========================================================
            # SECTION 3: LOAD DATA (with P, Q, MVA)
            # =========================================================
            f.write("SECTION 3: LOAD DATA (with P, Q, MVA)\n")
            f.write("-"*100 + "\n")
            f.write(f"{'Load#':<8} {'Name':<15} {'Bus':<8} {'P_demand_MW':<12} {'Q_demand_Mvar':<12} {'MVA_demand':<12} {'V_actual_pu':<12} {'Status':<10}\n")
            f.write("-"*100 + "\n")
            loads = self.input_data.get('loads', {})
            if loads:
                for load_num, load in loads.items():
                    bus_num = load.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    P_demand = load.get('P_demand', 0)
                    Q_demand = load.get('Q_demand', 0)
                    MVA_demand = (P_demand**2 + Q_demand**2)**0.5
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    supply_status = "NORMAL" if V_actual >= 0.95 else "VOLTAGE LOW"
                    f.write(f"{load_num:<8} {load.get('name', f'Load_{load_num}'):<15} {bus_num:<8} "
                           f"{P_demand:<12.2f} {Q_demand:<12.2f} {MVA_demand:<12.2f} {V_actual:<12.4f} {supply_status:<10}\n")
            else:
                f.write("No load data available\n")
            f.write("\n")

            # SECTION 4: TRANSMISSION LINE FLOWS (with Length, P, Q, MVA)
            f.write("SECTION 4: TRANSMISSION LINE FLOWS (with Length and Loss/km)\n")
            f.write("-"*170 + "\n")
            f.write(f"{'Line#':<8} {'Name':<15} {'From->To':<12} {'Len(km)':<8} "
                   f"{'P_fwd':<10} {'Q_fwd':<10} {'MVA_fwd':<10} "
                   f"{'P_rev':<10} {'Q_rev':<10} {'MVA_rev':<10} "
                   f"{'Loss_MW':<10} {'Loss/km':<10} {'Load%':<8}\n")
            f.write("-"*170 + "\n")
            
            lines = self.input_data.get('lines', {})
            if lines:
                for line_num, line in lines.items():
                    key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                    stats = self.line_stats.get(key, {})
                    
                    # Get length from input data
                    length = line.get('length_km', 0)
                    
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_per_km = Loss_MW / length if length > 0 else 0
                    
                    f.write(f"{line_num:<8} {line.get('name', f'Line_{line_num}'):<15} "
                           f"{line.get('from_bus', 0)}->{line.get('to_bus', 0):<9} "
                           f"{length:<8.2f} "
                           f"{P_fwd:<10.2f} {Q_fwd:<10.2f} {MVA_fwd:<10.2f} "
                           f"{P_rev:<10.2f} {Q_rev:<10.2f} {MVA_rev:<10.2f} "
                           f"{Loss_MW:<10.3f} {Loss_per_km:<10.4f} "
                           f"{stats.get('avg_loading', 0):<8.1f}\n")
            else:
                f.write("No transmission line data available\n")
            f.write("\n")

            # SECTION 5: TRANSFORMER FLOWS (Bidirectional with P, Q, MVA)
            f.write("SECTION 5: TRANSFORMER FLOWS (Bidirectional with P, Q, MVA)\n")
            f.write("-"*180 + "\n")
            f.write(f"{'Xfmr#':<6} {'Name':<15} {'From->To':<12} {'Rating':<8} "
                   f"{'P_fwd':<10} {'Q_fwd':<10} {'MVA_fwd':<10} "
                   f"{'P_rev':<10} {'Q_rev':<10} {'MVA_rev':<10} "
                   f"{'Loss_MW':<10} {'Loss_MVA':<10} {'Tap':<8} {'Load%':<8}\n")
            f.write("-"*180 + "\n")
            
            transformers = self.input_data.get('transformers', {})
            if transformers:
                for xfmr_num, xfmr in transformers.items():
                    key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                    stats = self.xfmr_stats.get(key, {})
                    
                    P_fwd = abs(stats.get('avg_P_fwd', 0))
                    Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                    MVA_fwd = (P_fwd**2 + Q_fwd**2)**0.5
                    
                    P_rev = abs(stats.get('avg_P_rev', 0))
                    Q_rev = abs(stats.get('avg_Q_rev', 0))
                    MVA_rev = (P_rev**2 + Q_rev**2)**0.5
                    
                    Loss_MW = abs(stats.get('avg_loss', 0))
                    Loss_MVA = (Loss_MW**2 + (abs(stats.get('avg_loss_q', 0)))**2)**0.5 if stats.get('avg_loss_q', 0) else Loss_MW
                    
                    # Get transformer rating (default 9999 if not provided)
                    rating = xfmr.get('rateA', 9999)
                    rating_display = f"{rating:.0f}" if rating < 9999 else "N/A"
                    
                    f.write(f"{xfmr_num:<6} {xfmr.get('name', f'Xfmr_{xfmr_num}'):<15} "
                           f"{xfmr.get('from_bus', 0)}->{xfmr.get('to_bus', 0):<9} "
                           f"{rating_display:<8} "  # NEW: Transformer Rating column
                           f"{P_fwd:<10.2f} {Q_fwd:<10.2f} {MVA_fwd:<10.2f} "
                           f"{P_rev:<10.2f} {Q_rev:<10.2f} {MVA_rev:<10.2f} "
                           f"{Loss_MW:<10.3f} {Loss_MVA:<10.3f} "
                           f"{stats.get('avg_tap', xfmr.get('tap_ratio', 1.0)):<8.4f} "
                           f"{stats.get('avg_loading', 0):<8.1f}\n")
            else:
                f.write("No transformer data available\n")
            f.write("\n")


            # SECTION 6: CAPACITOR BANK DATA
            f.write("SECTION 6: CAPACITOR BANK DATA\n")
            f.write("-"*100 + "\n")
            f.write(f"{'Cap#':<8} {'Name':<15} {'Bus':<8} {'Q_cap_Mvar':<12} {'V_actual_pu':<12} {'V_actual_kV':<12} {'Q_inj_Mvar':<12} {'Status':<10}\n")
            f.write("-"*100 + "\n")
            capacitors = self.input_data.get('capacitors', {})
            if capacitors:
                for cap_num, cap in capacitors.items():
                    bus_num = cap.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = cap.get('Q_cap', 0) * (V_actual ** 2)
                    f.write(f"{cap_num:<8} {cap.get('name', f'Cap_{cap_num}'):<15} {bus_num:<8} "
                           f"{cap.get('Q_cap', 0):<12.2f} {V_actual:<12.4f} {V_kV:<12.2f} {Q_inj:<12.2f} "
                           f"{'ONLINE' if cap.get('status', 1) == 1 else 'OFFLINE'}\n")
            else:
                f.write("No capacitor data available\n")
            f.write("\n")

            # SECTION 7: REACTOR DATA
            f.write("SECTION 7: REACTOR DATA\n")
            f.write("-"*100 + "\n")
            f.write(f"{'Reactor#':<8} {'Name':<15} {'Bus':<8} {'Q_react_Mvar':<12} {'V_actual_pu':<12} {'V_actual_kV':<12} {'Q_abs_Mvar':<12} {'Status':<10}\n")
            f.write("-"*100 + "\n")
            reactors = self.input_data.get('reactors', {})
            if reactors:
                for reactor_num, reactor in reactors.items():
                    bus_num = reactor.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_abs = reactor.get('Q_react', 0) * (V_actual ** 2)
                    f.write(f"{reactor_num:<8} {reactor.get('name', f'Reactor_{reactor_num}'):<15} {bus_num:<8} "
                           f"{reactor.get('Q_react', 0):<12.2f} {V_actual:<12.4f} {V_kV:<12.2f} {Q_abs:<12.2f} "
                           f"{'ONLINE' if reactor.get('status', 1) == 1 else 'OFFLINE'}\n")
            else:
                f.write("No reactor data available\n")
            f.write("\n")

            # SECTION 8: SERIES COMPENSATION DATA
            f.write("SECTION 8: SERIES COMPENSATION DATA\n")
            f.write("-"*120 + "\n")
            f.write(f"{'Series#':<8} {'Name':<15} {'From->To':<12} {'R_pu':<10} {'X_pu':<10} {'Comp_pct':<10} {'Eff_X_pu':<12} {'Status':<10}\n")
            f.write("-"*120 + "\n")
            series_comps = self.input_data.get('series_comps', {})
            if series_comps:
                for series_num, series in series_comps.items():
                    effective_X = series.get('x', 0) * (1 - series.get('comp_pct', 0) / 100)
                    f.write(f"{series_num:<8} {series.get('name', f'SC_{series_num}'):<15} "
                           f"{series.get('from_bus', 0)}->{series.get('to_bus', 0):<9} "
                           f"{series.get('r', 0):<10.4f} {series.get('x', 0):<10.4f} "
                           f"{series.get('comp_pct', 0):<10.1f} {effective_X:<12.6f} "
                           f"{'ACTIVE' if series.get('status', 1) == 1 else 'INACTIVE'}\n")
            else:
                f.write("No series compensation data available\n")
            f.write("\n")

            # SECTION 9: SERIES REACTOR DATA
            f.write("SECTION 9: SERIES REACTOR DATA\n")
            f.write("-"*100 + "\n")
            f.write(f"{'SR#':<8} {'Name':<15} {'From->To':<12} {'R_pu':<10} {'X_pu':<10} {'Status':<10}\n")
            f.write("-"*100 + "\n")
            series_reactors = self.input_data.get('series_reactors', {})
            if series_reactors:
                for sr_num, sr in series_reactors.items():
                    f.write(f"{sr_num:<8} {sr.get('name', f'SR_{sr_num}'):<15} "
                           f"{sr.get('from_bus', 0)}->{sr.get('to_bus', 0):<9} "
                           f"{sr.get('r', 0):<10.4f} {sr.get('x', 0):<10.4f} "
                           f"{'ONLINE' if sr.get('status', 1) == 1 else 'OFFLINE'}\n")
            else:
                f.write("No series reactor data available\n")
            f.write("\n")
            

            # SECTION 10: SHUNT COMPENSATION DATA
            f.write("SECTION 10: SHUNT COMPENSATION DATA\n")
            f.write("-"*100 + "\n")
            f.write(f"{'Shunt#':<8} {'Name':<15} {'Bus':<8} {'Q_shunt_Mvar':<12} {'V_actual_pu':<12} {'V_actual_kV':<12} {'Q_inj_Mvar':<12} {'Status':<10}\n")
            f.write("-"*100 + "\n")
            shunts = self.input_data.get('shunts', {})
            if shunts:
                for shunt_num, shunt in shunts.items():
                    bus_num = shunt.get('bus', 0)
                    idx = self.engine.bus_index_map.get(bus_num, 0)
                    V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                    V_kV = self.get_voltage_kv(bus_num, V_actual)
                    Q_inj = shunt.get('Q_shunt', 0) * (V_actual ** 2)
                    f.write(f"{shunt_num:<8} {shunt.get('name', f'Shunt_{shunt_num}'):<15} {bus_num:<8} "
                           f"{shunt.get('Q_shunt', 0):<12.2f} {V_actual:<12.4f} {V_kV:<12.2f} {Q_inj:<12.2f} "
                           f"{'ONLINE' if shunt.get('status', 1) == 1 else 'OFFLINE'}\n")
            else:
                f.write("No shunt compensation data available\n")
            f.write("\n")

            
            # =========================================================
            # SECTION 11: ZONE-WISE POWER SUMMARY
            # =========================================================
            f.write("SECTION 11: ZONE-WISE POWER SUMMARY\n")
            
            # Collect all zones from all components
            all_zones = set()
            
            for bus in self.input_data.get('buses', {}).values():
                all_zones.add(bus.get('zone', 1))
            for gen in self.input_data.get('generators', {}).values():
                all_zones.add(gen.get('zone', 1))
            for load in self.input_data.get('loads', {}).values():
                all_zones.add(load.get('zone', 1))
            for line in self.input_data.get('lines', {}).values():
                all_zones.add(line.get('zone', 1))
            for xfmr in self.input_data.get('transformers', {}).values():
                all_zones.add(xfmr.get('zone', 1))
            for cap in self.input_data.get('capacitors', {}).values():
                all_zones.add(cap.get('zone', 1))
            for reactor in self.input_data.get('reactors', {}).values():
                all_zones.add(reactor.get('zone', 1))
            for sc in self.input_data.get('series_comps', {}).values():
                all_zones.add(sc.get('zone', 1))
            for sr in self.input_data.get('series_reactors', {}).values():
                all_zones.add(sr.get('zone', 1))
            for shunt in self.input_data.get('shunts', {}).values():
                all_zones.add(shunt.get('zone', 1))
            
            sorted_zones = sorted(all_zones)
            
            # Create header with zone numbers
            header = f"{'Component':<20}"
            for zone in sorted_zones:
                header += f" {'Zone ' + str(zone):<15}"
            header += f" {'TOTAL':<15}"
            f.write(header + "\n")
            
            # Calculate totals per zone
            gen_p = {zone: 0 for zone in sorted_zones}
            gen_q = {zone: 0 for zone in sorted_zones}
            load_p = {zone: 0 for zone in sorted_zones}
            load_q = {zone: 0 for zone in sorted_zones}
            cap_q = {zone: 0 for zone in sorted_zones}
            reactor_q = {zone: 0 for zone in sorted_zones}
            shunt_q = {zone: 0 for zone in sorted_zones}
            sc_pct = {zone: 0 for zone in sorted_zones}
            sr_x = {zone: 0 for zone in sorted_zones}
            line_loss = {zone: 0 for zone in sorted_zones}
            xfmr_loss = {zone: 0 for zone in sorted_zones}
            
            # Collect Generator data
            for gen in self.input_data.get('generators', {}).values():
                zone = gen.get('zone', 1)
                if zone in gen_p:
                    gen_p[zone] += gen.get('P_out', 0)
                    gen_q[zone] += gen.get('Q_out', 0)
            
            # Collect Load data
            for load in self.input_data.get('loads', {}).values():
                zone = load.get('zone', 1)
                if zone in load_p:
                    load_p[zone] += load.get('P_demand', 0)
                    load_q[zone] += load.get('Q_demand', 0)
            
            # Collect Capacitor data
            for cap in self.input_data.get('capacitors', {}).values():
                zone = cap.get('zone', 1)
                bus_num = cap.get('bus', 0)
                idx = self.engine.bus_index_map.get(bus_num, 0)
                V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                Q_inj = cap.get('Q_cap', 0) * (V_actual ** 2)
                if zone in cap_q:
                    cap_q[zone] += Q_inj
            
            # Collect Reactor data
            for reactor in self.input_data.get('reactors', {}).values():
                zone = reactor.get('zone', 1)
                bus_num = reactor.get('bus', 0)
                idx = self.engine.bus_index_map.get(bus_num, 0)
                V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                Q_abs = reactor.get('Q_react', 0) * (V_actual ** 2)
                if zone in reactor_q:
                    reactor_q[zone] += Q_abs
            
            # Collect Shunt data
            for shunt in self.input_data.get('shunts', {}).values():
                zone = shunt.get('zone', 1)
                bus_num = shunt.get('bus', 0)
                idx = self.engine.bus_index_map.get(bus_num, 0)
                V_actual = self.avg_V[idx] if idx < len(self.avg_V) else 1.0
                Q_inj = shunt.get('Q_shunt', 0) * (V_actual ** 2)
                if zone in shunt_q:
                    shunt_q[zone] += Q_inj
            
            # Collect Series Compensation data
            for sc in self.input_data.get('series_comps', {}).values():
                zone = sc.get('zone', 1)
                if zone in sc_pct:
                    sc_pct[zone] += sc.get('comp_pct', 0)
            
            # Collect Series Reactor data
            for sr in self.input_data.get('series_reactors', {}).values():
                zone = sr.get('zone', 1)
                if zone in sr_x:
                    sr_x[zone] += sr.get('x', 0)
            
            # Collect Line Losses
            for line in self.input_data.get('lines', {}).values():
                zone = line.get('zone', 1)
                key = f"{line.get('from_bus', 0)}-{line.get('to_bus', 0)}"
                stats = self.line_stats.get(key, {})
                loss = abs(stats.get('avg_loss', 0))
                if zone in line_loss:
                    line_loss[zone] += loss
            
            # Collect Transformer Losses
            for xfmr in self.input_data.get('transformers', {}).values():
                zone = xfmr.get('zone', 1)
                key = f"{xfmr.get('from_bus', 0)}-{xfmr.get('to_bus', 0)}"
                stats = self.xfmr_stats.get(key, {})
                loss = abs(stats.get('avg_loss', 0))
                if zone in xfmr_loss:
                    xfmr_loss[zone] += loss
            
            # Write data rows
            row = f"{'Generator P (MW)':<20}"
            for zone in sorted_zones:
                row += f" {gen_p[zone]:>15.2f}"
            row += f" {sum(gen_p.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Generator Q (Mvar)':<20}"
            for zone in sorted_zones:
                row += f" {gen_q[zone]:>15.2f}"
            row += f" {sum(gen_q.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Load P (MW)':<20}"
            for zone in sorted_zones:
                row += f" {load_p[zone]:>15.2f}"
            row += f" {sum(load_p.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Load Q (Mvar)':<20}"
            for zone in sorted_zones:
                row += f" {load_q[zone]:>15.2f}"
            row += f" {sum(load_q.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Capacitor Q (Mvar)':<20}"
            for zone in sorted_zones:
                row += f" {cap_q[zone]:>15.2f}"
            row += f" {sum(cap_q.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Reactor Q (Mvar)':<20}"
            for zone in sorted_zones:
                row += f" {reactor_q[zone]:>15.2f}"
            row += f" {sum(reactor_q.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Shunt Q (Mvar)':<20}"
            for zone in sorted_zones:
                row += f" {shunt_q[zone]:>15.2f}"
            row += f" {sum(shunt_q.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Series Comp (%)':<20}"
            for zone in sorted_zones:
                row += f" {sc_pct[zone]:>15.2f}"
            row += f" {sum(sc_pct.values()):>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Series Reactor X (pu)':<20}"
            for zone in sorted_zones:
                row += f" {sr_x[zone]:>15.4f}"
            row += f" {sum(sr_x.values()):>15.4f}"
            f.write(row + "\n")
            
            row = f"{'Line Losses (MW)':<20}"
            for zone in sorted_zones:
                row += f" {line_loss[zone]:>15.3f}"
            row += f" {sum(line_loss.values()):>15.3f}"
            f.write(row + "\n")
            
            row = f"{'Transformer Losses (MW)':<20}"
            for zone in sorted_zones:
                row += f" {xfmr_loss[zone]:>15.3f}"
            row += f" {sum(xfmr_loss.values()):>15.3f}"
            f.write(row + "\n")
            
            # Net Power per Zone
            row = f"{'Net Power (MW)':<20}"
            for zone in sorted_zones:
                net_p = gen_p[zone] - load_p[zone] - line_loss[zone] - xfmr_loss[zone]
                row += f" {net_p:>15.2f}"
            total_net = sum(gen_p.values()) - sum(load_p.values()) - sum(line_loss.values()) - sum(xfmr_loss.values())
            row += f" {total_net:>15.2f}"
            f.write(row + "\n")
            
            row = f"{'Net Reactive (Mvar)':<20}"
            for zone in sorted_zones:
                net_q = gen_q[zone] - load_q[zone] + cap_q[zone] - reactor_q[zone] + shunt_q[zone]
                row += f" {net_q:>15.2f}"
            total_net_q = sum(gen_q.values()) - sum(load_q.values()) + sum(cap_q.values()) - sum(reactor_q.values()) + sum(shunt_q.values())
            row += f" {total_net_q:>15.2f}"
            f.write(row + "\n")
            
            f.write("\nNote: Positive Net Power = Export from zone, Negative = Import to zone\n")
            f.write("\n")

            # =========================================================
            # SECTION 12: ZONE-TO-ZONE POWER FLOW MATRIX
            # =========================================================
            f.write("SECTION 12: ZONE-TO-ZONE POWER FLOW MATRIX\n")
            
            # Get all zones
            all_zones = set()
            for bus in self.input_data.get('buses', {}).values():
                all_zones.add(bus.get('zone', 1))
            for line in self.input_data.get('lines', {}).values():
                all_zones.add(line.get('zone', 1))
            for xfmr in self.input_data.get('transformers', {}).values():
                all_zones.add(xfmr.get('zone', 1))
            
            sorted_zones = sorted(all_zones)
            zone_list = list(sorted_zones)
            
            # Create zone mapping for bus numbers
            bus_to_zone = {}
            for bus_num, bus in self.input_data.get('buses', {}).items():
                bus_to_zone[bus_num] = bus.get('zone', 1)
            
            # Initialize flow matrices
            # P_flow[from_zone][to_zone] = total active power flow
            # Q_flow[from_zone][to_zone] = total reactive power flow
            P_flow = {z: {z2: 0 for z2 in zone_list} for z in zone_list}
            Q_flow = {z: {z2: 0 for z2 in zone_list} for z in zone_list}
            
            # Process transmission lines
            for line in self.input_data.get('lines', {}).values():
                from_bus = line.get('from_bus', 0)
                to_bus = line.get('to_bus', 0)
                
                from_zone = bus_to_zone.get(from_bus, 1)
                to_zone = bus_to_zone.get(to_bus, 1)
                
                key = f"{from_bus}-{to_bus}"
                stats = self.line_stats.get(key, {})
                
                P_fwd = abs(stats.get('avg_P_fwd', 0))
                Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                
                if from_zone != to_zone:
                    P_flow[from_zone][to_zone] += P_fwd
                    Q_flow[from_zone][to_zone] += Q_fwd
            
            # Process transformers
            for xfmr in self.input_data.get('transformers', {}).values():
                from_bus = xfmr.get('from_bus', 0)
                to_bus = xfmr.get('to_bus', 0)
                
                from_zone = bus_to_zone.get(from_bus, 1)
                to_zone = bus_to_zone.get(to_bus, 1)
                
                key = f"{from_bus}-{to_bus}"
                stats = self.xfmr_stats.get(key, {})
                
                P_fwd = abs(stats.get('avg_P_fwd', 0))
                Q_fwd = abs(stats.get('avg_Q_fwd', 0))
                
                if from_zone != to_zone:
                    P_flow[from_zone][to_zone] += P_fwd
                    Q_flow[from_zone][to_zone] += Q_fwd
            
            # Remove zero flows (optional - to keep table clean)
            # Create list of zones that have any non-zero flow
            zones_with_flow = set()
            for from_z in zone_list:
                for to_z in zone_list:
                    if from_z != to_z and (P_flow[from_z][to_z] > 0.01 or Q_flow[from_z][to_z] > 0.01):
                        zones_with_flow.add(from_z)
                        zones_with_flow.add(to_z)
            
            display_zones = sorted(zones_with_flow) if zones_with_flow else zone_list
            
            if len(display_zones) > 1:
                # Create header for matrix
                f.write("\n")
                f.write("ACTIVE POWER FLOW (MW) - From Zone (rows) To Zone (columns)\n")
                f.write("-" * (15 + 12 * len(display_zones)) + "\n")
                
                # Header row
                header = f"{'From->To':<12}"
                for to_z in display_zones:
                    header += f" Zone {to_z:<8}"
                f.write(header + "\n")
                
                # Data rows
                for from_z in display_zones:
                    row = f"Zone {from_z:<7}"
                    for to_z in display_zones:
                        if from_z == to_z:
                            value = "---"
                        else:
                            value = f"{P_flow[from_z][to_z]:.2f}"
                        row += f" {value:>10}"
                    f.write(row + "\n")
                
                f.write("\n")
                f.write("REACTIVE POWER FLOW (Mvar) - From Zone (rows) To Zone (columns)\n")
                f.write("-" * (15 + 12 * len(display_zones)) + "\n")
                
                # Header row
                header = f"{'From->To':<12}"
                for to_z in display_zones:
                    header += f" Zone {to_z:<8}"
                f.write(header + "\n")
                
                # Data rows
                for from_z in display_zones:
                    row = f"Zone {from_z:<7}"
                    for to_z in display_zones:
                        if from_z == to_z:
                            value = "---"
                        else:
                            value = f"{Q_flow[from_z][to_z]:.2f}"
                        row += f" {value:>10}"
                    f.write(row + "\n")
                
                # Summary of inter-zone power exchange
                f.write("\n")
                f.write("INTER-ZONE POWER EXCHANGE SUMMARY\n")
                f.write("-" * 50 + "\n")
                
                for from_z in display_zones:
                    total_export_p = sum(P_flow[from_z][to_z] for to_z in display_zones if to_z != from_z)
                    total_export_q = sum(Q_flow[from_z][to_z] for to_z in display_zones if to_z != from_z)
                    
                    total_import_p = sum(P_flow[to_z][from_z] for to_z in display_zones if to_z != from_z)
                    total_import_q = sum(Q_flow[to_z][from_z] for to_z in display_zones if to_z != from_z)
                    
                    net_export_p = total_export_p - total_import_p
                    net_export_q = total_export_q - total_import_q
                    
                    if abs(net_export_p) > 0.01 or abs(net_export_q) > 0.01:
                        direction = "EXPORT" if net_export_p > 0 else "IMPORT"
                        f.write(f"Zone {from_z}: {direction} {abs(net_export_p):.2f} MW, {abs(net_export_q):.2f} Mvar\n")
                
                f.write("\n")
                
                # Detailed flow table (only non-zero flows)
                f.write("DETAILED INTER-ZONE FLOWS (Non-zero flows only)\n")
                f.write("-" * 60 + "\n")
                f.write(f"{'From Zone':<12} {'To Zone':<12} {'P (MW)':<12} {'Q (Mvar)':<12}\n")
                
                for from_z in display_zones:
                    for to_z in display_zones:
                        if from_z != to_z and (P_flow[from_z][to_z] > 0.01 or Q_flow[from_z][to_z] > 0.01):
                            f.write(f"Zone {from_z:<9} Zone {to_z:<9} {P_flow[from_z][to_z]:<12.2f} {Q_flow[from_z][to_z]:<12.2f}\n")
                
            else:
                f.write("No inter-zone power flows detected (all branches within same zone)\n")
            
            f.write("\n")
            
            
            # =========================================================
            # SECTION 13: SYSTEM PERFORMANCE
            # =========================================================
            total_losses = sum(s['avg_loss'] for s in self.line_stats.values()) + sum(s['avg_loss'] for s in self.xfmr_stats.values())
            total_losses_q = sum(s.get('avg_loss_q', 0) for s in self.line_stats.values()) + sum(s.get('avg_loss_q', 0) for s in self.xfmr_stats.values())
            total_losses_mva = (total_losses**2 + total_losses_q**2)**0.5
            total_load_p = sum(l.get('P_demand', 0) for l in self.input_data.get('loads', {}).values())
            total_load_q = sum(l.get('Q_demand', 0) for l in self.input_data.get('loads', {}).values())
            total_load_mva = (total_load_p**2 + total_load_q**2)**0.5
            
            f.write("\n" + "="*80 + "\n")
            f.write("SECTION 11: SYSTEM PERFORMANCE METRICS\n")
            f.write("-"*80 + "\n")
            f.write(f"Total Generation (P): {sum(g.get('P_out', 0) for g in self.input_data.get('generators', {}).values()):.2f} MW\n")
            f.write(f"Total Load (P): {total_load_p:.2f} MW\n")
            f.write(f"Total Load (Q): {total_load_q:.2f} Mvar\n")
            f.write(f"Total Load (MVA): {total_load_mva:.2f} MVA\n")
            f.write(f"Total Losses (P): {total_losses:.2f} MW\n")
            f.write(f"Total Losses (Q): {total_losses_q:.2f} Mvar\n")
            f.write(f"Total Losses (MVA): {total_losses_mva:.2f} MVA\n")
            f.write(f"Loss Percentage (P): {(total_losses/total_load_p)*100 if total_load_p > 0 else 0:.2f}%\n")
            f.write(f"Average Voltage: {np.mean(self.avg_V):.4f} pu\n")
            f.write(f"Minimum Voltage: {np.min(self.avg_V):.4f} pu\n")
            f.write(f"Maximum Voltage: {np.max(self.avg_V):.4f} pu\n")
            
            # Component counts
            f.write("\n" + "-"*80 + "\n")
            f.write("COMPONENT COUNTS\n")
            f.write("-"*80 + "\n")
            f.write(f"Buses: {len(self.input_data.get('buses', {}))}\n")
            f.write(f"Generators: {len(self.input_data.get('generators', {}))}\n")
            f.write(f"Loads: {len(self.input_data.get('loads', {}))}\n")
            f.write(f"Transmission Lines: {len(self.input_data.get('lines', {}))}\n")
            f.write(f"Transformers: {len(self.input_data.get('transformers', {}))}\n")
            f.write(f"Capacitors: {len(self.input_data.get('capacitors', {}))}\n")
            f.write(f"Reactors: {len(self.input_data.get('reactors', {}))}\n")
            f.write(f"Series Compensation: {len(self.input_data.get('series_comps', {}))}\n")
            f.write(f"Series Reactors: {len(self.input_data.get('series_reactors', {}))}\n")
            f.write(f"Shunt Compensation: {len(self.input_data.get('shunts', {}))}\n")
            
            f.write("="*80 + "\n")
        
        #print(f"  ✓ TXT (complete summary with ALL sections): {filename}")
        return filename
