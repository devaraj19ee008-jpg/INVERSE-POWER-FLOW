"""
Inverse Power Flow Engine - Core Calculation Logic
This file should NEVER be modified when changing input/output formats
"""

import numpy as np

class PowerFlowEngine:
    """
    Core engine for inverse power flow calculations
    Independent of input/output formats
    """
    
    def __init__(self, baseMVA=100.0):
        self.baseMVA = baseMVA
        self.Ybus = None
        self.buses = None
        self.branches = None
        self.n_buses = 0
        self.bus_index_map = {}
        
    def initialize(self, bus_data, branch_data):
        """
        Initialize engine with bus and branch data
        
        bus_data: dict {bus_num: {'V_init': float, 'angle_init': float, 'base_kV': float, ...}}
        branch_data: list of dict with keys: from_bus, to_bus, r, x, b, ratio, rateA, type
        """
        self.buses = bus_data
        self.branches = branch_data
        self.n_buses = len(bus_data)
        self.build_index_map()
        self.build_admittance_matrix()
        
    def build_index_map(self):
        """Build mapping from bus number to index"""
        bus_numbers = sorted(self.buses.keys())
        self.bus_index_map = {bus_num: idx for idx, bus_num in enumerate(bus_numbers)}
    
    def build_admittance_matrix(self):
        """Build Ybus matrix"""
        self.Ybus = np.zeros((self.n_buses, self.n_buses), dtype=complex)
        
        for branch in self.branches:
            f_idx = self.bus_index_map[branch['from_bus']]
            t_idx = self.bus_index_map[branch['to_bus']]
            r, x, b, ratio = branch['r'], branch['x'], branch['b'], branch['ratio']
            
            z = complex(r, x)
            if z != 0:
                y = 1.0 / z
                if ratio != 0:  # Transformer
                    self.Ybus[f_idx][f_idx] += y / (ratio ** 2)
                    self.Ybus[t_idx][t_idx] += y
                    self.Ybus[f_idx][t_idx] -= y / ratio
                    self.Ybus[t_idx][f_idx] -= y / ratio
                else:  # Transmission line
                    self.Ybus[f_idx][f_idx] += y + complex(0, b)
                    self.Ybus[t_idx][t_idx] += y + complex(0, b)
                    self.Ybus[f_idx][t_idx] -= y
                    self.Ybus[t_idx][f_idx] -= y
        
        return self.Ybus
    
    def get_base_voltages(self):
        """Get base voltage magnitudes and angles"""
        V_mag = np.zeros(self.n_buses)
        V_angle = np.zeros(self.n_buses)
        for bus_num, bus in self.buses.items():
            idx = self.bus_index_map[bus_num]
            V_mag[idx] = bus.get('V_init', 1.0)
            V_angle[idx] = bus.get('angle_init', 0.0)
        return V_mag, V_angle
    
    def calculate_power(self, V_mag, V_angle_deg):
        """Calculate power injection from voltages"""
        V_angle_rad = V_angle_deg * np.pi / 180.0
        V_complex = V_mag * np.exp(1j * V_angle_rad)
        I_inj = self.Ybus @ V_complex
        S = V_complex * np.conj(I_inj)
        return S.real, S.imag
    
    def calculate_bidirectional_flows(self, V_mag, V_angle_deg):
        """Calculate bidirectional power flows on all branches"""
        V_angle_rad = V_angle_deg * np.pi / 180.0
        V_complex = V_mag * np.exp(1j * V_angle_rad)
        
        flows = []
        for branch in self.branches:
            f_idx = self.bus_index_map[branch['from_bus']]
            t_idx = self.bus_index_map[branch['to_bus']]
            Vf, Vt = V_complex[f_idx], V_complex[t_idx]
            
            z = complex(branch['r'], branch['x'])
            y = 1.0 / z if z != 0 else 0
            
            if branch['ratio'] != 0:  # Transformer
                ratio = branch['ratio']
                # Forward (from → to)
                If_fwd = (Vf / ratio - Vt) * y
                S_fwd = Vf * np.conj(If_fwd)
                # Reverse (to → from)
                If_rev = (Vt * ratio - Vf) * (y / (ratio ** 2))
                S_rev = Vt * np.conj(If_rev)
                losses = S_fwd + S_rev
                
                flows.append({
                    'type': 'transformer',
                    'num': branch.get('num'),
                    'name': branch.get('name'),
                    'from_bus': branch['from_bus'],
                    'to_bus': branch['to_bus'],
                    'P_fwd_MW': S_fwd.real * self.baseMVA,
                    'Q_fwd_Mvar': S_fwd.imag * self.baseMVA,
                    'P_rev_MW': S_rev.real * self.baseMVA,
                    'Q_rev_Mvar': S_rev.imag * self.baseMVA,
                    'losses_MW': losses.real * self.baseMVA,
                    #'loading_pct': abs(S_fwd) / branch.get('rateA', 9999) * 100 if branch.get('rateA', 0) > 0 else 0,
                    'loading_pct': (abs(S_fwd) * self.baseMVA) / branch.get('rateA', 9999) * 100 if branch.get('rateA', 0) > 0 else 0,
                    'tap_ratio': ratio
                })
            else:  # Transmission line
                y_shunt = complex(0, branch['b'])
                # Forward (from → to)
                Ift = (Vf - Vt) * y
                Ish_f = Vf * y_shunt
                S_fwd = Vf * np.conj(Ift + Ish_f)
                # Reverse (to → from)
                Itf = (Vt - Vf) * y
                Ish_t = Vt * y_shunt
                S_rev = Vt * np.conj(Itf + Ish_t)
                losses = S_fwd + S_rev
                
                flows.append({
                    'type': 'line',
                    'num': branch.get('num'),
                    'name': branch.get('name'),
                    'from_bus': branch['from_bus'],
                    'to_bus': branch['to_bus'],
                    'P_fwd_MW': S_fwd.real * self.baseMVA,
                    'Q_fwd_Mvar': S_fwd.imag * self.baseMVA,
                    'P_rev_MW': S_rev.real * self.baseMVA,
                    'Q_rev_Mvar': S_rev.imag * self.baseMVA,
                    'losses_MW': losses.real * self.baseMVA,
                    'loading_pct': (abs(S_fwd) * self.baseMVA) / branch.get('rateA', 9999) * 100 if branch.get('rateA', 0) > 0 else 0
                    # 'loading_pct': abs(S_fwd) / branch.get('rateA', 9999) * 100 if branch.get('rateA', 0) > 0 else 0
                })
        
        return flows
    
    def generate_sample(self, variation_strength=0.05):
        """Generate one power flow sample"""
        V_mag_base, V_angle_base = self.get_base_voltages()
        
        variation = np.random.uniform(-variation_strength, variation_strength, self.n_buses)
        V_mag = V_mag_base * (1 + variation)
        V_mag = np.clip(V_mag, 0.90, 1.1)
        
        angle_variation = np.random.uniform(-variation_strength * 30, variation_strength * 30, self.n_buses)
        V_angle = V_angle_base + angle_variation
        
        P_calc, Q_calc = self.calculate_power(V_mag, V_angle)
        branch_flows = self.calculate_bidirectional_flows(V_mag, V_angle)
        
        return {
            'V_mag': V_mag, 'V_angle': V_angle,
            'P_calc': P_calc, 'Q_calc': Q_calc,
            'branch_flows': branch_flows
        }
    
    def validate_sample(self, sample):
        """Validate sample"""
        if np.any(np.isnan(sample['P_calc'])) or np.any(np.isinf(sample['P_calc'])):
            return False
        if np.any(sample['V_mag'] < 0.85) or np.any(sample['V_mag'] > 1.15):
            return False
        return True
    
    def run_batch(self, n_samples=100, variation_strength=0.10):
        """Run batch of samples"""
        samples = []
        valid_count = 0
        
        for i in range(n_samples):
            sample = self.generate_sample(variation_strength=variation_strength)
            if self.validate_sample(sample):
                valid_count += 1
                samples.append(sample)
        
        return samples, valid_count