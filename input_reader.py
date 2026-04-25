"""
Input Reader - Handles various input formats
Can be modified to support different input file types without changing engine
"""

import importlib.util
import os

class InputReader:
    """
    Reads power system data from various formats
    Currently supports: Python file with data lists
    """
    
    @staticmethod
    def read_from_python(file_path):
        """
        Read data from Python format input file
        Format includes: BUS_DATA, GENERATOR_DATA, LOAD_DATA, LINE_DATA, 
                         TRANSFORMER_DATA, CAPACITOR_DATA, REACTOR_DATA,
                         SERIES_COMP_DATA, SERIES_REACTOR_DATA, SHUNT_DATA
        
        Extended with: AREA, ZONE, OWNER fields
        """
        #print(f"\n[READING] {file_path}")
        
        try:
            # Load Python module
            case_name = os.path.splitext(os.path.basename(file_path))[0]
            spec = importlib.util.spec_from_file_location(case_name, file_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            # Extract data
            bus_data = {}
            if hasattr(module, 'BUS_DATA'):
                for row in module.BUS_DATA:
                    bus_num = row[0]
                    bus_data[bus_num] = {
                        'bus_num': bus_num,
                        'name': row[1] if len(row) > 1 else f"Bus_{bus_num}",
                        'type': row[2] if len(row) > 2 else 1,
                        'base_kV': row[3] if len(row) > 3 else 132.0,
                        'V_init': row[4] if len(row) > 4 else 1.0,
                        'angle_init': row[5] if len(row) > 5 else 0.0,
                        'area': row[8] if len(row) > 8 else 1,  # AREA field
                        'zone': row[9] if len(row) > 9 else 1,  # ZONE field
                        'owner': row[10] if len(row) > 10 else 1,  # OWNER field
                        'shunt_G': row[6] if len(row) > 6 else 0,
                        'shunt_B': row[7] if len(row) > 7 else 0
                    }
            
            # Generator data
            generator_data = {}
            if hasattr(module, 'GENERATOR_DATA'):
                for row in module.GENERATOR_DATA:
                    generator_data[row[0]] = {
                        'num': row[0], 'name': row[1], 'type': row[2],
                        'bus': row[3], 'area': row[10] if len(row) > 10 else 1,
                        'zone': row[11] if len(row) > 11 else 1,
                        'P_out': row[4], 'Q_out': row[5], 'V_set': row[6],
                        'Qmin': row[7], 'Qmax': row[8], 'status': row[9]
                    }
            
            # Load data
            load_data = {}
            if hasattr(module, 'LOAD_DATA'):
                for row in module.LOAD_DATA:
                    load_data[row[0]] = {
                        'num': row[0], 'name': row[1], 'bus': row[2],
                        'area': row[6] if len(row) > 6 else 1,
                        'zone': row[7] if len(row) > 7 else 1,
                        'P_demand': row[3], 'Q_demand': row[4], 'model': row[5]
                    }

            # Line data with length calculation
            line_data = {}
            if hasattr(module, 'LINE_DATA'):
                for row in module.LINE_DATA:
                    length_km = row[4] if len(row) > 4 else 0
                    R_per_km = row[5] if len(row) > 5 else 0
                    X_per_km = row[6] if len(row) > 6 else 0
                    B_per_km = row[7] if len(row) > 7 else 0
                    
                    # Calculate total values from per-km and length
                    R_total = R_per_km * length_km
                    X_total = X_per_km * length_km
                    B_total = B_per_km * length_km
                    
                    line_data[row[0]] = {
                        'num': row[0], 'name': row[1],
                        'from_bus': row[2], 'to_bus': row[3],
                        'length_km': length_km,
                        'R_per_km': R_per_km, 'X_per_km': X_per_km, 'B_per_km': B_per_km,
                        'r': R_total, 'x': X_total, 'b': B_total,  # These go to engine
                        'rateA': row[8] if len(row) > 8 else 0,
                        'status': row[9] if len(row) > 9 else 1,
                        'area': row[10] if len(row) > 10 else 1,
                        'zone': row[11] if len(row) > 11 else 1,
                        'owner': row[12] if len(row) > 12 else 1
                    }

            # Transformer data (NO length - direct R, X)
            transformer_data = {}
            if hasattr(module, 'TRANSFORMER_DATA'):
                for row in module.TRANSFORMER_DATA:
                    transformer_data[row[0]] = {
                        'num': row[0], 'name': row[1],
                        'from_bus': row[2], 'to_bus': row[3],
                        'r': row[4] if len(row) > 4 else 0,
                        'x': row[5] if len(row) > 5 else 0,
                        'tap_ratio': row[6] if len(row) > 6 else 1.0,
                        'rateA': row[7] if len(row) > 7 else 9999,  # NEW: Transformer rating
                        'phase_shift': row[8] if len(row) > 8 else 0,
                        'min_tap': row[9] if len(row) > 9 else 0.9,
                        'max_tap': row[10] if len(row) > 10 else 1.1,
                        'step_size': row[11] if len(row) > 11 else 0.01,
                        'status': row[12] if len(row) > 12 else 1,
                        'area': row[13] if len(row) > 13 else 1,
                        'zone': row[14] if len(row) > 14 else 1,
                        'owner': row[15] if len(row) > 15 else 1
                    }


            # Capacitor data
            capacitor_data = {}
            if hasattr(module, 'CAPACITOR_DATA'):
                for row in module.CAPACITOR_DATA:
                    capacitor_data[row[0]] = {
                        'num': row[0], 'name': row[1], 'bus': row[2],
                        'area': row[5] if len(row) > 5 else 1,
                        'zone': row[6] if len(row) > 6 else 1,
                        'Q_cap': row[3], 'status': row[4]
                    }
            
            # Reactor data
            reactor_data = {}
            if hasattr(module, 'REACTOR_DATA'):
                for row in module.REACTOR_DATA:
                    reactor_data[row[0]] = {
                        'num': row[0], 'name': row[1], 'bus': row[2],
                        'area': row[5] if len(row) > 5 else 1,
                        'zone': row[6] if len(row) > 6 else 1,
                        'Q_react': row[3], 'status': row[4]
                    }
            
            # Series Compensation data
            series_comp_data = {}
            if hasattr(module, 'SERIES_COMP_DATA'):
                for row in module.SERIES_COMP_DATA:
                    series_comp_data[row[0]] = {
                        'num': row[0], 'name': row[1],
                        'from_bus': row[2], 'to_bus': row[3],
                        'area': row[8] if len(row) > 8 else 1,
                        'zone': row[9] if len(row) > 9 else 1,
                        'owner': row[10] if len(row) > 10 else 1,
                        'r': row[4], 'x': row[5], 'comp_pct': row[6], 'status': row[7]
                    }
            
            # Series Reactor data
            series_reactor_data = {}
            if hasattr(module, 'SERIES_REACTOR_DATA'):
                for row in module.SERIES_REACTOR_DATA:
                    series_reactor_data[row[0]] = {
                        'num': row[0], 'name': row[1],
                        'from_bus': row[2], 'to_bus': row[3],
                        'area': row[7] if len(row) > 7 else 1,
                        'zone': row[8] if len(row) > 8 else 1,
                        'owner': row[9] if len(row) > 9 else 1,
                        'r': row[4], 'x': row[5], 'status': row[6]
                    }
            
            # Shunt data
            shunt_data = {}
            if hasattr(module, 'SHUNT_DATA'):
                for row in module.SHUNT_DATA:
                    shunt_data[row[0]] = {
                        'num': row[0], 'name': row[1], 'bus': row[2],
                        'area': row[5] if len(row) > 5 else 1,
                        'zone': row[6] if len(row) > 6 else 1,
                        'Q_shunt': row[3], 'status': row[4]
                    }
            
            baseMVA = getattr(module, 'BASE_MVA', 100.0)
            
            print(
                f"✓ Read {len(bus_data)} buses, "
                f"✓ Read {len(generator_data)} generators, "
                f"✓ Read {len(load_data)} loads, "
                f"✓ Read {len(line_data)} lines, "
                f"✓ Read {len(transformer_data)} transformers"
            )            
            return {
                'baseMVA': baseMVA,
                'buses': bus_data,
                'generators': generator_data,
                'loads': load_data,
                'lines': line_data,
                'transformers': transformer_data,
                'capacitors': capacitor_data,
                'reactors': reactor_data,
                'series_comps': series_comp_data,
                'series_reactors': series_reactor_data,
                'shunts': shunt_data
            }
            
        except Exception as e:
            print(f"  ✗ Error: {e}")
            return None

    @staticmethod
    def convert_to_engine_format(input_data):
        """
        Convert input data to engine format
        Extracts bus and branch data needed for power flow
        """
        bus_data = {}
        for bus_num, bus in input_data['buses'].items():
            bus_data[bus_num] = {
                'V_init': bus['V_init'],
                'angle_init': bus['angle_init'],
                'base_kV': bus['base_kV'],
                'type': bus['type'],
                'name': bus['name'],
                'area': bus.get('area', 1),
                'zone': bus.get('zone', 1),
                'owner': bus.get('owner', 1)
            }
        
        branch_data = []
        
        # Add lines
        for line in input_data['lines'].values():
            if line['status'] == 1:
                branch_data.append({
                    'type': 'line',
                    'num': line['num'],
                    'name': line['name'],
                    'from_bus': line['from_bus'],
                    'to_bus': line['to_bus'],
                    'r': line['r'],
                    'x': line['x'],
                    'b': line['b'],
                    'ratio': 0,
                    'rateA': line.get('rateA', 0),
                    'area': line.get('area', 1),
                    'zone': line.get('zone', 1),
                    'owner': line.get('owner', 1)
                })
        
        # Add transformers - FIXED
        for xfmr in input_data['transformers'].values():
            if xfmr['status'] == 1:
                branch_data.append({
                    'type': 'transformer',
                    'num': xfmr['num'],
                    'name': xfmr['name'],
                    'from_bus': xfmr['from_bus'],
                    'to_bus': xfmr['to_bus'],
                    'r': xfmr['r'],
                    'x': xfmr['x'],
                    'b': 0,
                    'ratio': xfmr['tap_ratio'],
                    'rateA': xfmr.get('rateA', 9999),  # FIXED: Use input value
                    'min_tap': xfmr.get('min_tap', 0.9),
                    'max_tap': xfmr.get('max_tap', 1.1),
                    'step_size': xfmr.get('step_size', 0.01),
                    'area': xfmr.get('area', 1),
                    'zone': xfmr.get('zone', 1),
                    'owner': xfmr.get('owner', 1)
                })
        
        return bus_data, branch_data, input_data

        