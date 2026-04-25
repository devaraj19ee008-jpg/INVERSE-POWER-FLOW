"""
Main Runner - Orchestrates input reading, engine execution, and report generation
Supports both interactive and command-line argument modes
"""

import os
import sys
import argparse
from datetime import datetime

from power_flow_engine import PowerFlowEngine
from input_reader import InputReader
from report_writer import ReportWriter

def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(
        description='Inverse Power Flow System - Run power flow analysis',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (asks for all inputs)
  python run_power_flow.py

  # Command line mode with all parameters
  python run_power_flow.py -i "E:\\invlfa\\eq\\converted_case300.py" -v 10 -n 200

  # Short form
  python run_power_flow.py -i case300.py -v 5 -n 100
        """
    )
    
    parser.add_argument('-i', '--input', 
                        type=str, 
                        help='Path to input Python file')
    parser.add_argument('-v', '--variation', 
                        #type=int,
                        type=float, 
                        #choices=[5, 10, 15],
                        help='Variation strength in percent (5, 10, or 15)')
    parser.add_argument('-n', '--samples', 
                        type=int, 
                        help='Number of samples to generate')
    
    return parser.parse_args()

def get_parameters_interactive():
    """Get parameters interactively from user"""
    print("\n" + "="*80)
    print("INTERACTIVE MODE")
    print("="*80)
    
    # Get input file
    print("\nEnter Python input file path:")
    print("(Supports AREA, ZONE, OWNER fields)")
    input_file = input("\nPath: ").strip().strip('"').strip("'")
    
    # Get variation strength
    print("\n" + "="*80)
    print("SIMULATION PARAMETERS")
    print("="*80)
    print("\nVariation Strength:")
    print("  [1] Small (5%)")
    print("  [2] Medium (10%) - Recommended")
    print("  [3] Large (15%)")
    
    choice = input("\nChoice (1-3) [default: 2]: ").strip()
    variation_map = {'1': 0.05, '2': 0.10, '3': 0.15}
    variation = variation_map.get(choice, 0.10)
    
    # Get number of samples
    n_samples = input("\nNumber of samples [default: 200]: ").strip()
    n_samples = int(n_samples) if n_samples else 200
    
    return input_file, variation, n_samples

def get_parameters_from_args(args):
    """Get parameters from command line arguments"""
    input_file = args.input
    variation = args.variation / 100.0 if args.variation else None
    n_samples = args.samples
    
    return input_file, variation, n_samples

def main():
    print("\n" + "="*80)
    print("INVERSE POWER FLOW SYSTEM - MODULAR ARCHITECTURE")
    print("Engine | Input Reader | Report Writer | Main Runner")
    print("="*80)
    
    # Parse command line arguments
    args = parse_arguments()
    
    # Determine mode: interactive or command-line
    if args.input and args.variation and args.samples:
        # Command-line mode - all parameters provided
        #print("\n" + "="*80)
        #print("COMMAND LINE MODE")
        #print("="*80)
        input_file, variation, n_samples = get_parameters_from_args(args)
        #print(f"  Input file: {input_file}")
        #print(f"  Variation: {variation*100:.0f}%")
        #print(f"  Samples: {n_samples}")
    else:
        # Interactive mode - ask for parameters
        input_file, variation, n_samples = get_parameters_interactive()
    
    # Validate input file
    if not os.path.isfile(input_file):
        print(f"\n[ERROR] File not found: {input_file}")
        return
    
    if not input_file.endswith('.py'):
        print(f"\n[ERROR] File must be .py format")
        return
    
    # Step 1: Read input data
    input_data = InputReader.read_from_python(input_file)
    if not input_data:
        return
    
    # Step 2: Convert to engine format
    bus_data, branch_data, full_data = InputReader.convert_to_engine_format(input_data)
    
    # Step 3: Initialize engine
    engine = PowerFlowEngine(baseMVA=input_data['baseMVA'])
    engine.initialize(bus_data, branch_data)
    
    # Step 4: Run batch simulation
    #print("\n" + "="*80)
    print("\n","RUNNING SIMULATION")
    #print("="*80)
    print(f"  Variation: {variation*100:.0f}% | Samples: {n_samples}")
    
    samples, valid_count = engine.run_batch(n_samples=n_samples, variation_strength=variation)
    
    print(f"\n  Valid samples: {valid_count}/{n_samples} ({100*valid_count/n_samples:.0f}%)")
    
    if len(samples) == 0:
        print("\n[ERROR] No valid samples.")
        return
    
    # # Step 5: Generate reports
    # print("\n" + "="*80)
    # print("GENERATING REPORTS")
    # print("="*80)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    case_name = os.path.splitext(os.path.basename(input_file))[0]
    base_name = f"{case_name}_var{int(variation*100)}_{timestamp}"
    
    # Get input file directory and create output folder with input file name
    input_dir = os.path.dirname(os.path.abspath(input_file))
    output_folder = os.path.join(input_dir, case_name)
    
    # Create the folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # ========== ADD THIS BLOCK IF U WANT TO SAVE SAMPLES FOR NN TRAINING ==========
    # Save samples for NN training (optional)
    # import pickle
    # samples_path = os.path.join(output_folder, f"{base_name}_SAMPLES.pkl")
    # with open(samples_path, 'wb') as f:
    #     pickle.dump(samples, f)
    # print(f"  💾 Samples saved for NN training: {samples_path}")
    # ========== END OF BLOCK ==========
    
    # Create full output paths inside the new folder
    csv_path = os.path.join(output_folder, f"{base_name}_ALL_DATA.csv")
    py_all_path = os.path.join(output_folder, f"{base_name}_ALL_DATA.py")
    py_output_path = os.path.join(output_folder, f"{base_name}_OUTPUTS_ONLY.py")
    txt_path = os.path.join(output_folder, f"{base_name}_SUMMARY.invout")
    #ieee_path = os.path.join(output_folder, f"{base_name}_IEEE.IEEE")

    reporter = ReportWriter(engine, full_data, samples, variation)
    
    # Generate 4 different output files
    csv_all = reporter.write_csv_all_data(csv_path)
    py_all = reporter.write_py_all_data(py_all_path)
    #py_output = reporter.write_py_output_only(py_output_path)
    txt_summary = reporter.write_txt_summary(txt_path)
    #ieee_report = reporter.write_ieee_report(ieee_path)



    
    print("\n" + "="*80)
    print("COMPLETE! FILES GENERATED:")
    print("="*80)
    print(f"   CSV (all data):      {csv_all}")
    print(f"   PY (all data):       {py_all}")
    #print(f"  PY (outputs only):   {py_output}")
    print(f"   INVERSE FORMATE :    {txt_summary}")
    #print(f"   IEEE (report):       {ieee_report}")
    print(f"\n  📁 All results saved in: {output_folder}")
    print("="*80)
    # print("\n✅ Features included:")
    # print("   • Bidirectional flows (from→to AND to→from)")
    # print("   • Voltage in pu AND kV")
    # print("   • Transformer tap ratio AND step number")
    # print("   • AREA, ZONE, OWNER fields")
    # print("   • Complete status outputs")
    # print("="*80)


if __name__ == "__main__":
    main()



"""
USAGE EXAMPLES:
# With 5% variation and 100 samples
python run_power_flow.py -i case300.py -v 5 -n 100

# With 10% variation and 500 samples
python run_power_flow.py -i "E:\invlfa\eq\converted_case300.py" -v 10 -n 500

# With 15% variation and default 200 samples
python run_power_flow.py -i my_system.py -v 15

# With default variation (10%) and custom samples
python run_power_flow.py -i my_system.py -n 300

# Short form (if you don't specify all, it falls back to interactive for missing ones)
python run_power_flow.py -i my_system.py
# This will ask for variation and samples interactively
"""