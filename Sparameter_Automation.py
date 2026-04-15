import json
import math
import time
import traceback
from typing import Dict, List, Tuple
import clr as pynet
from ClassFile import S2P
import copy
import datetime
import os
import uuid
from pathlib import Path
import sys


# Initialise RapidMS API
script_dir = Path(__file__).parent.absolute()
dll_path = script_dir / "RapidMS.ClientAPI.dll"
sys.path.append(str(script_dir))
pynet.AddReference(str(dll_path))
import RapidMS
rlp = RapidMS.ClientAPI.RapidMSClientAPI()

# Configuration Constants
PORT = 8733
KEY = ""


def load_gold_standard(filename: str) -> Dict[float, Tuple[float, float]]:
    """
    Loads and parses the S2P gold standard reference file.
    
    Args:
        filename: Path to the .s2p gold standard file
        
    Returns:
        Dictionary mapping frequency to S2P objects containing S-parameters
    """
    data_dict = {}
    frequency_unit = None
    format_type = None

    with open(filename, 'r') as file:
        for line in file:
            line = line.strip()

            # Skip empty lines and comments
            if not line or line.startswith('!'):
                continue

            # Parse header information
            if line.startswith('#'):
                parts = line.split()
                if len(parts) >= 4:
                    frequency_unit = parts[1].upper()
                    format_type = parts[3].upper()
                    continue
                
            # Parse measurement data
            values = list(map(float, line.split()))
            if len(values) >= 9:
                freq = float(values[0])
                s2p_object = S2P()
                s2p_object.frequency_unit = frequency_unit
                s2p_object.format_type = format_type
                s2p_object.set_all_params(
                    s11=(values[1], values[2]),
                    s21=(values[3], values[4]),
                    s12=(values[5], values[6]),
                    s22=(values[7], values[8]),
                )
                data_dict[freq] = s2p_object
    file.close()
    return data_dict


def load_measurement_data(measurement_data: str) -> Dict[float, Tuple[float, float]]:
    """
    Converts JSON measurement results to S2P object dictionary.
    
    Args:
        measurement_data: JSON string containing measurement results
        
    Returns:
        Dictionary mapping frequency to S2P objects containing measured S-parameters
    """
    data_dict = {}
    frequency_unit = 'GHz'
    format_type = 'dB' 
    results_json = json.loads(measurement_data)

    for idx in range(len(results_json["Frequency_GHz"])):
        freq = float(results_json["Frequency_GHz"][idx])
        s2p_object = S2P()
        s2p_object.frequency_unit = frequency_unit
        s2p_object.format_type = format_type

        s2p_object.set_all_params(
            s11=(float(results_json["S11_Mag_dB"][idx]), float(results_json["S11_Phase_deg"][idx])),
            s21=(float(results_json["S21_Mag_dB"][idx]), float(results_json["S21_Phase_deg"][idx])),
            s12=(float(results_json["S12_Mag_dB"][idx]), float(results_json["S12_Phase_deg"][idx])),
            s22=(float(results_json["S22_Mag_dB"][idx]), float(results_json["S22_Phase_deg"][idx]))
        )
        
        data_dict[freq] = s2p_object

    return data_dict


def align_initial_phase(gs: List[float], meas: List[float]) -> List[float]:
    """
    Aligns the initial phase of measured data to gold standard reference.
    
    Args:
        gs: Gold standard phase data
        meas: Measured phase data
        
    Returns:
        Phase-aligned measured data
    """
    aligned = meas.copy()
    diff = gs[0] - aligned[0]
    
    while diff > 180:
        aligned = [phase + 360 for phase in aligned]
        diff = gs[0] - aligned[0]
    while diff < -180:
        aligned = [phase - 360 for phase in aligned]
        diff = gs[0] - aligned[0]
    
    return aligned


def custom_unwrap(phase_list: List[float]) -> List[float]:
    """
    Custom phase unwrapping algorithm to handle discontinuities.
    
    Args:
        phase_list: List of phase values in degrees
        
    Returns:
        Unwrapped phase values
    """
    unwrapped_phase_list = [phase_list[0]]
    pad = 0

    for prev, curr in zip(phase_list, phase_list[1:]):
        diff = curr - prev
        if diff > 180:
            pad -= 360
        elif diff < -180:
            pad += 360
        unwrapped_phase_list.append(curr + pad)
    
    return unwrapped_phase_list


def align_unwrapped_phase(g_unwrapped: List[float], m_unwrapped: List[float]) -> List[float]:
    """
    Aligns unwrapped phase data between gold standard and measured data.
    
    Args:
        g_unwrapped: Gold standard unwrapped phase data
        m_unwrapped: Measured unwrapped phase data
        
    Returns:
        Aligned unwrapped measured phase data
    """
    offset = g_unwrapped[0] - m_unwrapped[0]
    return [phase + offset for phase in m_unwrapped]


def process_both(gold: Dict, meas: Dict) -> Tuple[Dict, Dict]:
    """
    Processes and aligns phase data between gold standard and measured S-parameters.
    
    Args:
        gold: Gold standard S-parameter data
        meas: Measured S-parameter data
        
    Returns:
        Tuple of processed (gold_standard, measured) dictionaries with aligned phases
    """
    freqs = sorted(set(gold) & set(meas))
    g_ph = {'s11': [], 's21': [], 's12': [], 's22': []}
    m_ph = {'s11': [], 's21': [], 's12': [], 's22': []}

    # Extract phase data for each S-parameter
    for freq in freqs:
        g_o, m_o = gold[freq], meas[freq]
        for phase in g_ph:
            g_ph[phase].append(getattr(g_o, phase)[1])
            m_ph[phase].append(getattr(m_o, phase)[1])
    
    # Align and unwrap phases
    m_wr = {phase: align_initial_phase(g_ph[phase], m_ph[phase]) for phase in g_ph}
    g_unw = {phase: custom_unwrap(g_ph[phase]) for phase in g_ph}
    m_unw = {phase: custom_unwrap(m_wr[phase]) for phase in m_ph}

    def build(src, phases):
        """Rebuilds S-parameter objects with processed phase data."""
        out = {}
        for i, f in enumerate(freqs):
            obj = copy.deepcopy(src[f])
            for p in phases:
                mag = getattr(src[f], p)[0]
                setattr(obj, p, (mag, phases[p][i]))
            out[f] = obj
        return out
    
    return build(gold, g_unw), build(meas, m_unw)


def compare(gold_filt: Dict, meas_filt: Dict, mag_tolerance: float, phase_tolerance: float) -> Tuple[str, Dict, float, float]:
    """
    Compares measured S-parameters against gold standard with specified tolerances.
    
    Args:
        gold_filt: Filtered gold standard data
        meas_filt: Filtered measured data
        mag_tolerance: Magnitude tolerance in dB
        phase_tolerance: Phase tolerance in degrees
        
    Returns:
        Tuple containing (verdict, error_dict, mag_tolerance, phase_tolerance)
    """
    params = ["s11", "s21", "s12", "s22"]
    params = ["s11", "s21", "s12", "s22"]
    test_verdict = "PASSED"
    error_dict = {}

    for f in sorted(set(gold_filt) & set(meas_filt)):
        g_o, m_o = gold_filt[f], meas_filt[f]
        
        for param in params:
            gm, gp = getattr(g_o, param)
            mm, mp = getattr(m_o, param)
            
            # Adjust tolerance for low signal levels
            if mm < -40:
                mag_tolerance = 0.6
                phase_tolerance = 5

            mag_diff = abs(gm - mm)
            
            # Normalise phase 
            phase1_norm = gp % 360
            phase2_norm = mp % 360
            phase_diff = abs(phase1_norm - phase2_norm)
            if phase_diff > 180:
                phase_diff = 360 - phase_diff

            if mag_diff > mag_tolerance or phase_diff > phase_tolerance:
                test_verdict = "FAILED"
                error_dict.setdefault(f, {})[param] = (round(mag_diff, 2), round(phase_diff, 2))

    return test_verdict, error_dict, mag_tolerance, phase_tolerance


def start_hardware_and_bias_device(source_cal: str, vector_cal: str, mode: int, frequency_list: List[float], power: float, input_voltage: float,
                                   output_voltage: float, aux1_voltage: float, aux2_voltage: float, bandwidth: int, bias_control: bool, pul_period: float, pul_width: float,
                                   meas_window: float, meas_delay: float, avg: int):
    """
    Initialises hardware and configures measurement parameters.
    
    Args:
        source_cal: Source calibration file path
        vector_cal: Vector calibration file path
        mode: Measurement mode (0=CW, 1=Pulsed, 2=Modulated)
        frequency_list: List of frequencies to measure
        power: Power level setting
        voltage: DC bias voltage
        bandwidth : IF Bandwidth (integer)
        bias_control: Enable/disable bias control
    """
    # Load calibration files and configure measurement
    rlp.SetMeasurementMode(mode, 1)  # 0 for Impedance measurement
    rlp.LoadSourceCalibration(source_cal)
    rlp.LoadVectorCalibration(vector_cal)
    if type(frequency_list) == list:
        rlp.SetSparameterFrequencyList(frequency_list)
    else:
        rlp.SetSparameterFrequencyList([frequency_list])
    if mode == 0:
        rlp.SetIFBandwidth(bandwidth)
    elif mode == 1:
        rlp.SetPulsedSettings(pul_period, pul_width, meas_window, meas_delay, avg)
    rlp.StartHardware()
    rlp.SetPowerLevel(power)
    
    if bias_control:
        if input_voltage != 0.0:
            rlp.SetDCBias(0, input_voltage) 
            rlp.TurnDCOnOff(0, True)
        if output_voltage != 0.0:
            rlp.SetDCBias(1, output_voltage)
            rlp.TurnDCOnOff(1, True)
        if aux1_voltage != 0.0:
            rlp.SetDCBias(2, aux1_voltage) 
            rlp.TurnDCOnOff(0, True)
        if aux2_voltage != 0.0:
            rlp.SetDCBias(3, aux2_voltage)
            rlp.TurnDCOnOff(1, True)
        
        # Safety check: ensure minimum current draw
        current = rlp.GetDCCurrent(1)
        if current < 30:
            print('Current below 30mA, stopping test for safety...')
            rlp.StopHardware()
            exit()


def measurement(input_voltage: float, output_voltage: float, aux1_voltage: float, aux2_voltgae: float, bias_control: bool) -> str:
    """
    Executes S-parameter measurement and returns results.
    
    Args:
        bias_control: Whether bias control is enabled
        
    Returns:
        JSON string containing measurement results
    """
    rlp.DoMeasurement("")
    
    # Wait for measurement completion
    print("Measuring...")
    while not rlp.IsMeasurementCompleted():
        pass
    
    result = rlp.GetSingleMeasuredData()
    print("Measurement complete")
    
    if bias_control:
        if aux2_voltgae != 0.0:
            rlp.TurnDCOnOff(3, False)
        if aux1_voltage != 0.0:
            rlp.TurnDCOnOff(2, False)
        if output_voltage != 0.0:
            rlp.TurnDCOnOff(1, False)
        if input_voltage != 0.0:
            rlp.TurnDCOnOff(0, False)
    rlp.StopHardware()

    return result


def output_measured_data(meas: Dict) -> str:
    """
    Outputs measured data to S2P format file.
    
    Args:
        meas: Dictionary containing measured S-parameter data
        
    Returns:
        Absolute path to the created output file
    """
    time_str = time.strftime("%Y%m%d-%H%M%S")
    filename = f"Meas_{time_str}.s2p"

    with open(filename, 'w') as f:
        f.write("!s2p data\n")
        f.write("# GHz S dB R 50\n\n")
        
        for freq in meas:
            s_params = meas[freq]
            f.write(f"{freq} {s_params.s11[0]} {s_params.s11[1]} "
                   f"{s_params.s21[0]} {s_params.s21[1]} "
                   f"{s_params.s12[0]} {s_params.s12[1]} "
                   f"{s_params.s22[0]} {s_params.s22[1]}\n")

    print(f"Output file created: {filename}")
    return os.path.abspath(filename)


def run_s_param_test(vals: List) -> Dict:
    """
    Main function to execute complete S-parameter test sequence.
    
    Args:
        vals: List containing test configuration parameters
        
    Returns:
        Dictionary containing test results and metadata
    """
    # System information for test traceability
    SYSTEM_INFO = {
        "jira": "RVT-633",
        "commit": "2fa5d9b9", 
        "hardware": "RapidVT",
        "sn": "1130",
        "fw": "RapidTS-0.1"
    }
    
    # Unpack test parameters
    (freq_list, pwr, t1, source_cal, vector_cal, init, gs, IP, mode, 
     inp_v, out_v, aux1_v, aux2_v, max_gain, stop_at_compression, 
     power_sweep, imp_sweep, bias_control, if_bw, mod_bw, mod_wf, pulsed_period, pulse_width, meas_window, meas_delay, pulsed_avg) = vals
    
    test_passed = True
    start_time = time.time()
    
    # Attempt connection to measurement system
    connection_result = rlp.Connect(IP, PORT, KEY)
    
    if connection_result:
        print("Connected")
        
        try:
            # Execute measurement sequence
            start_hardware_and_bias_device(source_cal, vector_cal, mode, freq_list, pwr, inp_v, out_v, aux1_v, aux2_v, if_bw, bias_control, pulsed_period,
                                           pulse_width, meas_window, meas_delay, pulsed_avg)
            measurement_data = measurement(inp_v, out_v, aux1_v, aux2_v, bias_control)
            execution_time = round(time.time() - start_time)
            
            # Process and analyse results
            measured = load_measurement_data(measurement_data)
            file_path = output_measured_data(measured)
            gold_standard = load_gold_standard(gs)
            filtered_gs, filtered_meas = process_both(gold_standard, measured)
            verdict, err_dict, mag_tol, phase_tol = compare(filtered_gs, filtered_meas, 0.2, 3)

            print(f"Test {verdict}")
            print(f"Execution time: {execution_time}s")
            
            if verdict == 'FAILED':
                test_passed = False
                print(f"Error details: {err_dict}")

        except Exception as ex:
            print(f"Test execution error: {ex}")
            print(traceback.format_exc())
            test_passed = False
            verdict = "FAILED"
            err_dict = {}
            file_path = ""
            mag_tol = 0.0
            phase_tol = 0.0

        finally:
            rlp.Disconnect()
            print("Disconnected")
            
    else:
        print("Connection failed!")
        test_passed = False
        verdict = "FAILED"
        err_dict = {}
        file_path = ""
        mag_tol = 0.0
        phase_tol = 0.0
    
    # Compile test results
    results = {
        "test_passed": test_passed,
        "verdict": verdict,
        "err_dict": err_dict,
        "file": file_path,
        "mag_tol": mag_tol,
        "phase_tol": phase_tol,
        "system_info": SYSTEM_INFO,
        "test_config": {
            "mode": mode,
            "mode_name": "CW" if mode == 0 else "Pulsed" if mode == 1 else "Modulated", 
            "source_cal": source_cal,
            "vector_cal": vector_cal,
            "if_bandwidth": f"{if_bw} kHz" if if_bw else "N/A",
            "frequency_list": freq_list
        }
    }
    
    return results


if __name__ == "__main__":
    pass