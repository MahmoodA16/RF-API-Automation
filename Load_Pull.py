"""
LoadPull.py

RapidMS load-pull automation + optional WLAN EVM (RFmx) processing.

This file is intentionally “all-in-one” for customer delivery:
- Connects to RapidMS via .NET Client API (pythonnet/clr).
- Runs CW / pulsed / modulated measurements.
- Optionally extracts B2 waveform IQ, computes WLAN EVM via RFmx,
  and plots an FFT spectrum annotated with EVM results.
- Loads gold-standard reference files and compares results.

"""

from __future__ import annotations
import itertools
import json
import math
import os
import socket
import sys
import time
import traceback
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import clr as pynet
import matplotlib.pyplot as plt
import nirfmxinstr
import nirfmxwlan
import numpy as np

from ClassFile import IdxClass


# RapidMS API bootstrap (pythonnet)

SCRIPT_DIR = Path(__file__).parent.absolute()
DLL_PATH = SCRIPT_DIR / "RapidMS.ClientAPI.dll"
sys.path.append(str(SCRIPT_DIR))
pynet.AddReference(str(DLL_PATH))

import RapidMS  

rlp = RapidMS.ClientAPI.RapidMSClientAPI()


# Configuration constants

PORT = 8733
KEY = ""



# Signal processing functions


def resample_iq(
    iq: np.ndarray,
    original_rate_mhz: float,
    target_rate_mhz: float,
    clear_dc: bool = False,
) -> np.ndarray:
    """
    Resample IQ data using FFT-domain cropping (matches a common C# approach).

    Args:
        iq: Complex IQ samples.
        original_rate_mhz: Original sample rate in MHz.
        target_rate_mhz: Target sample rate in MHz.
        clear_dc: If True, zero out DC and adjacent bins in the FFT domain.

    Returns:
        Resampled complex IQ array (complex64).
    """
    size = len(iq)
    new_size = int(round(size / original_rate_mhz * target_rate_mhz))
    skip = (size - new_size) // 2

    spectrum = np.fft.fft(iq)
    spectrum = np.fft.fftshift(spectrum)  # shift DC to center

    # Crop to new size (center portion)
    spectrum = spectrum[skip:skip + new_size]

    if clear_dc:
        mid = len(spectrum) // 2
        spectrum[mid] = 0
        spectrum[mid - 1] = 0
        spectrum[mid + 1] = 0

    spectrum = np.fft.ifftshift(spectrum)
    resampled = np.fft.ifft(spectrum)

    return resampled.astype(np.complex64)


def get_measured_ab_wave(harmonic: int, waveform_enum: int) -> Tuple[np.ndarray, np.ndarray]:
    """
    Retrieve AB wave data from RapidMS.

    waveform_enum mapping:
        A1 = 0
        A2 = 1
        B1 = 2
        B2 = 3

    Args:
        harmonic: 1, 2, or 3 (first/second/third harmonic).
        waveform_enum: Waveform enum value (0..3).

    Returns:
        (real_wave, imag_wave) as float64 NumPy arrays.

    Raises:
        ValueError: If RapidMS returns a failure status.
    """
    result = rlp.GetMeasuredABWave(int(harmonic), str(waveform_enum))

    success = result[0]
    if not success:
        raise ValueError("GetMeasuredABWave returned failure status")

    real_wave = np.asarray(result[1], dtype=np.float64)
    imag_wave = np.asarray(result[2], dtype=np.float64)

    return real_wave, imag_wave


def get_evm_from_iq(
    iq: np.ndarray,
    fs_hz: float,
    channel_bandwidth_hz: float = 160e6,
    freq_ghz: Optional[float] = None,
) -> dict:
    """
    Compute WLAN OFDM ModAcc EVM from IQ samples using RFmx WLAN.

    This function is configured to match the C# reference implementation:
    - Enable OFDM ModAcc measurement and traces
    - EVM unit in percentage
    - Configure channel bandwidth
    - Configure frequency band (2.4 GHz or 5 GHz)
    - Configure standard: 802.11ax
    - Analyze IQ waveform with dx = 1/fs

    Args:
        iq: Complex IQ samples (complex64/complex128 recommended).
        fs_hz: Sample rate in Hz.
        channel_bandwidth_hz: WLAN channel bandwidth in Hz (default 160e6).
        freq_ghz: Center frequency in GHz (used only for band selection).

    Returns:
        Dictionary containing EVM results or an error, plus success flag.
    """
    instr = nirfmxinstr.Session(resource_name="", option_string="AnalysisOnly=1")
    wlan = instr.get_wlan_signal_configuration()
    wlan.ofdmmodacc.configuration.set_all_traces_enabled("", True)
    wlan.ofdmmodacc.configuration.set_measurement_enabled("", True)
    wlan.ofdmmodacc.configuration.configure_evm_unit(
        "",
        nirfmxwlan.OfdmModAccEvmUnit.PERCENTAGE,
    )

    wlan.configure_channel_bandwidth("", channel_bandwidth_hz)

    # Frequency band selection
    if freq_ghz is not None and freq_ghz < 3.0:
        wlan.set_ofdm_frequency_band(
            "",
            nirfmxwlan.OfdmFrequencyBand.OFDM_FREQUENCY_BAND_2_4GHZ,
        )
    else:
        wlan.set_ofdm_frequency_band(
            "",
            nirfmxwlan.OfdmFrequencyBand.OFDM_FREQUENCY_BAND_5GHZ,
        )

    wlan.configure_standard("", nirfmxwlan.Standard.STANDARD_802_11_AX)

    # Analyse IQ waveform
    wlan.analyze_iq_1_waveform(
        selector_string="",
        result_name="",
        iq=iq,
        x0=0.0,
        dx=1.0 / fs_hz,
        reset=True,
    )

    try:
        # Fetch first, then get
        wlan.ofdmmodacc.results.fetch_composite_rms_evm("", 10.0)

        composite_rms_evm_mean, _ = wlan.ofdmmodacc.results.get_composite_rms_evm_mean("")
        composite_data_rms_evm_mean, _ = wlan.ofdmmodacc.results.get_composite_data_rms_evm_mean("")
        composite_pilot_rms_evm_mean, _ = wlan.ofdmmodacc.results.get_composite_pilot_rms_evm_mean("")

        evm_results = {
            "composite_rms_evm": composite_rms_evm_mean,
            "composite_data_rms_evm": composite_data_rms_evm_mean,
            "composite_pilot_rms_evm": composite_pilot_rms_evm_mean,
            "success": True,
        }
    except Exception as exc:
        evm_results = {"error": str(exc), "success": False}
    finally:
        instr.close()

    return evm_results


def fft_complex_to_dbm(wave: np.ndarray, freq: float, resis: float = 50.0):
    """
    Compute FFT magnitude (dB) from a complex waveform.

    Note: This function returns magnitude in dB (20*log10(|X|)).
    It is named dbm historically in this project, but it is not
    calibrated to dBm without additional scaling.

    Args:
        wave: Complex time-domain samples.
        freq: Sample rate (Hz).
        resis: Unused placeholder (kept for API compatibility).

    Returns:
        (freqs_hz, magnitude_db) arrays.
    """
    wave_size = wave.size
    if wave_size < 16:
        raise ValueError("Not enough data")

    window = np.hanning(wave_size)
    coherent_gain = np.sum(window) / wave_size
    xw = wave * window

    X = np.fft.fft(xw)
    freqs = np.fft.fftfreq(wave_size, d=1 / freq)

    # Shift so 0 Hz is centered
    X = np.fft.fftshift(X)
    freqs = np.fft.fftshift(freqs)

    magnitude = np.abs(X) / (wave_size * coherent_gain)
    magnitude[magnitude < 1e-20] = 1e-20

    power_db = 20 * np.log10(magnitude)
    return freqs, power_db


def plot_dbm_spectrum(
    freqs,
    dbm,
    title: str = "AB Wave Spectrum",
    filename: Optional[str] = None,
    evm_results: Optional[dict] = None,
):
    """
    Plot FFT spectrum and optionally annotate with EVM results.

    Args:
        freqs: Frequency axis in Hz.
        dbm: Magnitude in dB (see fft_complex_to_dbm note).
        title: Plot title.
        filename: If provided, save to file; otherwise show.
        evm_results: Dictionary returned by get_evm_from_iq().
    """
    plt.figure(figsize=(12, 7))
    plt.plot(freqs / 1e6, dbm)
    plt.xlabel("Frequency (MHz)")
    plt.ylabel("Magnitude (dB)")
    plt.title(title)
    plt.grid(True)

    plt.ylim(-130, np.max(dbm) + 10)
    y_max = int(np.ceil((np.max(dbm) + 10) / 10) * 10)
    plt.yticks(np.arange(-130, y_max + 1, 10))

    # Add EVM annotation
    if evm_results and evm_results.get("success"):
        evm_text = (
            "EVM Results:\n"
            f"  Composite Rms EVM Mean: {evm_results['composite_rms_evm']:.2f}\n"
            f"  Data Rms EVM Mean:      {evm_results['composite_data_rms_evm']:.2f}\n"
            f"  Pilot Rms EVM Mean:     {evm_results['composite_pilot_rms_evm']:.2f}"
        )
        props = dict(boxstyle="round", facecolor="wheat", alpha=0.8)
        plt.text(
            0.98,
            0.97,
            evm_text,
            transform=plt.gca().transAxes,
            fontsize=9,
            verticalalignment="top",
            horizontalalignment="right",
            bbox=props,
            family="monospace",
        )
    elif evm_results and not evm_results.get("success"):
        evm_text = f"EVM Error: {evm_results.get('error', 'Unknown')}"
        props = dict(boxstyle="round", facecolor="lightcoral", alpha=0.8)
        plt.text(
            0.98,
            0.97,
            evm_text,
            transform=plt.gca().transAxes,
            fontsize=9,
            verticalalignment="top",
            horizontalalignment="right",
            bbox=props,
        )

    plt.tight_layout()

    if filename:
        plt.savefig(filename, dpi=150)
        plt.close()
    else:
        plt.show()



# Reference data parsing and comparison helper functions

def load_gold_standard_files(
    filename: str,
) -> Dict[IdxClass, Tuple[List[str], List[float]]]:
    """
    Load and parse gold-standard load pull reference files.

    Supports file types:
        - .lpcwave / .lpc
        - .sat / .satwave
        - .lpwave / .lp
        - .lpd

    Args:
        filename: Path to the gold-standard reference file.

    Returns:
        Dictionary: IdxClass -> (headers, measurement_data)
    """
    result_dict: Dict[IdxClass, Tuple[List[str], List[float]]] = {}
    frequency = None
    headers: List[str] = []
    current_point = None
    current_gamma = None
    current_phase = None

    try:
        file_catering = 0
        data_section_started = False
        headers_found = False

        with open(filename, "r") as file:
            lines = file.readlines()
            if file.name.endswith((".lpcwave", ".lpc")):
                file_catering = 1
            if file.name.endswith((".sat", ".satwave")):
                file_catering = 2

        for line in lines:
            line = line.strip()

            # Extract frequency from header
            if line.startswith("! Frequency ="):
                freq_part = line.split("=")[1].strip()
                freq_value = freq_part.split()[0]
                frequency = float(freq_value)
                continue

            # Parse headers
            if not headers_found:
                if file_catering == 2:
                    if line.startswith("! TITLES:"):
                        headers = line.replace("! TITLES:", "").strip().split()
                        headers_found = True
                        data_section_started = True
                        continue
                else:
                    if not line.startswith("!") and line:
                        headers = line.split()
                        headers_found = True
                        data_section_started = True
                        continue

            if not headers_found and line.startswith("!"):
                continue

            if data_section_started and headers_found and line:
                data_values: List = []

                # Measurement point definition
                if line.startswith("#"):
                    data_values = line.split()
                    if len(data_values) >= 4:
                        current_point = float(data_values[1])
                        current_gamma = float(data_values[2])
                        current_phase = float(data_values[3])
                    continue

                if not line.startswith("!"):
                    data_values = line.split()

                if len(data_values) < 1:
                    continue

                try:
                    psource_val = None
                    target_gamma_1f0 = None
                    target_phase_1f0 = None

                    for index, header in enumerate(headers):
                        adjusted_index = index
                        if file_catering == 1:
                            adjusted_index = index - 3

                        if 0 <= adjusted_index < len(data_values):
                            if "Psource" in header:
                                psource_val = float(data_values[adjusted_index])
                            if "TargetGamma_1F0" in header:
                                target_gamma_1f0 = float(data_values[adjusted_index])
                            if "TargetPhase_1F0[deg]" in header:
                                target_phase_1f0 = float(data_values[adjusted_index])
                            if "OutputEff[%]" in header:
                                if data_values[adjusted_index] in ["âˆž", "∞"]:
                                    data_values[adjusted_index] = float("inf")
                                elif data_values[adjusted_index] in ["-âˆž", "-∞"]:
                                    data_values[adjusted_index] = float("-inf")

                    data_values = string_to_float_handling(data_values)
                    idx_obj = IdxClass(frequency, psource_val, target_gamma_1f0, target_phase_1f0)

                    if file_catering == 1:
                        complete_row_data = [
                            float(current_point),
                            float(current_gamma),
                            float(current_phase),
                        ] + data_values
                        row_tuple = (headers.copy(), complete_row_data)
                    else:
                        row_tuple = (headers.copy(), data_values.copy())

                    if idx_obj not in result_dict:
                        result_dict[idx_obj] = row_tuple
                    else:
                        print(f"Warning: Duplicate key {idx_obj} found in gold standard")
                        exit()

                except (ValueError, IndexError) as exc:
                    print(f"Error processing line: {line}")
                    print(f"Error details: {exc}")
                    continue

    except FileNotFoundError:
        print(f"Gold standard file not found: {filename}")
        return {}
    except Exception as exc:
        print(f"Error reading gold standard file: {exc}")
        return {}

    return result_dict


def create_key_objects_concise(
    freq: float,
    pwr: List[float],
    t1: List[Tuple[float, float]],
) -> List[IdxClass]:
    """
    Create IdxClass key objects for measurement configuration.

    Args:
        freq: Measurement frequency.
        pwr: Power level(s) (single value or list).
        t1: List of (real, imag) tuples for impedance targets.

    Returns:
        List of IdxClass objects representing measurement points.
    """
    key_list: List[IdxClass] = []

    if isinstance(pwr, (int, float)):
        for t in t1:
            t_gamma, t_phase = get_gamma_phase(t[0], t[1])
            key_list.append(IdxClass(freq, pwr, t_gamma, t_phase))
    else:
        for power, t in itertools.product(pwr, t1):
            t_gamma, t_phase = get_gamma_phase(t[0], t[1])
            key_list.append(IdxClass(freq, power, t_gamma, t_phase))

    return key_list


def string_to_float_handling(data: List) -> List:
    """
    Convert list entries to float when possible.

    Args:
        data: List of values.

    Returns:
        Same list, with numeric strings converted to float.
    """
    for index, value in enumerate(data):
        try:
            data[index] = float(value)
        except (ValueError, TypeError):
            continue
    return data


def read_meas_to_dict(
    meas_results: str,
    target_real1: float,
    target_imaginary1: float,
    data_dict: Optional[Dict] = None,
) -> Dict[IdxClass, Tuple[List[str], List[Optional[float]]]]:
    """
    Convert JSON measurement results to dictionary format matching gold standard.

    Args:
        meas_results: JSON string containing measurement data.
        target_real1: Real part of impedance target.
        target_imaginary1: Imaginary part of impedance target.
        data_dict: Existing dictionary to append to (optional).

    Returns:
        Dictionary with IdxClass keys and (headers, data) tuple values.
    """
    if not data_dict:
        data_dict = {}

    data_values: List = []
    headers: List[str] = []
    target_gamma1, target_phase1 = get_gamma_phase(target_real1, target_imaginary1)

    results_json = json.loads(meas_results)

    for header in results_json:
        if "Frequency" in header:
            frequency = float(results_json[header])
        if "Psource" in header:
            psource = float(results_json[header])
        if "OutputEff[%]" in header:
            if results_json[header] in ("âˆž", "∞"):
                results_json[header] = float("inf")

        # Special handling for ACPR
        if header == "ACPR":
            try:
                acpr_data = json.loads(results_json[header])
                for offset_key, offset_data in acpr_data.items():
                    offset_data = json.loads(offset_data)
                    if isinstance(offset_data, dict):
                        if offset_key.startswith("Offset-"):
                            offset_num = offset_key.replace("Offset-", "")
                            formatted_offset = f"Offset_neg_{offset_num}"
                        elif offset_key.startswith("Offset") and not offset_key.startswith("Offset-"):
                            offset_num = offset_key.replace("Offset", "")
                            formatted_offset = f"Offset_{offset_num}"
                        else:
                            formatted_offset = offset_key

                        if "Power_dBm" in offset_data:
                            headers.append(f"{formatted_offset}_Power[dBm]")
                            data_values.append(offset_data["Power_dBm"])

                        if "Power_dBc" in offset_data:
                            headers.append(f"{formatted_offset}_Power[dBc]")
                            data_values.append(offset_data["Power_dBc"])

            except (json.JSONDecodeError, TypeError):
                headers.append(header)
                data_values.append(results_json[header])
        else:
            headers.append(header)
            data_values.append(results_json[header])

    idx_obj = IdxClass(frequency, psource, target_gamma1, target_phase1)
    data_values = string_to_float_handling(data_values)
    row_tuple = (headers.copy(), data_values.copy())

    if idx_obj not in data_dict:
        data_dict[idx_obj] = row_tuple
    else:
        print(f"Error, key : {idx_obj} already in dictionary")
        exit()

    return data_dict


def compare(
    gs_dict: Dict,
    meas_dict: Dict,
    compare_list: List[Tuple[str, float]],
) -> Tuple[str, Dict]:
    """
    Compare measured results against gold standard with specified tolerances.

    Args:
        gs_dict: Gold standard data dictionary.
        meas_dict: Measured data dictionary.
        compare_list: List of (parameter_name, tolerance) tuples to compare.

    Returns:
        (test_verdict, error_dictionary)
    """
    test_verdict = "PASSED"
    error_dict: Dict = {}

    for key, value_gs in gs_dict.items():
        if key not in meas_dict:
            continue

        meas_headers, meas_data = meas_dict[key]
        gs_headers, gs_data = value_gs
        key_errors = []

        for target_info in compare_list:
            if isinstance(target_info, tuple):
                target, tolerance = target_info
            else:
                target = target_info
                tolerance = 0

            try:
                index_gs = gs_headers.index(target)
                data_gs = gs_data[index_gs]
            except ValueError:
                print(f"Error: '{target}' not found in gold standard headers")
                return None, None
            except IndexError:
                print(f"Error: Index {index_gs} out of range in gold standard data")
                return None, None

            try:
                index_meas = meas_headers.index(target)
                data_meas = meas_data[index_meas]
            except ValueError:
                print(f"Error: '{target}' not found in measured data headers")
                return None, None
            except IndexError:
                print(f"Error: Index {index_meas} out of range in measured data")
                return None, None

            if data_gs != data_meas:
                if isinstance(data_gs, (int, float)) and isinstance(data_meas, (int, float)):
                    if target in ("PhiLWaves@F0[deg]", "PhiLoad@F0[deg]"):
                        phase1_norm = data_gs % 360
                        phase2_norm = data_meas % 360
                        phase_diff = abs(phase1_norm - phase2_norm)
                        if phase_diff > 180:
                            phase_diff = 360 - phase_diff
                        diff = phase_diff
                    else:
                        diff = abs(data_gs - data_meas)
                else:
                    diff = 1
            else:
                diff = 0

            if diff > tolerance:
                test_verdict = "FAILED"
                key_errors.append((target, round(diff, 2), tolerance))

        if key_errors:
            error_dict[key] = key_errors

    return test_verdict, error_dict



# Measurement flow

def measure_one(
    gs: Dict,
    freq: float,
    pwr_list: List[float],
    t1: List[Tuple[float, float]],
    source_cal: str,
    vector_cal: str,
    init: str,
    mode: int,
    inp_v: float,
    out_v: float,
    aux1v: float,
    aux2v: float,
    ispowersweep: bool,
    isimpedance_sweep: bool,
    maxgain: float,
    compression: bool,
    bias_control: bool,
    if_bandwidth: Optional[int],
    mod_bandwidth: Optional[int],
    mod_waveform: Optional[str],
    pul_period: Optional[float],
    pul_width: Optional[float],
    meas_window: Optional[float],
    meas_delay: Optional[float],
    pul_avg: Optional[int],
) -> Dict:
    """
    Execute a load-pull measurement sequence for specified test conditions.

    Args:
        gs: Gold standard reference data.
        freq: Measurement frequency (GHz in RapidMS context).
        pwr_list: Power level configuration (single or sweep specification).
        t1: Impedance target points [(real, imag), ...].
        source_cal: Source calibration file path.
        vector_cal: Vector calibration file path.
        init: Initialization file path.
        mode: 0=CW, 1=Pulsed, 2=Modulated.
        inp_v/out_v/aux1v/aux2v: Bias voltages.
        ispowersweep: Enable power sweep.
        isimpedance_sweep: Enable impedance sweep.
        maxgain: Maximum gain setting.
        compression: Compression sweep parameter.
        bias_control: Enable bias control.
        if_bandwidth: IF bandwidth (CW mode).
        mod_bandwidth: Modulated bandwidth (Modulated mode).
        mod_waveform: Modulated waveform file.
        pul_period/pul_width/meas_window/meas_delay/pul_avg: Pulsed settings.

    Returns:
        Dictionary containing measured data in gold-standard-like structure.
    """
    issweep = isimpedance_sweep or ispowersweep

    all_keys = create_key_objects_concise(freq, pwr_list, t1)
    keys_to_measure = [key for key in all_keys if key in gs]

    if not keys_to_measure:
        print("No measurement points match gold standard reference")
        return {}

    print(f"Measuring {len(keys_to_measure)} points from gold standard")

    # Configure measurement system
    rlp.SetMeasurementMode(mode, 0)  # 0 for impedance measurement
    rlp.LoadSourceCalibration(source_cal)
    rlp.LoadVectorCalibration(vector_cal)
    rlp.LoadInitializationFile(init)
    rlp.SetCenterFrequency(freq)

    if mode == 0 and if_bandwidth:
        rlp.SetIFBandwidth(if_bandwidth)
    elif mode == 1:
        rlp.SetPulsedSettings(pul_period, pul_width, meas_window, meas_delay, pul_avg)
    elif mode == 2 and mod_bandwidth:
        rlp.SetModulatedBandwidth(mod_bandwidth)

    rlp.StartHardware()

    if mode == 2 and mod_waveform:
        rlp.SetModulatedWaveform(mod_waveform)

    rlp.SetMeasurementType(issweep)
    rlp.SetLoadSourcepull(1, True, False, 10, 40)

    # Configure measurement based on sweep type
    if not issweep:
        rlp.SetLoadTargetPoint(1, t1[0][0], t1[0][1])
        rlp.SetPowerLevel(pwr_list)
    else:
        if isimpedance_sweep:
            real_list = [ri[0] for ri in t1]
            imaginary_list = [ri[1] for ri in t1]
            rlp.SetSweepTargetPoints(True, real_list, imaginary_list)
        else:
            rlp.SetSweepTargetPoints(True, [t1[0][0]], [t1[0][1]])

        if ispowersweep:
            rlp.SetPowerSweep(True, pwr_list[0], pwr_list[1], pwr_list[2])
            if compression:
                rlp.SetPowerSweepCompression(True, False, float(compression))
            else:
                rlp.SetPowerSweepCompression(False, False, 0.0)
        else:
            rlp.SetPowerSweep(False, pwr_list, pwr_list, 0)
            rlp.SetPowerLevel(pwr_list)

    # Configure bias control if enabled
    if bias_control:
        if inp_v != 0.0:
            rlp.SetDCBias(0, inp_v)
            rlp.TurnDCOnOff(0, True)
        if out_v != 0.0:
            rlp.SetDCBias(1, out_v)
            rlp.TurnDCOnOff(1, True)
        if aux1v != 0.0:
            rlp.SetDCBias(2, inp_v)
            rlp.TurnDCOnOff(0, True)
        if aux2v != 0.0:
            rlp.SetDCBias(3, out_v)
            rlp.TurnDCOnOff(1, True)

        current = rlp.GetDCCurrent(1)
        if current < 30:
            print("Current below 30mA, stopping for safety...")
            rlp.StopHardware()
            exit()

    # Execute measurement
    rlp.DoMeasurement("")
    while not rlp.IsMeasurementCompleted():
        time.sleep(0.1)

    # Modulated mode: extract B2 IQ, compute EVM, plot spectrum
    if mode == 2 and mod_bandwidth:
        try:
            # B2 waveform enum = 3
            real_wave, imag_wave = get_measured_ab_wave(harmonic=1, waveform_enum=3)
            iq = real_wave.astype(np.float32) + 1j * imag_wave.astype(np.float32)

            print(f"IQ length: {len(iq)}")
            print(f"mod_bandwidth: {mod_bandwidth}")

            original_rate_mhz = float(mod_bandwidth)

            # Use sampled data for EVM
            fs_hz = original_rate_mhz * 1e6
            channel_bw = 20e6

            print("Computing EVM...")
            evm_results = get_evm_from_iq(iq.astype(np.complex64), fs_hz, channel_bw, freq)

            if evm_results.get("success"):
                print(f"\n{'=' * 40}")
                print(f"Composite Rms EVM Mean: {evm_results['composite_rms_evm']:.2f}")
                print(f"Data Rms EVM Mean:      {evm_results['composite_data_rms_evm']:.2f}")
                print(f"Pilot Rms EVM Mean:     {evm_results['composite_pilot_rms_evm']:.2f}")
                print(f"{'=' * 40}\n")
            else:
                print(f"EVM computation failed: {evm_results.get('error')}")

            freqs, spec_dbm = fft_complex_to_dbm(iq, fs_hz)
            filename_plot = f"Spectrum_{freq}GHz_Harm1.png"
            plot_dbm_spectrum(
                freqs,
                spec_dbm,
                f"B2 Wave Spectrum @ {freq} GHz",
                filename=filename_plot,
                evm_results=evm_results,
            )
            print(f"Spectrum plot saved to {filename_plot}")

        except Exception as exc:
            print(f"Failed to extract/plot AB Wave: {exc}")
            traceback.print_exc()

    # Process measurement results
    if not issweep:
        result_json = rlp.GetSingleMeasuredData()
        measured_dict = read_meas_to_dict(result_json, t1[0][0], t1[0][1], None)
    else:
        result_json = rlp.GetSweepData()
        measured_dict = None

        if ispowersweep and isimpedance_sweep:
            results_per_impedance = len(result_json) // len(real_list)
            result_index = 0
            for target_real, target_imaginary in zip(real_list, imaginary_list):
                for _ in range(results_per_impedance):
                    if result_index < len(result_json):
                        measured_dict = read_meas_to_dict(
                            result_json[result_index],
                            target_real,
                            target_imaginary,
                            measured_dict,
                        )
                        result_index += 1
        elif ispowersweep:
            target_real = t1[0][0]
            target_imaginary = t1[0][1]
            for result in result_json:
                measured_dict = read_meas_to_dict(
                    result,
                    target_real,
                    target_imaginary,
                    measured_dict,
                )
        else:
            for result, target_real, target_imaginary in zip(result_json, real_list, imaginary_list):
                measured_dict = read_meas_to_dict(
                    result,
                    target_real,
                    target_imaginary,
                    measured_dict,
                )

    print("Measurement sequence complete")

    # Clean up hardware state
    if bias_control:
        if aux2v != 0.0:
            rlp.TurnDCOnOff(3, False)
        if aux1v != 0.0:
            rlp.TurnDCOnOff(2, False)
        if out_v != 0.0:
            rlp.TurnDCOnOff(1, False)
        if inp_v != 0.0:
            rlp.TurnDCOnOff(0, False)

    rlp.StopHardware()
    return measured_dict


def get_gamma_phase(real: float, imaginary: float) -> Tuple[float, float]:
    """
    Convert real/imaginary impedance values to gamma magnitude and phase.

    Args:
        real: Real part of impedance.
        imaginary: Imaginary part of impedance.

    Returns:
        (gamma_magnitude, phase_degrees)
    """
    gamma = round(math.sqrt(real**2 + imaginary**2), 3)
    phase = round(math.degrees(math.atan2(imaginary, real)), 3)
    return gamma, phase


def output_file(meas: Dict, ismod: bool) -> str:
    """
    Write measured data to load-pull wave format file.

    Args:
        meas: Dictionary containing measured load pull data.
        ismod: If True, write modulated extension (.lpd), else (.lpwave).

    Returns:
        Absolute path to created output file.
    """
    extension = "lpd" if ismod else "lpwave"

    time_str = time.strftime("%Y%m%d-%H%M%S")
    filename = f"Measured_{time_str}.{extension}"
    print(f"Creating output file: {filename}")

    with open(filename, "w") as f:
        first_key = next(iter(meas))
        headers = meas[first_key][0]

        headers_with_data = ["Point"] + headers + ["TargetGamma_1F0"] + ["TargetPhase_1F0[deg]"]
        header_line = "  ".join(headers_with_data)
        f.write(header_line + "\n")

        separator = "!" + "-" * (len(header_line) - 1)
        f.write(separator + "\n")

        point = 0
        for key, value in meas.items():
            target_gamma1_str = str(key.targetGamma_1f0)
            target_phase1_str = str(key.targetPhase_1f0)

            data_lines = value[1:]
            for data_line in data_lines:
                point_str = f"{point:03d}"
                complete_data_line = (
                    [point_str]
                    + [str(val) for val in data_line]
                    + [target_gamma1_str]
                    + [target_phase1_str]
                )
                f.write("  ".join(complete_data_line) + "\n")
                point += 1

    return os.path.abspath(filename)


def write_measured_dict_to_file(measured_dict, filename):
    """
    Write measured dictionary to a text file in transposed format.

    Each header appears once followed by all its values across measurements.

    Args:
        measured_dict: Dictionary containing measurement data.
        filename: Output text file path.
    """
    if not measured_dict:
        return

    frequency_values = []
    psource_values = []
    gain_values = []
    gl_waves_values = []
    phi_waves_values = []

    for _, measurement_data in measured_dict.items():
        headers = measurement_data[0]
        values = measurement_data[1]
        data_map = dict(zip(headers, values))

        frequency_values.append(str(data_map.get("Frequency[GHz]", "")))
        psource_values.append(str(data_map.get("Psource[dBm]", "")))
        gain_values.append(str(data_map.get("GainWavesTrd[dB]", "")))
        gl_waves_values.append(str(data_map.get("|GLWaves@F0|", "")))
        phi_waves_values.append(str(data_map.get("PhiLWaves@F0[deg]", "")))

    with open(filename, "w") as f:
        f.write(f"Frequency : {', '.join(frequency_values)}\n")
        f.write(f"Psource : {', '.join(psource_values)}\n")
        f.write(f"GainWavesTrd[dB] : {', '.join(gain_values)}\n")
        f.write(f"|GLWaves@F0| : {', '.join(gl_waves_values)}\n")
        f.write(f"PhiLWaves@F0[deg] : {', '.join(phi_waves_values)}\n")


# Test runner

def run_test(vals: List) -> Dict:
    """
    Execute the complete load-pull test sequence.

    Args:
        vals: List of test configuration parameters.

    Returns:
        Dictionary containing test results and metadata.
    """
    system_info = {
        "jira": "RVT-633",
        "commit": "2fa5d9b9",
        "hardware": "RapidVT",
        "sn": "1130",
        "fw": "RapidTS-0.1",
    }

    (
        freq,
        pwr_list,
        t1,
        source_cal,
        vector_cal,
        init,
        gs,
        IP,
        mode,
        inp_v,
        out_v,
        aux1_v,
        aux2_v,
        max_gain,
        stop_at_compression,
        power_sweep,
        imp_sweep,
        bias_control,
        if_bw,
        mod_bw,
        mod_waveform,
        pulse_period,
        pulse_width,
        meas_window,
        meas_delay,
        pulse_average,
    ) = vals

    test_passed = True
    start_time = time.time()

    if not IP:
        IP = socket.gethostbyname(socket.gethostname())

    connection_result = rlp.Connect(IP, PORT, KEY)

    if connection_result:
        print("Connected")
        try:
            gold_standard_dict = load_gold_standard_files(gs)
            meas_dict = measure_one(
                gold_standard_dict,
                freq,
                pwr_list,
                t1,
                source_cal,
                vector_cal,
                init,
                mode,
                inp_v,
                out_v,
                aux1_v,
                aux2_v,
                power_sweep,
                imp_sweep,
                max_gain,
                stop_at_compression,
                bias_control,
                if_bw,
                mod_bw,
                mod_waveform,
                pulse_period,
                pulse_width,
                meas_window,
                meas_delay,
                pulse_average,
            )

            output_file_path = None
            ismod = mode == 2

            if len(meas_dict) == 1:
                output_file_path = output_file(meas_dict, ismod)

            execution_time = round(time.time() - start_time)

            # Comparison parameter setup
            if bias_control:
                if ismod:
                    comparison_parameters = [
                        ("PhiLoad@F0[deg]", 1.5),
                        ("|GLoad@F0|", 0.01),
                        ("OutputEff[%]", 0.2),
                        ("GainPwr[dB]", 0.3),
                    ]
                else:
                    comparison_parameters = [
                        ("PhiLWaves@F0[deg]", 1.5),
                        ("|GLWaves@F0|", 0.01),
                        ("OutputEff[%]", 0.2),
                        ("GainWavesPwr[dB]", 0.2),
                    ]
            else:
                if ismod:
                    comparison_parameters = [
                        ("PhiLoad@F0[deg]", 1.5),
                        ("|GLoad@F0|", 0.01),
                        ("GainPwr[dB]", 0.3),
                    ]
                else:
                    comparison_parameters = [
                        ("PhiLWaves@F0[deg]", 1.5),
                        ("|GLWaves@F0|", 0.01),
                        ("GainWavesPwr[dB]", 0.2),
                    ]

            verdict, err_dict = compare(gold_standard_dict, meas_dict, comparison_parameters)

            print(f"Test {verdict}")
            print(f"Execution time: {execution_time}s")

            if verdict == "FAILED":
                test_passed = False
                print(f"Error details: {err_dict}")

            results = {
                "test_passed": test_passed,
                "verdict": verdict,
                "err_dict": err_dict,
                "gold_standard_dict": gold_standard_dict,
                "measured_dict": meas_dict,
                "output_file": output_file_path,
                "system_info": system_info,
                "test_config": {
                    "mode": mode,
                    "mode_name": "CW" if mode == 0 else "Pulsed" if mode == 1 else "Modulated",
                    "source_cal": source_cal,
                    "vector_cal": vector_cal,
                    "init_file": init,
                    "if_bandwidth": f"{if_bw} kHz" if if_bw else "N/A",
                    "mod_bandwidth": f"{mod_bw} kHz" if mod_bw else "N/A",
                    "waveform": mod_waveform if mode == 2 else None,
                    "frequency": freq,
                },
            }

        except Exception as exc:
            print(f"Test execution error: {exc}")
            print(traceback.format_exc())
            test_passed = False

            results = {
                "test_passed": False,
                "verdict": "FAILED",
                "err_dict": {"system_error": str(exc)},
                "gold_standard_dict": {},
                "measured_dict": {},
                "output_file": None,
                "system_info": system_info,
                "test_config": {},
            }

        finally:
            rlp.Disconnect()
            print("Disconnected")

    else:
        print("Connection failed!")
        results = {
            "test_passed": False,
            "verdict": "FAILED",
            "err_dict": {"connection_error": "Failed to connect to measurement system"},
            "gold_standard_dict": {},
            "measured_dict": {},
            "output_file": None,
            "system_info": system_info,
            "test_config": {},
        }

    return results


if __name__ == "__main__":
    pass

