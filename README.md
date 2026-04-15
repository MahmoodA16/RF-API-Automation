# RF API Automation

Automated test controller for running S-Parameter and Load Pull measurements on RF/microwave systems. You define your test matrix in a CSV, point the tool at it, and it runs everything through the RapidMS API, then gives you a web report with pass/fail results.

I built this to replace the manual process of configuring and running tests one at a time through the vendor GUI. A full regression that used to take the better part of a day now runs unattended.

## How it works

The application has two layers: a Tkinter desktop GUI for loading CSV test configurations and controlling execution, and a Flask web server that generates HTML test reports you can view in a browser.

**Test flow:**

1. Load a CSV file that defines your test matrix (device serial numbers, test types, modes, calibration files, frequencies, etc.)
1. Select which test types (S-Parameter, F0 Load Pull) and modes (CW, Pulsed, Modulated) to run
1. Hit Start. Tests execute in a background thread so the GUI stays responsive.
1. Results accumulate in the console log. When done, open the web report for a summary page with drill-down into individual test results.

The system talks to the measurement hardware through `RapidMS.ClientAPI.dll` (a .NET DLL accessed via Python.NET). The YAML config (`API_Meas.yaml`) defines the API connection and measurement parameters.

## Test types

**S-Parameter** (`Sparameter_Automation.py`): Runs S-parameter measurements and compares results against a gold standard `.s2p` reference file. Pass/fail is based on magnitude and phase tolerances at each frequency point.

**Load Pull** (`Load_Pull.py`): Runs load pull measurements in CW, Pulsed, or Modulated modes. Configures source/vector calibrations, sets impedance tuning, runs the measurement, and compares output against expected values. Supports different IF bandwidths for CW/Pulsed vs modulated bandwidth for MOD mode.

## Files

|File                      |What it does                                                                                                                                  |
|--------------------------|----------------------------------------------------------------------------------------------------------------------------------------------|
|`main.py`                 |The main application. Tkinter GUI, CSV parsing, test orchestration, Flask web server for reports. ~1200 lines.                                |
|`Sparameter_Automation.py`|S-parameter test execution. Loads gold standard, runs measurement via RapidMS API, compares results within tolerances.                        |
|`Load_Pull.py`            |Load pull test execution. Handles CW/Pulsed/Modulated modes, calibration loading, impedance configuration, measurement, and result validation.|
|`ClassFile.py`            |Data model for S2P (Touchstone) file parsing. Holds S-parameter data indexed by frequency.                                                    |
|`API_Meas.yaml`           |API configuration and measurement parameter definitions.                                                                                      |
|`Template.csv`            |Example CSV test matrix showing the expected column format.                                                                                   |
|`templates/`              |Flask HTML templates for the web report (summary page, S-param detail report, load pull detail report).                                       |
|`RapidMS.ClientAPI.dll`   |Vendor .NET API for controlling the measurement hardware. Accessed via pythonnet.                                                             |
|`requirements.txt`        |Python dependencies.                                                                                                                          |

## Requirements

- Python 3.x
- pythonnet (`clr`) for .NET interop with the RapidMS API
- pandas, numpy, flask
- The RapidMS measurement software running and accessible on the configured port
- Actual RF measurement hardware connected (this won’t do anything useful without it)

## CSV format

Use `Template.csv` as a starting point. Each row defines a test, with columns for the system name, serial number, test type, mode, calibration file paths, frequency list, tolerances, and so on. The `main.py` CSV parser (`extract_information`) reads these into a dictionary keyed by test name and type.

## Web report

After running tests, click “View Web Report” in the GUI. Flask spins up on a random free port and opens your browser. The summary page shows pass/fail for each test with links to detailed reports. S-parameter reports show per-frequency error breakdowns. Load pull reports show measured vs expected values.

The report also links to data files. Clicking a data file link opens it in Load Pull Explorer (if installed) via a PowerShell subprocess.

## Note

This was built for a specific hardware setup and vendor API. The `RapidMS.ClientAPI.dll` and calibration file paths are environment-specific. You’ll need the matching hardware and software to actually run measurements, but the test orchestration, CSV parsing, and reporting code is general enough to adapt.
