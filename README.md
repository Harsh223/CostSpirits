# CostSpirits - Spacecraft Cost Estimation Tool

A Streamlit-based web application for spacecraft subsystem cost estimation and analysis. This tool helps aerospace engineers and project managers estimate costs for various spacecraft subsystems using historical data and parametric models.

Live Demo: The code is also hosted at https://costspirits.streamlit.app/ thanks to Streamlit Community Cloud.


## Features

- **Easy Launch**: Quick start script (`run_costspirits.py`) for hassle-free application startup with dependency checking

- **Subsystem Cost Estimation**: Generate cost estimates for various spacecraft subsystems including:
  - Structural/Mechanical components
  - Electrical Power & Distribution
  - Command, Control & Data Handling (CC&DH)
  - Propulsion systems
  - Thermal control
  - And many more...

- **Excel Template Generation**: Create customized Excel templates for data collection with subsystem-specific headers

- **Mass Budget Analysis**: Generate mass budget templates and analysis from uploaded data

- **Inflation Adjustment**: Built-in inflation calculations for cost projections using NASA New Start Inflation Index

- **Interactive Data Grid**: User-friendly interface for data entry and editing

- **Professional Export**: Generate styled Excel reports with multiple sheets and formatting

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Harsh223/CostSpirits.git
cd CostSpirits
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:

### Option A: Quick Start (Recommended)
Use the provided quick start script for the easiest launch experience:
```bash
python run_costspirits.py
```
This script will automatically check dependencies and launch the application with helpful status messages.

### Option B: Direct Launch
Alternatively, you can run the application directly:
```bash
streamlit run CostSpirits.py
```

## Usage

1. **Select Subsystems**: Choose the spacecraft subsystems you want to analyze
2. **Generate Templates**: Download Excel templates with appropriate headers for each subsystem
3. **Upload Data**: Upload your completed Excel files with cost and technical data
4. **Analyze Results**: View cost estimates and generate reports

## Troubleshooting

### Application Won't Start?
If you encounter issues launching the application:

1. **Use the Quick Start Script**: Try `python run_costspirits.py` - it will check for common issues
2. **Check Dependencies**: Ensure all requirements are installed with `pip install -r requirements.txt`
3. **Python Version**: Make sure you're using Python 3.8 or higher
4. **Streamlit Issues**: If Streamlit isn't working, try `pip install --upgrade streamlit`

### Common Issues:
- **"streamlit: command not found"**: Use `python -m streamlit run CostSpirits.py` instead
- **Import errors**: Reinstall dependencies with `pip install -r requirements.txt --force-reinstall`
- **File not found**: Make sure you're running the command from the CostSpirits directory

## File Structure

- `CostSpirits.py` - Main Streamlit application
- `run_costspirits.py` - Quick start script for easy application launch
- `subsystem_headers.json` - Configuration file containing headers for each subsystem type
- `Inflation Table.xlsx` - Historical inflation data for cost adjustments
- `requirements.txt` - Python package dependencies
- `README.md` - Project documentation
- `LICENSE` - MIT License file
- `.gitignore` - Git ignore rules

## Subsystem Types Supported

The tool supports cost estimation for the following subsystem categories:

### Structural/Mechanical Group
- Structure
- Mechanisms

### Electrical Power & Distribution Group
- Electrical Power
- Power Distribution/Regulation/Control

### CC&DH Group
- Data Management
- Communication
- Antennas
- Instrumentation Display & Control

### Other Subsystems
- Avionics
- ASE (Aerospace Support Equipment)
- Range Safety
- Separation
- Thermal Control
- Crew Accommodations
- ECLS (Environmental Control and Life Support)
- Launch & Landing Safety
- Miscellaneous
- Attitude Control/GN&C
- Engines
- Propulsion
- Reaction Control
- Solid/Kick Motor
- Thrust Vector Control

Feel free to customize the subsystem_headers file for more headers, for an expanded version of the headers file please contact the author!

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.



