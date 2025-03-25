# Real-Time Milestone SV(t) Forecasting with MS Project Integration

This application provides real-time schedule variance forecasting for project milestones by directly integrating with Microsoft Project Desktop. It allows you to import milestone data from your active MS Project file and perform Earned Schedule calculations to forecast completion dates and analyze schedule performance.

The development of this repository was inspired by the Earned Schedule Community in their development of Earned Schedule tools. To learn more visit https://www.earnedschedule.com/Calculator.shtml

![image](https://github.com/user-attachments/assets/cfee8f8f-8028-4143-8cfc-3e6fb67f80fe)



## Features

- Direct integration with MS Project via COM automation
- Automatic extraction of milestone tasks and their attributes
- Real-time calculation of Schedule Variance [SV(t)] and Schedule Performance Index [SPI(t)]
- To-Complete Schedule Performance Index (TSPI) calculation
- Milestone completion forecasting based on current performance
- Interactive timeline visualization of baseline vs. forecast dates
- Risk assessment for each milestone
- Support for both local and server-based Project files

## Prerequisites

- Windows operating system
- Microsoft Project Desktop installed
- Python 3.7+ installed
- Virtual environment (recommended)

## Installation

1. **Activate the virtual environment:**

   ```powershell
   # First, allow script execution (run PowerShell as Administrator)
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   
   # Then activate the virtual environment
   .\venv\Scripts\activate
   ```

2. **Install dependencies:**

   ```
   pip install -r requirements.txt
   ```

## Usage

1. **Start the application:**

   ```
   python app.py
   ```

2. **Access the web interface:**

   Open your browser and navigate to [http://localhost:5000](http://localhost:5000)

3. **Import milestones from MS Project:**

   - Open your project file in Microsoft Project Desktop
   - Click the "Import from MS Project" button in the web interface
   - Review the imported milestones and their forecast metrics

## How It Works

### MS Project Integration

The application connects to MS Project using COM automation (via pywin32) to access the currently open project. It extracts all tasks marked as milestones along with their relevant attributes:

- Name and ID
- Start and Finish dates
- Baseline dates
- Percent Complete
- Actual Start and Finish dates

### Earned Schedule Calculations

Using the Earned Schedule methodology for schedule variance analysis:

- SV(t): Schedule Variance in time units (days). Positive values indicate ahead of schedule, negative values indicate behind schedule.
- SPI(t): Schedule Performance Index in time. Values < 1 indicate behind schedule, values > 1 indicate ahead of schedule.
- TSPI: To-Complete Schedule Performance Index, the efficiency needed to meet the baseline date.
- Forecast Finish: Projected completion date based on current performance.

## Security and Privacy

- All data processing occurs locally on your machine
- No project data is sent to external servers
- Integration requires explicit user permission
- MS Project files are accessed in read-only mode

## Troubleshooting

- If you get a COM connection error, ensure MS Project is open with a project file loaded
- If PowerShell script execution is blocked, run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser` in an elevated PowerShell prompt

## License

This project is proprietary and confidential.
