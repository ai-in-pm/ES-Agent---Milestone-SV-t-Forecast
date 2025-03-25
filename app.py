import os
import json
from flask import Flask, render_template, request, jsonify
from flask_cors import CORS
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from dotenv import load_dotenv
import logging

# Import our custom modules
from msproject_integration import MSProjectIntegration
from earned_schedule import EarnedScheduleCalculator

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='app.log',
    filemode='a'
)
logger = logging.getLogger('app')

# Load environment variables
load_dotenv()

app = Flask(__name__)
CORS(app)  # Enable CORS

# Initialize project integration
project_integration = MSProjectIntegration()

# Initialize earned schedule calculator
earned_schedule_calc = EarnedScheduleCalculator()

# Global variable to store imported milestones
milestones_data = []

@app.route('/')
def index():
    """Main page route"""
    return render_template('index.html')

@app.route('/api/import-from-msproject', methods=['POST'])
def import_from_msproject():
    """API endpoint to import data from MS Project"""
    global milestones_data
    try:
        # Extract milestones from MS Project
        milestones = project_integration.extract_milestones()
        
        # Check if we got any milestones
        if not milestones or len(milestones) == 0:
            logger.warning("No milestones found in the project")
            return jsonify({
                'status': 'warning',
                'message': 'No milestones found in the project. Please ensure your project has tasks marked as milestones.'
            }), 200
        
        # Save to global variable and return success
        milestones_data = milestones
        
        # Calculate Earned Schedule metrics for each milestone
        for milestone in milestones_data:
            earned_schedule_calc.calculate_milestone_metrics(milestone)
        
        return jsonify({
            'status': 'success',
            'message': f'Successfully imported {len(milestones)} milestones',
            'milestones': milestones
        })
    except Exception as e:
        logger.error(f"Error importing milestones: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': f'Error importing milestones: {str(e)}'
        }), 500

@app.route('/api/msproject-status', methods=['GET'])
def get_msproject_status():
    """Check if MS Project is running and what projects are open"""
    try:
        # Try to connect to MS Project
        success, message = project_integration.connect_to_msproject()
        
        # Get list of open projects if connection was successful
        open_projects = []
        if success:
            open_projects = project_integration.get_currently_open_projects()
        
        return jsonify({
            'status': 'success' if success else 'error',
            'message': message,
            'connected': success,
            'open_projects': open_projects
        })
    except Exception as e:
        logger.error(f"Error checking MS Project status: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': f'Error checking MS Project status: {str(e)}',
            'connected': False,
            'open_projects': []
        }), 500

@app.route('/api/locate-milestones', methods=['GET'])
def locate_milestones():
    """Locate milestones in the currently open MS Project file"""
    try:
        # First check if MS Project is running
        success, message = project_integration.connect_to_msproject()
        if not success:
            logger.error(f"Could not connect to MS Project: {message}")
            return jsonify({
                'status': 'error',
                'message': f"Could not connect to MS Project: {message}"
            }), 500
            
        # Extract milestones from MS Project without importing them
        logger.info("Attempting to extract milestones from MS Project")
        milestones = project_integration.extract_milestones()
        
        # Format milestone info for the frontend
        milestone_info = []
        for m in milestones:
            status = "Complete" if m.get('percent_complete', 0) == 100 else "Not Complete"
            milestone_info.append({
                'name': m.get('name', 'Unnamed Milestone'),
                'wbs': m.get('wbs', ''),
                'id': m.get('id', ''),
                'baseline_finish': m.get('baseline_finish', 'No baseline'),
                'finish_date': m.get('finish_date', 'Not scheduled'),
                'status': status,
                'percent_complete': m.get('percent_complete', 0)
            })
        
        logger.info(f"Successfully located {len(milestone_info)} milestones")
        return jsonify({
            'status': 'success',
            'message': f'Found {len(milestone_info)} milestones in the current project',
            'milestones': milestone_info
        })
    except Exception as e:
        error_msg = f"Error locating milestones: {str(e)}"
        logger.error(error_msg)
        logger.exception("Stack trace:")
        return jsonify({
            'status': 'error',
            'message': error_msg
        }), 500

@app.route('/api/milestones', methods=['GET'])
def get_milestones():
    """API endpoint to get all milestones"""
    global milestones_data
    return jsonify(milestones_data)

@app.route('/api/calculate-forecast', methods=['POST'])
def calculate_forecast():
    """Calculate and return forecast for all milestones"""
    global milestones_data
    
    try:
        # Update data with forecasts
        forecasted_milestones = earned_schedule_calc.calculate_forecasts(milestones_data)
        
        # Update global data
        milestones_data = forecasted_milestones
        
        return jsonify({
            'status': 'success',
            'forecasts': forecasted_milestones
        })
    except Exception as e:
        logger.error(f"Error calculating forecasts: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': f'Error calculating forecasts: {str(e)}'
        }), 500

@app.route('/api/dashboard-data', methods=['GET'])
def get_dashboard_data():
    """Get processed data for dashboard visualizations"""
    global milestones_data
    
    if not milestones_data:
        return jsonify({
            'status': 'error',
            'message': 'No milestone data available. Please import from MS Project first.'
        }), 404
    
    # Process data for dashboard
    dashboard_data = earned_schedule_calc.prepare_dashboard_data(milestones_data)
    
    return jsonify(dashboard_data)

if __name__ == '__main__':
    app.run(debug=True)
