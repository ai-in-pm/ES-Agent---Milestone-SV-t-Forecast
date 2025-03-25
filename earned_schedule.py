import datetime
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import logging

class EarnedScheduleCalculator:
    """Class to calculate Earned Schedule metrics for milestones"""
    
    def __init__(self):
        """Initialize the Earned Schedule calculator"""
        self.setup_logging()
    
    def setup_logging(self):
        """Set up logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            filename='earned_schedule.log',
            filemode='a'
        )
        self.logger = logging.getLogger('EarnedScheduleCalculator')
    
    def calculate_milestone_metrics(self, milestone):
        """Calculate ES metrics for a single milestone"""
        try:
            # Parse dates from strings to datetime objects
            baseline_finish = self._parse_date(milestone.get('baseline_finish'))
            actual_finish = self._parse_date(milestone.get('actual_finish'))
            today = datetime.now()  # Status date (today)
            
            # Skip if no baseline
            if not baseline_finish:
                milestone['sv_t'] = None
                milestone['spi_t'] = None
                milestone['tspi'] = None
                milestone['forecast_finish'] = None
                milestone['status'] = 'No baseline'
                return milestone
            
            # Get planned duration (from project start to baseline finish)
            # For simplicity, we'll use a fixed project start date if not available
            project_start = self._parse_date(milestone.get('baseline_start')) or datetime(2024, 1, 1)
            planned_duration = (baseline_finish - project_start).days
            
            # Calculate actual time (AT) - days from project start until today or completion
            at_days = (today - project_start).days
            
            # Calculate earned schedule (ES) based on completion status
            percent_complete = milestone.get('percent_complete', 0)
            
            if actual_finish:  # Milestone is complete
                es_days = (actual_finish - project_start).days
                milestone['status'] = 'Complete'
            elif percent_complete == 100:  # Complete but no actual finish date
                es_days = at_days
                milestone['status'] = 'Complete'
            else:  # In progress or not started
                # For milestones, ES is typically 0 until complete, but we can interpolate
                # based on % complete of the milestone (assuming linear progress)
                es_days = planned_duration * (percent_complete / 100)
                milestone['status'] = 'In Progress' if percent_complete > 0 else 'Not Started'
            
            # Calculate Schedule Variance in time terms - SV(t)
            sv_t = es_days - at_days  # Positive is ahead of schedule
            
            # Calculate Schedule Performance Index in time terms - SPI(t)
            spi_t = es_days / at_days if at_days > 0 else 1.0
            
            # Calculate To Complete Schedule Performance Index (TSPI)
            remaining_planned = planned_duration - es_days
            remaining_time = (baseline_finish - today).days
            
            if remaining_planned > 0 and remaining_time > 0:
                tspi = remaining_planned / remaining_time
            else:
                tspi = 1.0  # Default if milestone is complete or no time remains
            
            # Calculate Forecast Finish Date using SPI(t)
            if percent_complete < 100 and spi_t > 0:
                # Independent Estimate at Completion (time)
                ieac_t = planned_duration / spi_t
                
                # Calculate how many more days are needed
                remaining_duration = ieac_t - es_days
                
                # Forecast finish date
                forecast_finish = today + timedelta(days=remaining_duration)
                forecast_finish_str = forecast_finish.strftime('%Y-%m-%d %H:%M:%S')
            else:
                forecast_finish_str = actual_finish.strftime('%Y-%m-%d %H:%M:%S') if actual_finish else None
            
            # Update milestone with calculated metrics
            milestone['sv_t'] = round(sv_t, 1)  # days ahead/behind
            milestone['spi_t'] = round(spi_t, 2)
            milestone['tspi'] = round(tspi, 2) if tspi else None
            milestone['forecast_finish'] = forecast_finish_str
            
            # Add risk indicator based on SPI(t) and TSPI
            if percent_complete < 100:
                if spi_t < 0.85 or tspi > 1.2:
                    milestone['risk'] = 'High'
                elif spi_t < 0.95 or tspi > 1.1:
                    milestone['risk'] = 'Medium'
                else:
                    milestone['risk'] = 'Low'
            else:
                milestone['risk'] = 'None'
            
            return milestone
        
        except Exception as e:
            self.logger.error(f"Error calculating metrics for milestone {milestone.get('name')}: {str(e)}")
            milestone['error'] = str(e)
            return milestone
    
    def calculate_forecasts(self, milestones):
        """Calculate forecasts for all milestones"""
        updated_milestones = []
        
        for milestone in milestones:
            updated_milestone = self.calculate_milestone_metrics(milestone)
            updated_milestones.append(updated_milestone)
        
        return updated_milestones
    
    def prepare_dashboard_data(self, milestones):
        """Prepare data for dashboard visualizations"""
        # Create summary stats
        total_milestones = len(milestones)
        completed = sum(1 for m in milestones if m.get('percent_complete') == 100)
        behind_schedule = sum(1 for m in milestones if m.get('sv_t', 0) < 0 and m.get('percent_complete') < 100)
        on_schedule = sum(1 for m in milestones if m.get('sv_t', 0) >= 0 and m.get('percent_complete') < 100)
        
        # Risk breakdown
        high_risk = sum(1 for m in milestones if m.get('risk') == 'High')
        medium_risk = sum(1 for m in milestones if m.get('risk') == 'Medium')
        low_risk = sum(1 for m in milestones if m.get('risk') == 'Low')
        
        # Average SPI(t)
        valid_spi = [m.get('spi_t') for m in milestones if m.get('spi_t') is not None]
        avg_spi_t = sum(valid_spi) / len(valid_spi) if valid_spi else 0
        
        # Timeline data for visualization
        timeline_data = []
        for m in milestones:
            if m.get('baseline_finish') and (m.get('forecast_finish') or m.get('actual_finish')):
                baseline = self._parse_date(m.get('baseline_finish'))
                forecast = self._parse_date(m.get('forecast_finish') or m.get('actual_finish'))
                
                timeline_data.append({
                    'name': m.get('name'),
                    'baseline': m.get('baseline_finish'),
                    'forecast': m.get('forecast_finish') or m.get('actual_finish'),
                    'variance_days': m.get('sv_t'),
                    'status': m.get('status'),
                    'risk': m.get('risk')
                })
        
        # Prepare return data
        dashboard_data = {
            'summary': {
                'total_milestones': total_milestones,
                'completed': completed,
                'behind_schedule': behind_schedule,
                'on_schedule': on_schedule,
                'avg_spi_t': round(avg_spi_t, 2),
                'high_risk': high_risk,
                'medium_risk': medium_risk,
                'low_risk': low_risk
            },
            'timeline': timeline_data,
            'milestones': milestones
        }
        
        return dashboard_data
    
    def _parse_date(self, date_str):
        """Parse date string to datetime object"""
        if not date_str:
            return None
        
        try:
            return datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
        except ValueError:
            try:
                return datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                self.logger.error(f"Could not parse date: {date_str}")
                return None
