import win32com.client
import pythoncom
import datetime
import logging
import json
import os
import sys
import time
from win32com.client import constants

class MSProjectIntegration:
    """Class to handle integration with MS Project via COM"""
    
    def __init__(self):
        """Initialize the MS Project integration"""
        self.app = None
        self.project = None
        self.setup_logging()
    
    def setup_logging(self):
        """Set up logging for the integration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            filename='msproject_integration.log',
            filemode='a'
        )
        self.logger = logging.getLogger('MSProjectIntegration')
    
    def connect_to_msproject(self):
        """Connect to MS Project application via COM"""
        try:
            # Initialize COM for this thread (important for multi-threaded apps)
            pythoncom.CoInitialize()
            
            # Try to get active instance first with increased timeout and error handling
            self.logger.info("Attempting to connect to MS Project...")
            try:
                # Try to connect to existing MS Project instance
                self.app = win32com.client.GetActiveObject("MSProject.Application")
                self.logger.info("Connected to existing MS Project instance")
            except Exception as e:
                self.logger.warning(f"No active MS Project instance found: {str(e)}")
                # Try different approach - create new instance
                try:
                    # Create new MS Project instance without the problematic parameter
                    self.app = win32com.client.Dispatch("MSProject.Application")
                    self.app.Visible = True
                    self.logger.info("Created new MS Project instance")
                    # Give MS Project a moment to initialize
                    time.sleep(2)
                except Exception as e2:
                    self.logger.error(f"Failed to create MS Project instance: {str(e2)}")
                    return False, f"MS Project could not be started. Please ensure it's installed correctly. Error: {str(e2)}"
            
            # Check if there's an active project
            if not self.app.ActiveProject:
                # Try to check if there are any open projects
                try:
                    # Try to get all open projects in Project
                    project_count = self.app.Projects.Count
                    if project_count > 0:
                        self.project = self.app.Projects(1)  # Get the first project
                        self.logger.info(f"Selected project: {self.project.Name}")
                    else:
                        # No projects open, suggest opening one
                        return False, "No projects are open in MS Project. Please open a project file and try again."
                except Exception as e:
                    self.logger.error(f"Error checking for open projects: {str(e)}")
                    return False, "No active project found. Please open a project in MS Project and try again."
            else:
                self.project = self.app.ActiveProject
                project_name = self.project.Name
                self.logger.info(f"Connected to active project: {project_name}")
            
            return True, f"Connected to MS Project: {self.project.Name}"
        
        except Exception as e:
            error_message = f"Error connecting to MS Project: {str(e)}"
            self.logger.error(error_message)
            return False, error_message
        
    def extract_milestones(self):
        """Extract milestone tasks from the active project"""
        milestones = []
        
        # First, connect to MS Project
        success, message = self.connect_to_msproject()
        if not success:
            raise Exception(message)
        
        try:
            # Get all tasks from the active project
            tasks = self.project.Tasks
            task_count = tasks.Count
            
            self.logger.info(f"Found {task_count} tasks in the project")
            
            # Use multiple methods to identify milestones
            for i in range(1, task_count + 1):
                task = tasks(i)  # 1-based indexing in COM
                
                # Skip summary tasks if needed
                if hasattr(task, 'Summary') and task.Summary and not task.Milestone:
                    continue
                
                # Multiple checks for milestones
                is_milestone = False
                
                # Method 1: Check explicit Milestone flag
                if hasattr(task, 'Milestone') and task.Milestone:
                    is_milestone = True
                    self.logger.info(f"Found milestone by flag: {task.Name}")
                
                # Method 2: Check if duration is zero (common milestone indicator)
                elif hasattr(task, 'Duration') and task.Duration == 0:
                    is_milestone = True
                    self.logger.info(f"Found milestone by zero duration: {task.Name}")
                
                # Method 3: Check for milestone in name (optional)
                elif 'milestone' in task.Name.lower():
                    is_milestone = True
                    self.logger.info(f"Found milestone by name: {task.Name}")
                
                if is_milestone:
                    # Extract all relevant fields
                    milestone_data = self._extract_task_data(task)
                    milestones.append(milestone_data)
            
            self.logger.info(f"Extracted {len(milestones)} milestones")
            
            # If no milestones found, create a helpful message
            if len(milestones) == 0:
                self.logger.warning("No milestones found in the project")
                # Get project stats to help debugging
                try:
                    total_tasks = task_count
                    completed_tasks = sum(1 for i in range(1, task_count + 1) if tasks(i).PercentComplete == 100)
                    self.logger.info(f"Project contains {total_tasks} total tasks, {completed_tasks} completed tasks")
                except:
                    pass
            
            # Save to file as a backup
            self._save_milestones_to_file(milestones)
            
            return milestones
        
        except Exception as e:
            error_message = f"Error extracting milestones: {str(e)}"
            self.logger.error(error_message)
            raise Exception(error_message)
        
        finally:
            # Clean up COM resources
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _extract_task_data(self, task):
        """Extract relevant data fields from a task"""
        milestone_data = {
            'id': task.UniqueID,
            'wbs': self._safe_get_property(task, 'WBS', ''),
            'name': task.Name,
            'percent_complete': self._safe_get_property(task, 'PercentComplete', 0),
            
            # Handle dates - convert to string representations
            'start_date': self._format_date(self._safe_get_property(task, 'Start', None)),
            'finish_date': self._format_date(self._safe_get_property(task, 'Finish', None)),
            
            # Baseline dates
            'baseline_start': self._format_date(self._safe_get_property(task, 'BaselineStart', None)),
            'baseline_finish': self._format_date(self._safe_get_property(task, 'BaselineFinish', None)),
            
            # Actual dates
            'actual_start': self._format_date(self._safe_get_property(task, 'ActualStart', None)),
            'actual_finish': self._format_date(self._safe_get_property(task, 'ActualFinish', None)),
            
            # Notes
            'notes': self._safe_get_property(task, 'Notes', '')
        }
        
        return milestone_data
    
    def _safe_get_property(self, obj, property_name, default_value):
        """Safely get a property or return default value if not available"""
        try:
            return getattr(obj, property_name)
        except Exception as e:
            self.logger.warning(f"Could not get property {property_name}: {str(e)}")
            return default_value
    
    def _format_date(self, date_value):
        """Format a date value from MS Project"""
        if date_value is None:
            return None
            
        try:
            # Handle both datetime objects and MS Project date objects
            if isinstance(date_value, datetime.datetime):
                return date_value.strftime('%Y-%m-%d %H:%M:%S')
            elif hasattr(date_value, 'Format'):
                try:
                    # Convert to Python datetime if possible
                    py_date = datetime.datetime(
                        date_value.Year, 
                        date_value.Month, 
                        date_value.Day, 
                        date_value.Hour if hasattr(date_value, 'Hour') else 0, 
                        date_value.Minute if hasattr(date_value, 'Minute') else 0, 
                        date_value.Second if hasattr(date_value, 'Second') else 0
                    )
                    return py_date.strftime('%Y-%m-%d %H:%M:%S')
                except Exception as e:
                    self.logger.warning(f"Error converting date: {str(e)}")
                    # Fall back to string representation
                    return str(date_value)
            else:
                # Last resort - return string representation
                return str(date_value)
        except Exception as e:
            self.logger.warning(f"Error formatting date: {str(e)}")
            return str(date_value) if date_value else None
    
    def _save_milestones_to_file(self, milestones):
        """Save milestone data to a backup JSON file"""
        try:
            with open('milestones_backup.json', 'w') as f:
                json.dump(milestones, f, indent=2)
            
            self.logger.info("Saved milestones backup to milestones_backup.json")
        except Exception as e:
            self.logger.error(f"Error saving milestones backup: {str(e)}")

    def disconnect(self):
        """Disconnect from MS Project"""
        try:
            # Release COM objects
            self.project = None
            self.app = None
            pythoncom.CoUninitialize()
            self.logger.info("Disconnected from MS Project")
        except Exception as e:
            self.logger.error(f"Error disconnecting from MS Project: {str(e)}")
            
    def get_currently_open_projects(self):
        """Get a list of all currently open projects in MS Project"""
        projects_list = []
        
        success, message = self.connect_to_msproject()
        if not success:
            return projects_list
            
        try:
            # Try to enumerate all open projects
            for i in range(1, self.app.Projects.Count + 1):
                project = self.app.Projects(i)
                projects_list.append({
                    'name': project.Name,
                    'path': project.FullName if hasattr(project, 'FullName') else 'Unknown',
                    'tasks': project.Tasks.Count if hasattr(project, 'Tasks') else 0
                })
                
            return projects_list
        except Exception as e:
            self.logger.error(f"Error getting open projects: {str(e)}")
            return projects_list
        finally:
            # Clean up COM resources
            try:
                pythoncom.CoUninitialize()
            except:
                pass
