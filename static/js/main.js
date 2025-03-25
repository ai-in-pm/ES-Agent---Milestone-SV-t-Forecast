// Main JavaScript for Real-Time Milestone SV(t) Forecasting with MS Project Integration

document.addEventListener('DOMContentLoaded', function() {
    // Elements
    const importBtn = document.getElementById('import-btn');
    const checkConnectionBtn = document.getElementById('check-connection-btn');
    const locateMilestonesBtn = document.getElementById('locate-milestones-btn');
    const importLocatedBtn = document.getElementById('import-located-btn');
    const connectionStatus = document.getElementById('connection-status');
    const lastImportTime = document.getElementById('last-import-time');
    const milestonesTable = document.getElementById('milestones-table');
    const dashboardSummary = document.getElementById('dashboard-summary');
    
    // Bootstrap modal elements
    const statusModal = new bootstrap.Modal(document.getElementById('status-modal'));
    const modalTitle = document.getElementById('modal-title');
    const loadingSpinner = document.getElementById('loading-spinner');
    const loadingMessage = document.getElementById('loading-message');
    const errorMessage = document.getElementById('error-message');
    
    // Milestone locate modal elements
    const milestoneLocateModal = new bootstrap.Modal(document.getElementById('milestone-locate-modal'));
    const milestoneLocateSpinner = document.getElementById('milestone-locate-spinner');
    const milestoneLocateError = document.getElementById('milestone-locate-error');
    const milestoneLocateResults = document.getElementById('milestone-locate-results');
    const milestoneLocateMessage = document.getElementById('milestone-locate-message');
    const locatedMilestonesTable = document.getElementById('located-milestones-table');
    
    // Summary elements
    const totalMilestones = document.getElementById('total-milestones');
    const avgSpi = document.getElementById('avg-spi');
    const behindCount = document.getElementById('behind-count');
    const highRiskCount = document.getElementById('high-risk-count');
    
    // Store located milestones for later use
    let locatedMilestones = [];
    
    // Event listeners
    importBtn.addEventListener('click', importFromMSProject);
    checkConnectionBtn.addEventListener('click', checkMSProjectStatus);
    locateMilestonesBtn.addEventListener('click', locateMilestones);
    importLocatedBtn.addEventListener('click', importLocatedMilestones);
    
    // Initial data load (if any exists from previous session)
    fetchMilestones();
    
    // Functions
    async function checkMSProjectStatus() {
        try {
            // Show modal with loading spinner
            modalTitle.textContent = 'Checking MS Project Status';
            loadingMessage.textContent = 'Checking MS Project connection status...';
            loadingSpinner.style.display = 'block';
            errorMessage.style.display = 'none';
            statusModal.show();
            
            // Call status API
            const response = await fetch('/api/msproject-status');
            const data = await response.json();
            
            // Hide spinner
            loadingSpinner.style.display = 'none';
            
            if (data.status === 'success') {
                // Show success message
                errorMessage.style.display = 'block';
                errorMessage.className = 'alert alert-success';
                
                // Create HTML for open projects
                let projectsHtml = '';
                if (data.open_projects && data.open_projects.length > 0) {
                    projectsHtml = '<h6 class="mt-3">Open Projects:</h6><ul>';
                    data.open_projects.forEach(project => {
                        projectsHtml += `<li><strong>${project.name}</strong> - ${project.tasks} tasks</li>`;
                    });
                    projectsHtml += '</ul>';
                } else {
                    projectsHtml = '<p class="mt-3">No open projects found. Please open a project in MS Project.</p>';
                }
                
                errorMessage.innerHTML = `
                    <div>
                        <p><strong>Status:</strong> ${data.message}</p>
                        ${projectsHtml}
                    </div>
                `;
                
                // Update connection status
                connectionStatus.textContent = data.message;
                connectionStatus.classList.remove('text-danger');
                connectionStatus.classList.add('text-success');
            } else {
                // Show error message
                errorMessage.style.display = 'block';
                errorMessage.className = 'alert alert-danger';
                errorMessage.textContent = data.message || 'Error checking MS Project status';
                
                // Update connection status
                connectionStatus.textContent = 'Not connected to MS Project';
                connectionStatus.classList.remove('text-success');
                connectionStatus.classList.add('text-danger');
            }
        } catch (error) {
            console.error('Status check error:', error);
            
            // Show error in modal
            loadingSpinner.style.display = 'none';
            errorMessage.style.display = 'block';
            errorMessage.className = 'alert alert-danger';
            errorMessage.textContent = error.message || 'Error checking MS Project status';
            
            // Update connection status
            connectionStatus.textContent = 'Not connected to MS Project';
            connectionStatus.classList.remove('text-success');
            connectionStatus.classList.add('text-danger');
        }
    }
    
    async function locateMilestones() {
        try {
            // Show modal with loading spinner
            milestoneLocateSpinner.style.display = 'block';
            milestoneLocateError.style.display = 'none';
            milestoneLocateResults.style.display = 'none';
            milestoneLocateModal.show();
            
            // Call locate API
            const response = await fetch('/api/locate-milestones');
            const data = await response.json();
            
            // Hide spinner
            milestoneLocateSpinner.style.display = 'none';
            
            if (data.status === 'success' && data.milestones && data.milestones.length > 0) {
                // Save the located milestones
                locatedMilestones = data.milestones;
                
                // Show results
                milestoneLocateResults.style.display = 'block';
                milestoneLocateMessage.textContent = data.message || `Found ${data.milestones.length} milestones in MS Project`;
                
                // Render milestones table
                renderLocatedMilestonesTable(data.milestones);
                
                // Update connection status
                connectionStatus.textContent = 'Connected to MS Project';
                connectionStatus.classList.remove('text-danger');
                connectionStatus.classList.add('text-success');
            } else if (data.status === 'success' && (!data.milestones || data.milestones.length === 0)) {
                // Show no milestones message
                milestoneLocateError.style.display = 'block';
                milestoneLocateError.className = 'alert alert-warning';
                milestoneLocateError.textContent = 'No milestones found in the open MS Project file. ' +
                    'Make sure your project has tasks marked as milestones (0 duration tasks or with milestone flag set).';
            } else {
                throw new Error(data.message || 'Failed to locate milestones');
            }
        } catch (error) {
            console.error('Locate milestones error:', error);
            
            // Add more detailed error logging
            console.log('Error details:', {
                name: error.name,
                message: error.message,
                stack: error.stack,
                toString: error.toString()
            });
            
            // Show error in modal
            milestoneLocateSpinner.style.display = 'none';
            milestoneLocateError.style.display = 'block';
            milestoneLocateError.className = 'alert alert-danger';
            milestoneLocateError.textContent = error.message || 'Error locating milestones in MS Project. Please make sure MS Project is open with a project file.';
        }
    }
    
    function renderLocatedMilestonesTable(milestones) {
        if (!milestones || milestones.length === 0) return;
        
        // Clear table
        const tbody = locatedMilestonesTable.querySelector('tbody');
        tbody.innerHTML = '';
        
        // Add rows
        milestones.forEach(milestone => {
            const row = document.createElement('tr');
            
            // Create percent complete display with progress bar
            const percentComplete = milestone.percent_complete || 0;
            const percentCompleteHtml = `
                <div class="progress" style="height: 20px;">
                    <div class="progress-bar ${percentComplete === 100 ? 'bg-success' : 'bg-primary'}" 
                         role="progressbar" 
                         style="width: ${percentComplete}%" 
                         aria-valuenow="${percentComplete}" 
                         aria-valuemin="0" 
                         aria-valuemax="100">${percentComplete}%</div>
                </div>
            `;
            
            // Build the row HTML
            row.innerHTML = `
                <td>${milestone.wbs || '-'}</td>
                <td>${milestone.name || 'Unnamed Milestone'}</td>
                <td>${formatDate(milestone.baseline_finish) || 'Not set'}</td>
                <td>${formatDate(milestone.finish_date) || 'Not scheduled'}</td>
                <td>${milestone.status || 'Not Started'}</td>
                <td>${percentCompleteHtml}</td>
            `;
            
            tbody.appendChild(row);
        });
    }
    
    function importLocatedMilestones() {
        // Close the locate modal
        milestoneLocateModal.hide();
        
        // If we have located milestones, import them
        if (locatedMilestones && locatedMilestones.length > 0) {
            importFromMSProject();
        }
    }
    
    async function importFromMSProject() {
        try {
            // Show modal with loading spinner
            modalTitle.textContent = 'Connecting to MS Project';
            loadingMessage.textContent = 'Connecting to MS Project and extracting milestones...';
            loadingSpinner.style.display = 'block';
            errorMessage.style.display = 'none';
            statusModal.show();
            
            // Call import API
            const response = await fetch('/api/import-from-msproject', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });
            
            const data = await response.json();
            
            if (data.status === 'success') {
                // Update connection status
                connectionStatus.textContent = 'Connected to MS Project';
                connectionStatus.classList.remove('text-danger');
                connectionStatus.classList.add('text-success');
                
                // Update last import time
                const now = new Date();
                lastImportTime.textContent = `Last import: ${now.toLocaleString()}`;
                
                // Close modal
                statusModal.hide();
                
                // Refresh data display
                await fetchMilestones();
                await calculateForecasts();
                await updateDashboard();
            } else if (data.status === 'warning') {
                // Show warning in modal
                loadingSpinner.style.display = 'none';
                errorMessage.style.display = 'block';
                errorMessage.className = 'alert alert-warning';
                errorMessage.textContent = data.message || 'No milestones found in MS Project';
                
                // No need to update connection status as it might be connected but just no milestones
            } else {
                throw new Error(data.message || 'Failed to import from MS Project');
            }
        } catch (error) {
            console.error('Import error:', error);
            
            // Show error in modal
            loadingSpinner.style.display = 'none';
            errorMessage.style.display = 'block';
            errorMessage.className = 'alert alert-danger';
            errorMessage.textContent = error.message || 'Error connecting to MS Project';
            
            // Update connection status
            connectionStatus.textContent = 'Not connected to MS Project';
            connectionStatus.classList.remove('text-success');
            connectionStatus.classList.add('text-danger');
        }
    }
    
    async function fetchMilestones() {
        try {
            const response = await fetch('/api/milestones');
            const milestones = await response.json();
            
            if (milestones && milestones.length > 0) {
                renderMilestonesTable(milestones);
                dashboardSummary.style.display = 'flex';
                return milestones;
            }
        } catch (error) {
            console.error('Error fetching milestones:', error);
        }
        
        return [];
    }
    
    async function calculateForecasts() {
        try {
            const response = await fetch('/api/calculate-forecast', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });
            
            const data = await response.json();
            
            if (data.status === 'success') {
                renderMilestonesTable(data.forecasts);
                return data.forecasts;
            }
        } catch (error) {
            console.error('Error calculating forecasts:', error);
        }
        
        return [];
    }
    
    async function updateDashboard() {
        try {
            const response = await fetch('/api/dashboard-data');
            const data = await response.json();
            
            // Update summary
            if (data.summary) {
                totalMilestones.textContent = data.summary.total_milestones;
                avgSpi.textContent = data.summary.avg_spi_t.toFixed(2);
                behindCount.textContent = data.summary.behind_schedule;
                highRiskCount.textContent = data.summary.high_risk;
            }
            
            // Generate timeline chart
            if (data.timeline && data.timeline.length > 0) {
                renderTimelineChart(data.timeline);
            }
        } catch (error) {
            console.error('Error updating dashboard:', error);
        }
    }
    
    function renderMilestonesTable(milestones) {
        if (!milestones || milestones.length === 0) return;
        
        // Clear table
        const tbody = milestonesTable.querySelector('tbody');
        tbody.innerHTML = '';
        
        // Add rows
        milestones.forEach(milestone => {
            const row = document.createElement('tr');
            
            // Status indicator for name column
            let statusClass = '';
            if (milestone.status === 'Complete') {
                statusClass = 'status-complete';
            } else if (milestone.status === 'In Progress') {
                statusClass = 'status-in-progress';
            } else {
                statusClass = 'status-not-started';
            }
            
            // SV(t) styling
            const svClass = milestone.sv_t < 0 ? 'sv-negative' : 'sv-positive';
            
            // Risk styling
            let riskClass = '';
            if (milestone.risk === 'High') {
                riskClass = 'risk-high';
            } else if (milestone.risk === 'Medium') {
                riskClass = 'risk-medium';
            } else if (milestone.risk === 'Low') {
                riskClass = 'risk-low';
            }
            
            row.innerHTML = `
                <td><span class="status-indicator ${statusClass}"></span> ${milestone.name}</td>
                <td>${formatDate(milestone.baseline_finish)}</td>
                <td>${formatDate(milestone.forecast_finish || milestone.actual_finish)}</td>
                <td>${milestone.status}</td>
                <td class="${svClass}">${milestone.sv_t !== null ? milestone.sv_t : 'N/A'}</td>
                <td>${milestone.spi_t !== null ? milestone.spi_t : 'N/A'}</td>
                <td>${milestone.tspi !== null ? milestone.tspi : 'N/A'}</td>
                <td class="${riskClass}">${milestone.risk || 'N/A'}</td>
            `;
            
            tbody.appendChild(row);
        });
    }
    
    function renderTimelineChart(timelineData) {
        // Sort milestones by baseline date
        timelineData.sort((a, b) => {
            return new Date(a.baseline) - new Date(b.baseline);
        });
        
        const milestoneNames = timelineData.map(m => m.name);
        const baselineDates = timelineData.map(m => m.baseline);
        const forecastDates = timelineData.map(m => m.forecast || m.actual);
        
        // Create Plotly chart
        const plotData = [
            {
                x: baselineDates,
                y: milestoneNames,
                mode: 'markers',
                type: 'scatter',
                name: 'Baseline',
                marker: {
                    color: 'blue',
                    size: 12,
                    symbol: 'diamond'
                }
            },
            {
                x: forecastDates,
                y: milestoneNames,
                mode: 'markers',
                type: 'scatter',
                name: 'Forecast',
                marker: {
                    color: 'red',
                    size: 12,
                    symbol: 'circle'
                }
            }
        ];
        
        const layout = {
            title: 'Milestone Timeline: Baseline vs Forecast',
            xaxis: {
                title: 'Date',
                type: 'date'
            },
            yaxis: {
                title: 'Milestones',
                automargin: true
            },
            height: 400,
            margin: {
                l: 150
            },
            hovermode: 'closest',
            legend: {
                orientation: 'h',
                y: -0.2
            }
        };
        
        Plotly.newPlot('timeline-chart', plotData, layout);
    }
    
    // Helper functions
    function formatDate(dateString) {
        if (!dateString) return 'N/A';
        
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        
        return date.toLocaleDateString();
    }
});
