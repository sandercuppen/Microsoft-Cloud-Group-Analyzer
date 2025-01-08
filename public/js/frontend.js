class GroupAnalyzer {
    constructor() {
        this.authStatus = document.getElementById('authStatus');
        this.startAnalysisBtn = document.getElementById('startAnalysis');
        this.exportCSVBtn = document.getElementById('exportCSV');
        this.exportJSONBtn = document.getElementById('exportJSON');
        this.resultsContainer = document.getElementById('results');
        this.groupIdInput = document.getElementById('groupId');
        
        this.bindEvents();
        this.checkAuth();
    }

    bindEvents() {
        this.startAnalysisBtn.addEventListener('click', () => this.startAnalysis());
        this.exportCSVBtn.addEventListener('click', () => this.exportData('csv'));
        this.exportJSONBtn.addEventListener('click', () => this.exportData('json'));
    }

    async checkAuth() {
        try {
            const response = await fetch('/api/auth-status');
            const data = await response.json();
            
            if (data.authenticated) {
                this.authStatus.className = 'alert alert-success';
                this.authStatus.textContent = `Authenticated as ${data.userPrincipalName}`;
                this.startAnalysisBtn.disabled = false;
            } else {
                this.authStatus.className = 'alert alert-warning';
                this.authStatus.textContent = 'Not authenticated. Please check your credentials.';
            }
        } catch (error) {
            this.authStatus.className = 'alert alert-danger';
            this.authStatus.textContent = 'Error checking authentication status';
        }
    }

    async startAnalysis() {
        this.startAnalysisBtn.disabled = true;
        this.resultsContainer.innerHTML = '<div class="text-center"><div class="spinner-border text-primary" role="status"></div><div class="mt-2">Analyzing groups...</div></div>';

        try {
            const groupId = this.groupIdInput?.value?.trim() || 'all';
            const response = await fetch(`/api/analyze?groupId=${encodeURIComponent(groupId)}`);
            const data = await response.json();
            
            this.displayResults(data);
            this.exportCSVBtn.disabled = false;
            this.exportJSONBtn.disabled = false;
        } catch (error) {
            this.resultsContainer.innerHTML = '<div class="alert alert-danger">Error during analysis. Please try again.</div>';
        } finally {
            this.startAnalysisBtn.disabled = false;
        }
    }

    displayResults(data) {
        this.resultsContainer.innerHTML = '';
        
        if (data.length === 0) {
            this.resultsContainer.innerHTML = '<div class="alert alert-info">No groups found.</div>';
            return;
        }
        
        data.forEach(group => {
            const groupElement = document.createElement('div');
            groupElement.className = 'group-item';
            
            groupElement.innerHTML = `
                <h6>${group.displayName}</h6>
                <div><strong>Type:</strong> ${group.groupTypes.join(', ') || 'Security'}</div>
                <div><strong>Members:</strong> ${group.memberCount || 0}</div>
                <div><strong>Owners:</strong> ${group.ownerCount || 0}</div>
                ${group.issues?.length ? `
                    <div class="mt-2">
                        <strong>Issues:</strong>
                        <ul class="mb-0">
                            ${group.issues.map(issue => `<li class="status-warning">${issue}</li>`).join('')}
                        </ul>
                    </div>
                ` : ''}
            `;
            
            this.resultsContainer.appendChild(groupElement);
        });
    }

    async exportData(format) {
        try {
            const groupId = this.groupIdInput?.value?.trim() || 'all';
            const response = await fetch(`/api/export?format=${format}&groupId=${encodeURIComponent(groupId)}`);
            const blob = await response.blob();
            
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `group-analysis.${format}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        } catch (error) {
            alert(`Error exporting data as ${format.toUpperCase()}`);
        }
    }
}

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    new GroupAnalyzer();
});
